// The one place the reader talks to an XML parser.
//
// Writing XML is a matter of putting the right characters in the right order,
// which is why `cell.ts` does it by hand. Reading it is not the same job:
// entities, numeric references, attribute quoting, text split across chunks
// and everything a producer is allowed to write that nobody expects. That is
// a solved problem and `saxes` solves it, so the whole of this module is the
// shape the rest of the reader wants it in.
//
// Pull-based on the outside, push-based on the inside: `saxes` calls back as
// it parses, and the reader wants to `for await` over rows. Nothing clever
// bridges the two — whoever feeds a chunk in drains whatever came out of it
// right after, which is the same order with the results in hand instead of on
// the stack.
import { SaxesParser } from 'saxes';

/** What a part's parser is told about, as it goes. */
export interface XmlHandlers {
    /** An element started. `attributes` is empty rather than absent. */
    open?(name: string, attributes: Record<string, string>): void;
    /**
     * Text inside the current element. It can arrive in pieces — a chunk
     * boundary is enough to split it — so a handler that wants it whole
     * accumulates until the element closes.
     */
    text?(text: string): void;
    close?(name: string): void;
}

/**
 * An element name with its namespace prefix dropped: `x:worksheet` is the
 * `worksheet` of a producer that writes prefixes, and the parts are read by
 * the names the spec gives them either way.
 *
 * The prefix is dropped rather than resolved because there is nothing to
 * resolve it against that matters here: an xlsx part has one vocabulary in
 * it, and a `<c>` inside a `<row>` inside a `<sheetData>` is a cell whatever
 * the producer decided to call the namespace. Reading only unprefixed names
 * would come back as a workbook with no rows in it, which is the kind of
 * quiet wrong answer worth going out of the way to avoid.
 */
function localName(name: string): string {
    const colon = name.indexOf(':');
    return colon === -1 ? name : name.slice(colon + 1);
}

/**
 * A part being parsed, fed with `write` and finished with `close`.
 *
 * `saxes` reports a malformed document through an `error` event and carries
 * on afterwards; here the first one is kept and thrown at the caller, on the
 * `write` that caused it. A part that does not parse is a file that cannot be
 * read, not something to return half of.
 */
export class XmlParser {
    private readonly parser: SaxesParser;
    private failure: Error | undefined;

    constructor(handlers: XmlHandlers, partName: string) {
        this.parser = new SaxesParser({ fileName: partName });
        this.parser.on('error', (err) => {
            this.failure ??= err;
        });
        if (handlers.open) {
            const open = handlers.open.bind(handlers);
            this.parser.on('opentag', (tag) => {
                open(localName(tag.name), tag.attributes as Record<string, string>);
            });
        }
        if (handlers.text) this.parser.on('text', handlers.text.bind(handlers));
        if (handlers.close) {
            const close = handlers.close.bind(handlers);
            this.parser.on('closetag', (tag) => close(localName(tag.name)));
        }
    }

    private checkFailure(): void {
        if (this.failure) throw this.failure;
    }

    write(chunk: string): void {
        this.parser.write(chunk);
        this.checkFailure();
    }

    /** Finishes the document, and with it the well-formedness checks. */
    close(): void {
        this.parser.close();
        this.checkFailure();
    }
}

/** A part small enough to be read in one piece. */
export function parseXml(text: string, handlers: XmlHandlers, partName: string): void {
    const parser = new XmlParser(handlers, partName);
    parser.write(text);
    parser.close();
}

/**
 * Bytes as text, in the same pieces they arrive in.
 *
 * `TextDecoder` in streaming mode is what makes this safe to do chunk by
 * chunk: a character whose bytes fall on both sides of a boundary is held
 * until the rest of it turns up, rather than coming out as a replacement
 * character in the middle of a cell.
 */
export async function* decodeChunks(chunks: AsyncIterable<Uint8Array>): AsyncIterable<string> {
    const decoder = new TextDecoder();
    for await (const chunk of chunks) {
        const text = decoder.decode(chunk, { stream: true });
        if (text) yield text;
    }
    const rest = decoder.decode();
    if (rest) yield rest;
}
