// The little bit of DOM `src/browser/index.ts` reaches for, faked in Node.
//
// This is not a stand-in for running the code in a real browser — that is the
// second stage, and `npm run example:browser:test` already drives the example
// page through Chromium. What it covers is the branch the browser itself
// decides: whether `showSaveFilePicker` is there, and what the fallback does
// with the `Blob` when it is not.

export interface FakeAnchor {
    href: string;
    download: string;
    clicks: number;
}

export interface DomStub {
    /** Anchors `document.createElement('a')` handed out. */
    anchors: FakeAnchor[];
    /** Object URLs created, and the blobs they were created from. */
    objectUrls: { url: string; blob: Blob }[];
    revokedUrls: string[];
    /** What the File System Access route was asked to save, once it is done. */
    saved: { suggestedName: string | undefined; bytes: Buffer } | undefined;
    restore(): void;
}

/**
 * Installs `window`, `document` and `URL.createObjectURL` for the duration of
 * a test. With `showSaveFilePicker: false` the picker is absent, which is what
 * Firefox and Safari look like from the inside.
 */
export function stubDom(options: { showSaveFilePicker: boolean }): DomStub {
    const stub: DomStub = {
        anchors: [],
        objectUrls: [],
        revokedUrls: [],
        saved: undefined,
        restore,
    };

    const picker = async (pickerOptions?: { suggestedName?: string }) => ({
        createWritable: async () => {
            const chunks: Buffer[] = [];
            return new WritableStream<Uint8Array>({
                write(chunk) {
                    chunks.push(Buffer.from(chunk));
                },
                close() {
                    stub.saved = {
                        suggestedName: pickerOptions?.suggestedName,
                        bytes: Buffer.concat(chunks),
                    };
                },
            });
        },
    });

    const window = options.showSaveFilePicker ? { showSaveFilePicker: picker } : {};

    const document = {
        createElement(tagName: string) {
            if (tagName !== 'a') throw new Error(`unexpected <${tagName}>`);
            const anchor: FakeAnchor = { href: '', download: '', clicks: 0 };
            stub.anchors.push(anchor);
            return Object.assign(anchor, {
                click() {
                    anchor.clicks++;
                },
            });
        },
    };

    const previousCreate = Reflect.get(URL, 'createObjectURL');
    const previousRevoke = Reflect.get(URL, 'revokeObjectURL');

    Reflect.set(globalThis, 'window', window);
    Reflect.set(globalThis, 'document', document);
    Reflect.set(URL, 'createObjectURL', (blob: Blob) => {
        const url = `blob:xlsx-now/${stub.objectUrls.length}`;
        stub.objectUrls.push({ url, blob });
        return url;
    });
    Reflect.set(URL, 'revokeObjectURL', (url: string) => {
        stub.revokedUrls.push(url);
    });

    function restore(): void {
        Reflect.deleteProperty(globalThis, 'window');
        Reflect.deleteProperty(globalThis, 'document');
        if (previousCreate) Reflect.set(URL, 'createObjectURL', previousCreate);
        else Reflect.deleteProperty(URL, 'createObjectURL');
        if (previousRevoke) Reflect.set(URL, 'revokeObjectURL', previousRevoke);
        else Reflect.deleteProperty(URL, 'revokeObjectURL');
    }

    return stub;
}
