/** Stream plumbing the tests need on both sides of the writer. */

/** Everything a `ReadableStream` produces, as one buffer. */
export async function collect(stream: ReadableStream<Uint8Array>): Promise<Buffer> {
    const chunks: Buffer[] = [];
    const reader = stream.getReader();
    for (;;) {
        const result = await reader.read();
        if (result.done) return Buffer.concat(chunks);
        chunks.push(Buffer.from(result.value));
    }
}

/** The values of `items`, reachable only through `for await`. */
export async function* asAsyncIterable<T>(items: readonly T[]): AsyncGenerator<T> {
    for (const item of items) yield item;
}

/** The values of `items`, out of a generator rather than an array. */
export function* asIterable<T>(items: readonly T[]): Generator<T> {
    for (const item of items) yield item;
}

/** `items` as a `ReadableStream`, to feed the writable side of a transform. */
export function asReadable<T>(items: readonly T[]): ReadableStream<T> {
    let index = 0;
    return new ReadableStream<T>({
        pull(controller) {
            if (index < items.length) controller.enqueue(items[index++] as T);
            else controller.close();
        },
    });
}

/** A sink that keeps every chunk it is handed, and how many it has seen. */
export function recordingSink(): {
    sink: (bytes: Uint8Array) => void;
    chunks: Uint8Array[];
    bytes: () => Buffer;
} {
    const chunks: Uint8Array[] = [];
    return {
        sink: (bytes) => {
            // `fflate` reuses its output buffer, so the bytes have to be
            // copied out of it before the next chunk overwrites them.
            chunks.push(Uint8Array.prototype.slice.call(bytes));
        },
        chunks,
        bytes: () => Buffer.concat(chunks.map((chunk) => Buffer.from(chunk))),
    };
}
