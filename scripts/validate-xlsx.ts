// Independent validation of a generated .xlsx.
//
// The point is to read the file back with implementations that had nothing to
// do with writing it: `yauzl` (plus Node's own `zlib`) for the container and
// `exceljs` for the workbook. Validating with `fflate` would only prove that
// the writer agrees with itself.
//
// Checks both the spreadsheet content (styles, values, row count) and the
// ZIP-level properties an OOXML consumer expects — compression, ZIP version
// 2.0, no ZIP64, and sizes that were not known when the entry was written.
//
// Usage: node dist/scripts/validate-xlsx.js out/example-node.xlsx [expected_rows]
import { readFile } from 'node:fs/promises';
import { crc32 } from 'node:zlib';
import ExcelJS from 'exceljs';
import { openPromise, type Entry, type ZipFile } from 'yauzl';

/** "Version needed to extract" value that means ZIP64. */
const ZIP64_VERSION = 45;
/** ZIP 2.0, the highest version Office accepts. */
const MAX_VERSION = 20;
/** General-purpose flag bit 3: sizes and CRC follow the data, not precede it. */
const FLAG_DATA_DESCRIPTOR = 0x0008;
const METHOD_DEFLATE = 8;

const SHEET_PATH = 'xl/worksheets/sheet1.xml';
const PK_FILL_ARGB = 'FFFFE699';

function check(condition: unknown, message: string): asserts condition {
    if (!condition) throw new Error(message);
}

async function readEntry(zip: ZipFile, entry: Entry): Promise<Buffer> {
    const stream = await zip.openReadStreamPromise(entry);
    const chunks: Buffer[] = [];
    for await (const chunk of stream) chunks.push(chunk as Buffer);
    return Buffer.concat(chunks);
}

interface ContainerReport {
    names: string[];
    sheetRawBytes: number;
    sheetStoredBytes: number;
}

/** Checks the ZIP container itself, below the spreadsheet level. */
async function checkContainer(path: string): Promise<ContainerReport> {
    const zip = await openPromise(path, { lazyEntries: true, validateEntrySizes: true });
    const names: string[] = [];
    let sheet: Entry | undefined;

    try {
        for await (const entry of zip.eachEntry()) {
            names.push(entry.fileName);

            // The low byte is the ZIP version; the high byte is the source OS.
            const madeBy = entry.versionMadeBy & 0xff;
            check(
                madeBy <= MAX_VERSION,
                `${entry.fileName}: created with ZIP ${madeBy / 10}, Office requires 2.0`,
            );
            check(
                entry.versionNeededToExtract <= MAX_VERSION,
                `${entry.fileName}: needs ZIP ${entry.versionNeededToExtract / 10} to extract, Office requires 2.0`,
            );
            check(entry.versionNeededToExtract < ZIP64_VERSION, `${entry.fileName}: ZIP64 entry`);

            // Reading the entry inflates it with Node's zlib and lets yauzl
            // verify the declared sizes; the CRC it does not check, so compare
            // it here against the one the archive carries.
            const data = await readEntry(zip, entry);
            check(
                crc32(data) === entry.crc32,
                `${entry.fileName}: CRC mismatch against the archive's own value`,
            );

            const header = await zip.readLocalFileHeaderPromise(entry);
            check(
                (header.generalPurposeBitFlag & FLAG_DATA_DESCRIPTOR) !== 0,
                `${entry.fileName}: sizes were known upfront (not streamed)`,
            );
            check(
                header.crc32 === 0 && header.compressedSize === 0 && header.uncompressedSize === 0,
                `${entry.fileName}: local header carries sizes despite bit 3`,
            );

            if (entry.fileName === SHEET_PATH) sheet = entry;
        }
    } finally {
        zip.close();
    }

    check(sheet, `missing ${SHEET_PATH}`);
    check(sheet.compressionMethod === METHOD_DEFLATE, 'worksheet is not deflated');

    // The trailing records are, at most, [ZIP64 EOCD][ZIP64 locator][EOCD];
    // looking only at the tail avoids matching these signatures by chance
    // inside compressed data.
    const raw = await readFile(path);
    const tail = raw.subarray(Math.max(0, raw.length - 128));
    check(!tail.includes(Buffer.from('PK\x06\x06')), 'ZIP64 end of central directory present');
    check(!tail.includes(Buffer.from('PK\x06\x07')), 'ZIP64 end of central directory locator present');

    return {
        names,
        sheetRawBytes: sheet.uncompressedSize,
        sheetStoredBytes: sheet.compressedSize,
    };
}

interface WorkbookReport {
    /** Every sheet the workbook declares, in order — the first one is checked. */
    sheetNames: string[];
    rows: number;
    header: string[];
    lastName: unknown;
    frozen: string;
}

function isPkFilled(cell: ExcelJS.Cell): boolean {
    return cell.fill?.type === 'pattern' && cell.fill.fgColor?.argb === PK_FILL_ARGB;
}

/**
 * The header row and the leading pk columns should be frozen, so they stay on
 * screen while the sheet scrolls — and no column at all when the pks are not
 * the first columns, or when every column is one.
 *
 * Which columns are pk is read back from the fills in the file, not taken
 * from the writer's own configuration.
 */
function checkFrozenPanes(sheet: ExcelJS.Worksheet): string {
    const headerRow = sheet.getRow(1);
    let leadingPks = 0;
    while (leadingPks < sheet.columnCount && isPkFilled(headerRow.getCell(leadingPks + 1))) {
        leadingPks++;
    }
    const expectedX = leadingPks === sheet.columnCount ? 0 : leadingPks;

    const view = sheet.views?.[0];
    check(view?.state === 'frozen', `the sheet view is ${view?.state ?? 'missing'}, not frozen`);
    check(view.ySplit === 1, `frozen rows: expected 1, found ${view.ySplit}`);
    check(
        (view.xSplit ?? 0) === expectedX,
        `frozen columns: expected ${expectedX} (leading pk columns), found ${view.xSplit ?? 0}`,
    );
    return `1 row + ${expectedX} column${expectedX === 1 ? '' : 's'}`;
}

async function checkWorkbook(path: string, expectedRows: number | undefined): Promise<WorkbookReport> {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(path);
    const sheet = workbook.worksheets[0];
    check(sheet, 'the workbook has no worksheets');
    const sheetNames = workbook.worksheets.map((each) => each.name);
    check(
        new Set(sheetNames.map((name) => name.toLowerCase())).size === sheetNames.length,
        `two sheets share a name: ${sheetNames.join(', ')}`,
    );

    const headerRow = sheet.getRow(1);
    const header: string[] = [];
    for (let i = 1; i <= sheet.columnCount; i++) {
        const cell = headerRow.getCell(i);
        header.push(String(cell.value));
        check(cell.font?.bold, `header cell ${cell.address} is not bold`);
    }

    const pkHeader = headerRow.getCell(1);
    check(isPkFilled(pkHeader), `PK header fill is ${JSON.stringify(pkHeader.fill)}`);

    const firstData = sheet.getRow(2);
    check(isPkFilled(firstData.getCell(1)), 'PK cell not filled');

    const plainCell = firstData.getCell(2);
    check(
        plainCell.fill === undefined || plainCell.fill.type === undefined,
        'non-PK cell should be unstyled',
    );
    check(!plainCell.font?.bold, 'data row should not be bold');

    if (expectedRows !== undefined) {
        check(
            sheet.rowCount === expectedRows,
            `expected ${expectedRows} rows, found ${sheet.rowCount}`,
        );
    }

    return {
        sheetNames,
        rows: sheet.rowCount,
        header,
        lastName: sheet.getRow(sheet.rowCount).getCell(2).value,
        frozen: checkFrozenPanes(sheet),
    };
}

async function main(): Promise<void> {
    const path = process.argv[2];
    if (!path) {
        console.error('Usage: node dist/scripts/validate-xlsx.js <file.xlsx> [expected_rows]');
        process.exitCode = 1;
        return;
    }
    const expectedRows = process.argv[3] ? Number(process.argv[3]) : undefined;

    const container = await checkContainer(path);
    const workbook = await checkWorkbook(path, expectedRows);

    const ratio = (container.sheetStoredBytes / container.sheetRawBytes) * 100;
    const n = (value: number): string => value.toLocaleString('en-US');

    console.log(`OK  ${path}`);
    console.log(`    parts        ${container.names.length}: ${container.names.join(', ')}`);
    console.log(`    sheets       ${workbook.sheetNames.length}: ${workbook.sheetNames.join(', ')}`);
    console.log(
        `    rows         ${workbook.rows} in the first sheet (header + ${workbook.rows - 1} data), ` +
            `last name ${JSON.stringify(workbook.lastName)}`,
    );
    console.log(`    columns      ${JSON.stringify(workbook.header)}`);
    console.log(`    frozen       ${workbook.frozen}`);
    console.log('    zip          2.0, no ZIP64, deflate, streamed (data descriptor), CRC verified');
    console.log(
        `    sheet1.xml   ${n(container.sheetRawBytes)} B -> ` +
            `${n(container.sheetStoredBytes)} B (${ratio.toFixed(1)}%)`,
    );
}

try {
    await main();
} catch (err) {
    console.error(`FAILED  ${process.argv[2] ?? ''}`);
    console.error(`        ${err instanceof Error ? err.message : String(err)}`);
    process.exitCode = 1;
}
