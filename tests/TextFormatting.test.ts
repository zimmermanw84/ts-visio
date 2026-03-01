import { describe, it, expect } from 'vitest';
import { VisioDocument } from '../src/VisioDocument';
import { XMLParser } from 'fast-xml-parser';

const parser = new XMLParser({ ignoreAttributes: false, attributeNamePrefix: '@_' });

/** Parse the first page XML from a saved doc buffer and return the raw shape list. */
async function getPageShapes(doc: VisioDocument): Promise<any[]> {
    const buf = await doc.save();
    const JSZip = (await import('jszip')).default;
    const zip = await JSZip.loadAsync(buf);
    const xml = await zip.file('visio/pages/page1.xml')!.async('string');
    const parsed = parser.parse(xml);
    const shapes = parsed.PageContents?.Shapes?.Shape;
    if (!shapes) return [];
    return Array.isArray(shapes) ? shapes : [shapes];
}

/** Find a cell by name in a shape's top-level Cell array. */
function findCell(shape: any, name: string): any {
    const cells = Array.isArray(shape.Cell) ? shape.Cell : [shape.Cell];
    return cells.find((c: any) => c['@_N'] === name);
}

/** Find a row's cell in a named section. */
function findSectionRowCell(shape: any, sectionName: string, cellName: string): any {
    const sections = Array.isArray(shape.Section) ? shape.Section : [shape.Section];
    const sec = sections?.find((s: any) => s['@_N'] === sectionName);
    if (!sec) return undefined;
    const rows = Array.isArray(sec.Row) ? sec.Row : [sec.Row];
    const row = rows[0];
    if (!row) return undefined;
    const cells = Array.isArray(row.Cell) ? row.Cell : [row.Cell];
    return cells.find((c: any) => c['@_N'] === cellName);
}

describe('Text Formatting — Font Size', () => {
    it('creates a Size cell in the Character section (addShape)', async () => {
        const doc = await VisioDocument.create();
        const page = doc.pages[0];
        await page.addShape({ text: 'Big Text', x: 2, y: 2, width: 3, height: 1, fontSize: 18 });

        const shapes = await getPageShapes(doc);
        const sizeCell = findSectionRowCell(shapes[0], 'Character', 'Size');
        expect(sizeCell).toBeDefined();
        // 18pt → 18/72 = 0.25 inches
        expect(parseFloat(sizeCell['@_V'])).toBeCloseTo(18 / 72, 8);
        expect(sizeCell['@_U']).toBe('PT');
    });

    it('creates a Size cell via setStyle', async () => {
        const doc = await VisioDocument.create();
        const page = doc.pages[0];
        const shape = await page.addShape({ text: 'Hello', x: 2, y: 2, width: 2, height: 1 });
        await shape.setStyle({ fontSize: 12 });

        const shapes = await getPageShapes(doc);
        const sizeCell = findSectionRowCell(shapes[0], 'Character', 'Size');
        expect(sizeCell).toBeDefined();
        expect(parseFloat(sizeCell['@_V'])).toBeCloseTo(12 / 72, 8);
    });

    it('combining fontSize with bold keeps both values', async () => {
        const doc = await VisioDocument.create();
        const page = doc.pages[0];
        await page.addShape({ text: 'Bold Big', x: 2, y: 2, width: 2, height: 1, fontSize: 14, bold: true });

        const shapes = await getPageShapes(doc);
        const sizeCell = findSectionRowCell(shapes[0], 'Character', 'Size');
        const styleCell = findSectionRowCell(shapes[0], 'Character', 'Style');
        expect(parseFloat(sizeCell['@_V'])).toBeCloseTo(14 / 72, 8);
        expect(styleCell['@_V']).toBe('1'); // Bold bit set
    });
});

describe('Text Formatting — Font Family', () => {
    it('creates a Font cell with FONT() formula (addShape)', async () => {
        const doc = await VisioDocument.create();
        const page = doc.pages[0];
        await page.addShape({ text: 'Arial Text', x: 2, y: 2, width: 3, height: 1, fontFamily: 'Arial' });

        const shapes = await getPageShapes(doc);
        const fontCell = findSectionRowCell(shapes[0], 'Character', 'Font');
        expect(fontCell).toBeDefined();
        expect(fontCell['@_F']).toBe('FONT("Arial")');
        expect(fontCell['@_V']).toBe('0'); // placeholder ID
    });

    it('creates a Font cell with FONT() formula via setStyle', async () => {
        const doc = await VisioDocument.create();
        const page = doc.pages[0];
        const shape = await page.addShape({ text: 'Hello', x: 2, y: 2, width: 2, height: 1 });
        await shape.setStyle({ fontFamily: 'Times New Roman' });

        const shapes = await getPageShapes(doc);
        const fontCell = findSectionRowCell(shapes[0], 'Character', 'Font');
        expect(fontCell?.['@_F']).toBe('FONT("Times New Roman")');
    });

    it('omitting fontFamily uses default Font cell (no formula)', async () => {
        const doc = await VisioDocument.create();
        const page = doc.pages[0];
        await page.addShape({ text: 'Default', x: 2, y: 2, width: 2, height: 1, bold: true });

        const shapes = await getPageShapes(doc);
        const fontCell = findSectionRowCell(shapes[0], 'Character', 'Font');
        expect(fontCell?.['@_V']).toBe('1'); // Default Calibri
        expect(fontCell?.['@_F']).toBeUndefined();
    });
});

describe('Text Formatting — Horizontal Alignment', () => {
    it('creates a Paragraph section with HorzAlign center (addShape)', async () => {
        const doc = await VisioDocument.create();
        const page = doc.pages[0];
        await page.addShape({ text: 'Centered', x: 2, y: 2, width: 3, height: 1, horzAlign: 'center' });

        const shapes = await getPageShapes(doc);
        const horzCell = findSectionRowCell(shapes[0], 'Paragraph', 'HorzAlign');
        expect(horzCell).toBeDefined();
        expect(horzCell['@_V']).toBe('1');
    });

    it('creates HorzAlign left (0)', async () => {
        const doc = await VisioDocument.create();
        const page = doc.pages[0];
        await page.addShape({ text: 'Left', x: 2, y: 2, width: 2, height: 1, horzAlign: 'left' });

        const shapes = await getPageShapes(doc);
        const horzCell = findSectionRowCell(shapes[0], 'Paragraph', 'HorzAlign');
        expect(horzCell['@_V']).toBe('0');
    });

    it('creates HorzAlign right (2)', async () => {
        const doc = await VisioDocument.create();
        const page = doc.pages[0];
        await page.addShape({ text: 'Right', x: 2, y: 2, width: 2, height: 1, horzAlign: 'right' });

        const shapes = await getPageShapes(doc);
        expect(findSectionRowCell(shapes[0], 'Paragraph', 'HorzAlign')['@_V']).toBe('2');
    });

    it('creates HorzAlign justify (3)', async () => {
        const doc = await VisioDocument.create();
        const page = doc.pages[0];
        await page.addShape({ text: 'Justified', x: 2, y: 2, width: 2, height: 1, horzAlign: 'justify' });

        const shapes = await getPageShapes(doc);
        expect(findSectionRowCell(shapes[0], 'Paragraph', 'HorzAlign')['@_V']).toBe('3');
    });

    it('sets horzAlign via setStyle', async () => {
        const doc = await VisioDocument.create();
        const page = doc.pages[0];
        const shape = await page.addShape({ text: 'Hello', x: 2, y: 2, width: 2, height: 1 });
        await shape.setStyle({ horzAlign: 'right' });

        const shapes = await getPageShapes(doc);
        expect(findSectionRowCell(shapes[0], 'Paragraph', 'HorzAlign')['@_V']).toBe('2');
    });
});

describe('Text Formatting — Vertical Alignment', () => {
    it('creates a VerticalAlign top-level cell (middle)', async () => {
        const doc = await VisioDocument.create();
        const page = doc.pages[0];
        await page.addShape({ text: 'Middle', x: 2, y: 2, width: 3, height: 2, verticalAlign: 'middle' });

        const shapes = await getPageShapes(doc);
        const vaCell = findCell(shapes[0], 'VerticalAlign');
        expect(vaCell).toBeDefined();
        expect(vaCell['@_V']).toBe('1');
    });

    it('creates VerticalAlign top (0)', async () => {
        const doc = await VisioDocument.create();
        const page = doc.pages[0];
        await page.addShape({ text: 'Top', x: 2, y: 2, width: 2, height: 2, verticalAlign: 'top' });

        const shapes = await getPageShapes(doc);
        expect(findCell(shapes[0], 'VerticalAlign')['@_V']).toBe('0');
    });

    it('creates VerticalAlign bottom (2)', async () => {
        const doc = await VisioDocument.create();
        const page = doc.pages[0];
        await page.addShape({ text: 'Bottom', x: 2, y: 2, width: 2, height: 2, verticalAlign: 'bottom' });

        const shapes = await getPageShapes(doc);
        expect(findCell(shapes[0], 'VerticalAlign')['@_V']).toBe('2');
    });

    it('sets verticalAlign via setStyle', async () => {
        const doc = await VisioDocument.create();
        const page = doc.pages[0];
        const shape = await page.addShape({ text: 'Hello', x: 2, y: 2, width: 2, height: 2 });
        await shape.setStyle({ verticalAlign: 'bottom' });

        const shapes = await getPageShapes(doc);
        expect(findCell(shapes[0], 'VerticalAlign')['@_V']).toBe('2');
    });

    it('updating verticalAlign via setStyle replaces existing cell (no duplicates)', async () => {
        const doc = await VisioDocument.create();
        const page = doc.pages[0];
        const shape = await page.addShape({ text: 'Hello', x: 2, y: 2, width: 2, height: 2, verticalAlign: 'top' });
        await shape.setStyle({ verticalAlign: 'middle' });

        const shapes = await getPageShapes(doc);
        const cells = Array.isArray(shapes[0].Cell) ? shapes[0].Cell : [shapes[0].Cell];
        const vaCells = cells.filter((c: any) => c['@_N'] === 'VerticalAlign');
        expect(vaCells).toHaveLength(1);
        expect(vaCells[0]['@_V']).toBe('1');
    });
});

describe('Text Formatting — Combined', () => {
    it('supports all four properties simultaneously on addShape', async () => {
        const doc = await VisioDocument.create();
        const page = doc.pages[0];
        await page.addShape({
            text: 'Full Style',
            x: 2, y: 2, width: 4, height: 2,
            fontSize: 16,
            fontFamily: 'Georgia',
            horzAlign: 'center',
            verticalAlign: 'middle',
        });

        const shapes = await getPageShapes(doc);
        const shape = shapes[0];

        expect(parseFloat(findSectionRowCell(shape, 'Character', 'Size')['@_V'])).toBeCloseTo(16 / 72, 8);
        expect(findSectionRowCell(shape, 'Character', 'Font')['@_F']).toBe('FONT("Georgia")');
        expect(findSectionRowCell(shape, 'Paragraph', 'HorzAlign')['@_V']).toBe('1');
        expect(findCell(shape, 'VerticalAlign')['@_V']).toBe('1');
    });
});
