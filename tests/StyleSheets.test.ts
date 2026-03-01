import { describe, it, expect, beforeEach } from 'vitest';
import { VisioDocument } from '../src/VisioDocument';
import { StyleSheetManager } from '../src/core/StyleSheetManager';
import { VisioPackage } from '../src/VisioPackage';
import { XMLParser } from 'fast-xml-parser';

// ── Helpers ────────────────────────────────────────────────────────────────────

async function createDoc() {
    const doc  = await VisioDocument.create();
    const page = await doc.addPage('Test');
    return { doc, page };
}

/** Parse document.xml out of a saved buffer and return the VisioDocument node. */
async function getDocXml(doc: VisioDocument): Promise<any> {
    const buf    = await doc.save();
    const pkg    = new VisioPackage();
    await pkg.load(buf);
    const xml    = pkg.getFileText('visio/document.xml');
    const parser = new XMLParser({ ignoreAttributes: false, attributeNamePrefix: '@_' });
    const parsed = parser.parse(xml);
    return parsed['VisioDocument'];
}

function normArray(val: any): any[] {
    if (!val) return [];
    return Array.isArray(val) ? val : [val];
}

// ── StyleSheetManager unit tests ──────────────────────────────────────────────

describe('StyleSheetManager', () => {
    it('createStyle returns a StyleRecord with id >= 2', async () => {
        const { doc } = await createDoc();
        const record = doc.createStyle('MyStyle', { fillColor: '#ff0000' });
        expect(record.id).toBeGreaterThanOrEqual(2);
        expect(record.name).toBe('MyStyle');
    });

    it('createStyle IDs are sequential', async () => {
        const { doc } = await createDoc();
        const a = doc.createStyle('A');
        const b = doc.createStyle('B');
        const c = doc.createStyle('C');
        expect(b.id).toBe(a.id + 1);
        expect(c.id).toBe(b.id + 1);
    });

    it('getStyles returns built-in styles 0 and 1', async () => {
        const { doc } = await createDoc();
        const styles = doc.getStyles();
        const ids = styles.map(s => s.id);
        expect(ids).toContain(0);
        expect(ids).toContain(1);
    });

    it('getStyles includes newly created style', async () => {
        const { doc } = await createDoc();
        const record = doc.createStyle('Custom');
        const ids = doc.getStyles().map(s => s.id);
        expect(ids).toContain(record.id);
    });

    it('createStyle with no props still creates a valid entry', async () => {
        const { doc } = await createDoc();
        expect(() => doc.createStyle('Empty')).not.toThrow();
        const styles = doc.getStyles();
        expect(styles.some(s => s.name === 'Empty')).toBe(true);
    });
});

// ── Document XML structure ─────────────────────────────────────────────────────

describe('StyleSheets in document.xml', () => {
    it('new document contains Style 0 and Style 1', async () => {
        const { doc } = await createDoc();
        const vDoc     = await getDocXml(doc);
        const sheets   = normArray(vDoc?.StyleSheets?.StyleSheet);
        const ids = sheets.map((s: any) => s['@_ID']);
        expect(ids).toContain('0');
        expect(ids).toContain('1');
    });

    it('createStyle persists fill cells to document.xml', async () => {
        const { doc } = await createDoc();
        doc.createStyle('Blue', { fillColor: '#4472C4' });
        const vDoc   = await getDocXml(doc);
        const sheets = normArray(vDoc?.StyleSheets?.StyleSheet);
        const custom = sheets.find((s: any) => s['@_Name'] === 'Blue');
        expect(custom).toBeDefined();
        const cells = normArray(custom.Cell);
        const fill  = cells.find((c: any) => c['@_N'] === 'FillForegnd');
        expect(fill?.['@_V']).toBe('#4472C4');
        expect(fill?.['@_F']).toMatch(/^RGB\(/);
    });

    it('createStyle persists line cells', async () => {
        const { doc } = await createDoc();
        doc.createStyle('Thick', { lineColor: '#cc0000', lineWeight: 2, linePattern: 1 });
        const vDoc   = await getDocXml(doc);
        const sheets = normArray(vDoc?.StyleSheets?.StyleSheet);
        const custom = sheets.find((s: any) => s['@_Name'] === 'Thick');
        const cells  = normArray(custom.Cell);
        const lc = cells.find((c: any) => c['@_N'] === 'LineColor');
        const lw = cells.find((c: any) => c['@_N'] === 'LineWeight');
        const lp = cells.find((c: any) => c['@_N'] === 'LinePattern');
        expect(lc?.['@_V']).toBe('#cc0000');
        expect(parseFloat(lw?.['@_V'])).toBeCloseTo(2 / 72);
        expect(lp?.['@_V']).toBe('1');
    });

    it('createStyle persists Character section with font properties', async () => {
        const { doc } = await createDoc();
        doc.createStyle('Header', {
            fontColor:  '#ffffff',
            fontSize:   14,
            bold:       true,
            fontFamily: 'Calibri',
        });
        const vDoc   = await getDocXml(doc);
        const sheets = normArray(vDoc?.StyleSheets?.StyleSheet);
        const custom = sheets.find((s: any) => s['@_Name'] === 'Header');
        const sections = normArray(custom.Section);
        const charSec  = sections.find((s: any) => s['@_N'] === 'Character');
        expect(charSec).toBeDefined();
        const row   = charSec.Row;
        const cells = normArray(row.Cell);
        expect(cells.find((c: any) => c['@_N'] === 'Color')?.['@_V']).toBe('#ffffff');
        expect(cells.find((c: any) => c['@_N'] === 'Style')?.['@_V']).toBe('1'); // bold
        expect(parseFloat(cells.find((c: any) => c['@_N'] === 'Size')?.['@_V'])).toBeCloseTo(14 / 72);
        expect(cells.find((c: any) => c['@_N'] === 'Font')?.['@_F']).toBe('FONT("Calibri")');
    });

    it('createStyle persists Paragraph section', async () => {
        const { doc } = await createDoc();
        doc.createStyle('Para', { horzAlign: 'center', spaceBefore: 6, lineSpacing: 1.5 });
        const vDoc   = await getDocXml(doc);
        const sheets = normArray(vDoc?.StyleSheets?.StyleSheet);
        const custom = sheets.find((s: any) => s['@_Name'] === 'Para');
        const sections = normArray(custom.Section);
        const paraSec  = sections.find((s: any) => s['@_N'] === 'Paragraph');
        expect(paraSec).toBeDefined();
        const cells = normArray(paraSec.Row.Cell);
        expect(cells.find((c: any) => c['@_N'] === 'HorzAlign')?.['@_V']).toBe('1');
        expect(parseFloat(cells.find((c: any) => c['@_N'] === 'SpBefore')?.['@_V'])).toBeCloseTo(6 / 72);
        expect(parseFloat(cells.find((c: any) => c['@_N'] === 'SpLine')?.['@_V'])).toBeCloseTo(-1.5);
    });

    it('createStyle persists TextBlock cells (margins, vertical align)', async () => {
        const { doc } = await createDoc();
        doc.createStyle('Box', {
            verticalAlign:    'middle',
            textMarginTop:    0.1,
            textMarginBottom: 0.1,
        });
        const vDoc   = await getDocXml(doc);
        const sheets = normArray(vDoc?.StyleSheets?.StyleSheet);
        const custom = sheets.find((s: any) => s['@_Name'] === 'Box');
        const cells  = normArray(custom.Cell);
        expect(cells.find((c: any) => c['@_N'] === 'VerticalAlign')?.['@_V']).toBe('1');
        expect(parseFloat(cells.find((c: any) => c['@_N'] === 'TopMargin')?.['@_V'])).toBeCloseTo(0.1);
        expect(parseFloat(cells.find((c: any) => c['@_N'] === 'BottomMargin')?.['@_V'])).toBeCloseTo(0.1);
    });

    it('Style row carries correct parent style refs', async () => {
        const { doc } = await createDoc();
        doc.createStyle('Child', {
            parentLineStyleId: 1,
            parentFillStyleId: 1,
            parentTextStyleId: 1,
        });
        const vDoc   = await getDocXml(doc);
        const sheets = normArray(vDoc?.StyleSheets?.StyleSheet);
        const custom = sheets.find((s: any) => s['@_Name'] === 'Child');
        expect(custom['@_LineStyle']).toBe('1');
        expect(custom['@_FillStyle']).toBe('1');
        expect(custom['@_TextStyle']).toBe('1');
    });

    it('Character Style bitmask encodes italic+underline correctly', async () => {
        const { doc } = await createDoc();
        doc.createStyle('ItalicUnderline', { italic: true, underline: true });
        const vDoc   = await getDocXml(doc);
        const sheets = normArray(vDoc?.StyleSheets?.StyleSheet);
        const custom = sheets.find((s: any) => s['@_Name'] === 'ItalicUnderline');
        const charSec = normArray(custom.Section).find((s: any) => s['@_N'] === 'Character');
        const cells  = normArray(charSec.Row.Cell);
        // italic=2, underline=4 → 6
        expect(cells.find((c: any) => c['@_N'] === 'Style')?.['@_V']).toBe('6');
    });
});

// ── Applying styles to shapes ─────────────────────────────────────────────────

describe('Applying stylesheets to shapes', () => {
    it('addShape with styleId sets LineStyle, FillStyle, TextStyle on shape XML', async () => {
        const { doc, page } = await createDoc();
        const style = doc.createStyle('Blue', { fillColor: '#4472C4' });
        await page.addShape({ text: 'Styled', x: 1, y: 1, width: 2, height: 1, styleId: style.id });
        const buf  = await doc.save();
        const pkg2 = new VisioPackage();
        await pkg2.load(buf);
        const xml = pkg2.getFileText('visio/pages/page2.xml'); // page added after default page 1
        const parser = new XMLParser({ ignoreAttributes: false, attributeNamePrefix: '@_' });
        const parsed = parser.parse(xml);
        const shapes = parsed.PageContents?.Shapes?.Shape;
        const shapeArr = Array.isArray(shapes) ? shapes : [shapes];
        const s = shapeArr.find((sh: any) => sh['@_LineStyle']);
        expect(s?.['@_LineStyle']).toBe(style.id.toString());
        expect(s?.['@_FillStyle']).toBe(style.id.toString());
        expect(s?.['@_TextStyle']).toBe(style.id.toString());
    });

    it('addShape with lineStyleId sets only LineStyle', async () => {
        const { doc, page } = await createDoc();
        const style = doc.createStyle('LineOnly', { lineColor: '#ff0000' });
        await page.addShape({ text: 'X', x: 1, y: 1, width: 2, height: 1, lineStyleId: style.id });
        const buf  = await doc.save();
        const pkg2 = new VisioPackage();
        await pkg2.load(buf);
        const xml = pkg2.getFileText('visio/pages/page2.xml');
        const parser = new XMLParser({ ignoreAttributes: false, attributeNamePrefix: '@_' });
        const parsed = parser.parse(xml);
        const shapes = normArray(parsed.PageContents?.Shapes?.Shape);
        const s = shapes.find((sh: any) => sh['@_LineStyle'] === style.id.toString());
        expect(s?.['@_LineStyle']).toBe(style.id.toString());
        expect(s?.['@_FillStyle']).toBeUndefined();
    });

    it('shape.applyStyle(id) sets all three attributes', async () => {
        const { doc, page } = await createDoc();
        const style = doc.createStyle('Red', { fillColor: '#ff0000' });
        const shape = await page.addShape({ text: 'A', x: 1, y: 1, width: 2, height: 1 });
        shape.applyStyle(style.id);
        const buf  = await doc.save();
        const pkg2 = new VisioPackage();
        await pkg2.load(buf);
        const xml = pkg2.getFileText('visio/pages/page2.xml');
        const parser = new XMLParser({ ignoreAttributes: false, attributeNamePrefix: '@_' });
        const parsed = parser.parse(xml);
        const shapes = normArray(parsed.PageContents?.Shapes?.Shape);
        const s = shapes.find((sh: any) => sh['@_LineStyle'] === style.id.toString());
        expect(s?.['@_LineStyle']).toBe(style.id.toString());
        expect(s?.['@_FillStyle']).toBe(style.id.toString());
        expect(s?.['@_TextStyle']).toBe(style.id.toString());
    });

    it('shape.applyStyle(id, "fill") sets only FillStyle', async () => {
        const { doc, page } = await createDoc();
        const style = doc.createStyle('FillOnly', { fillColor: '#00ff00' });
        const shape = await page.addShape({ text: 'B', x: 1, y: 1, width: 2, height: 1 });
        shape.applyStyle(style.id, 'fill');
        const buf  = await doc.save();
        const pkg2 = new VisioPackage();
        await pkg2.load(buf);
        const xml = pkg2.getFileText('visio/pages/page2.xml');
        const parser = new XMLParser({ ignoreAttributes: false, attributeNamePrefix: '@_' });
        const parsed = parser.parse(xml);
        const shapes = normArray(parsed.PageContents?.Shapes?.Shape);
        const s = shapes.find((sh: any) => sh['@_FillStyle'] === style.id.toString());
        expect(s?.['@_FillStyle']).toBe(style.id.toString());
        expect(s?.['@_LineStyle']).toBeUndefined();
        expect(s?.['@_TextStyle']).toBeUndefined();
    });

    it('shape.applyStyle is chainable', async () => {
        const { doc, page } = await createDoc();
        const style = doc.createStyle('Chain', { fillColor: '#123456' });
        const shape = await page.addShape({ text: 'C', x: 1, y: 1, width: 2, height: 1 });
        const result = shape.applyStyle(style.id);
        expect(result).toBe(shape);
    });

    it('applyStyle on unknown shapeId throws', async () => {
        const { doc } = await createDoc();
        const style = doc.createStyle('X', {});
        const pkg2  = await VisioPackage.create();
        const { ShapeModifier } = await import('../src/ShapeModifier');
        const mod = new ShapeModifier(pkg2);
        expect(() => mod.applyStyle('1', '999', style.id)).toThrow();
    });
});

// ── Full round-trip: create style → apply → save → reload ─────────────────────

describe('StyleSheet round-trip', () => {
    it('style definition survives save/reload cycle', async () => {
        const doc1 = await VisioDocument.create();
        doc1.createStyle('Persistent', { fillColor: '#abcdef', lineWeight: 3 });
        const buf = await doc1.save();

        const doc2   = await VisioDocument.load(buf);
        const styles = doc2.getStyles();
        const found  = styles.find(s => s.name === 'Persistent');
        expect(found).toBeDefined();
        expect(found!.id).toBeGreaterThanOrEqual(2);
    });

    it('multiple styles survive save/reload', async () => {
        const doc1 = await VisioDocument.create();
        const s1 = doc1.createStyle('Alpha', { fillColor: '#111111' });
        const s2 = doc1.createStyle('Beta',  { lineColor: '#222222' });
        const s3 = doc1.createStyle('Gamma', { fontColor: '#333333' });
        const buf = await doc1.save();

        const doc2   = await VisioDocument.load(buf);
        const styles = doc2.getStyles();
        const ids    = styles.map(s => s.id);
        expect(ids).toContain(s1.id);
        expect(ids).toContain(s2.id);
        expect(ids).toContain(s3.id);
    });
});
