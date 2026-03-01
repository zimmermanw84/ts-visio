import { describe, it, expect, afterEach } from 'vitest';
import { VisioDocument } from '../src/VisioDocument';
import { VisioPackage } from '../src/VisioPackage';
import { XMLParser } from 'fast-xml-parser';
import fs from 'fs';
import path from 'path';

// ── helpers ────────────────────────────────────────────────────────────────────

/** Save doc, reload, and return the parsed page XML object. */
async function getPageXml(doc: VisioDocument, pageFile = 'visio/pages/page1.xml') {
    const buf    = await doc.save();
    const pkg    = new VisioPackage();
    await pkg.load(buf);
    const xml    = pkg.getFileText(pageFile);
    const parser = new XMLParser({ ignoreAttributes: false, attributeNamePrefix: '@_' });
    return parser.parse(xml);
}

function normArray(val: any): any[] {
    if (!val) return [];
    return Array.isArray(val) ? val : [val];
}

// ── legacy smoke tests ─────────────────────────────────────────────────────────

describe('Styling', () => {
    const testFile = path.resolve(__dirname, 'style_test.vsdx');

    afterEach(() => {
        if (fs.existsSync(testFile)) {
            fs.unlinkSync(testFile);
        }
    });

    it('should create a shape with fill color', async () => {
        const doc = await VisioDocument.create();
        const page = doc.pages[0];

        const shape = await page.addShape({
            text: 'Colorful Box',
            x: 2,
            y: 4,
            width: 2,
            height: 1,
            fillColor: '#FF0000'
        });

        expect(shape).toBeDefined();

        await doc.save(testFile);
        expect(fs.existsSync(testFile)).toBe(true);
    });

    it('should create a shape with bold text and color', async () => {
        const doc = await VisioDocument.create();
        const page = doc.pages[0];

        await page.addShape({
            text: 'Bold Red Text',
            x: 5,
            y: 5,
            width: 2,
            height: 1,
            fillColor: '#FFFFFF',
            fontColor: '#FF0000',
            bold: true
        });

        await doc.save(testFile);
        expect(fs.existsSync(testFile)).toBe(true);
    });

    it('should allow fluent chaining of styles', async () => {
        const doc = await VisioDocument.create();
        const page = doc.pages[0];
        const shape = await page.addShape({ text: 'Fluent', x: 2, y: 2, width: 2, height: 1 });

        await (await shape.setStyle({ fillColor: '#00FF00' })).setStyle({ bold: true });

        const s2 = await shape.setStyle({ fontColor: '#0000FF' });
        expect(s2).toBe(shape);
    });

    it('should allow linking with connectTo and chaining', async () => {
        const doc = await VisioDocument.create();
        const page = doc.pages[0];
        const s1 = await page.addShape({ text: 'A', x: 0, y: 0, width: 1, height: 1 });
        const s2 = await page.addShape({ text: 'B', x: 4, y: 4, width: 1, height: 1 });

        const ret = await s1.connectTo(s2);
        expect(ret).toBe(s1);
    });
});

// ── setStyle line properties ───────────────────────────────────────────────────

describe('shape.setStyle() — line properties', () => {
    it('setStyle({ lineColor }) persists LineColor to page XML', async () => {
        const doc  = await VisioDocument.create();
        const page = doc.pages[0];
        const shape = await page.addShape({ text: 'A', x: 1, y: 1, width: 2, height: 1 });

        await shape.setStyle({ lineColor: '#cc0000' });

        const parsed = await getPageXml(doc);
        const shapes = normArray(parsed.PageContents?.Shapes?.Shape);
        const s = shapes.find((sh: any) => sh['@_ID'] === shape.id);
        const sections = normArray(s?.Section);
        const lineSec  = sections.find((sec: any) => sec['@_N'] === 'Line');
        const cells    = normArray(lineSec?.Cell);
        expect(cells.find((c: any) => c['@_N'] === 'LineColor')?.['@_V']).toBe('#cc0000');
    });

    it('setStyle({ lineWeight }) converts points to inches in XML', async () => {
        const doc  = await VisioDocument.create();
        const page = doc.pages[0];
        const shape = await page.addShape({ text: 'B', x: 1, y: 1, width: 2, height: 1 });

        await shape.setStyle({ lineWeight: 2 }); // 2 pt → 2/72 in

        const parsed = await getPageXml(doc);
        const shapes = normArray(parsed.PageContents?.Shapes?.Shape);
        const s = shapes.find((sh: any) => sh['@_ID'] === shape.id);
        const sections = normArray(s?.Section);
        const lineSec  = sections.find((sec: any) => sec['@_N'] === 'Line');
        const cells    = normArray(lineSec?.Cell);
        const lw = cells.find((c: any) => c['@_N'] === 'LineWeight');
        expect(parseFloat(lw?.['@_V'])).toBeCloseTo(2 / 72, 6);
    });

    it('setStyle({ linePattern }) persists pattern value', async () => {
        const doc  = await VisioDocument.create();
        const page = doc.pages[0];
        const shape = await page.addShape({ text: 'C', x: 1, y: 1, width: 2, height: 1 });

        await shape.setStyle({ linePattern: 2 }); // dashed

        const parsed = await getPageXml(doc);
        const shapes = normArray(parsed.PageContents?.Shapes?.Shape);
        const s = shapes.find((sh: any) => sh['@_ID'] === shape.id);
        const sections = normArray(s?.Section);
        const lineSec  = sections.find((sec: any) => sec['@_N'] === 'Line');
        const cells    = normArray(lineSec?.Cell);
        expect(cells.find((c: any) => c['@_N'] === 'LinePattern')?.['@_V']).toBe('2');
    });

    it('setStyle with all three line props together', async () => {
        const doc  = await VisioDocument.create();
        const page = doc.pages[0];
        const shape = await page.addShape({ text: 'D', x: 1, y: 1, width: 2, height: 1 });

        await shape.setStyle({ lineColor: '#0000ff', lineWeight: 3, linePattern: 3 });

        const parsed = await getPageXml(doc);
        const shapes = normArray(parsed.PageContents?.Shapes?.Shape);
        const s = shapes.find((sh: any) => sh['@_ID'] === shape.id);
        const sections = normArray(s?.Section);
        const lineSec  = sections.find((sec: any) => sec['@_N'] === 'Line');
        const cells    = normArray(lineSec?.Cell);
        expect(cells.find((c: any) => c['@_N'] === 'LineColor')?.['@_V']).toBe('#0000ff');
        expect(parseFloat(cells.find((c: any) => c['@_N'] === 'LineWeight')?.['@_V'])).toBeCloseTo(3 / 72, 6);
        expect(cells.find((c: any) => c['@_N'] === 'LinePattern')?.['@_V']).toBe('3');
    });

    it('setStyle line and fill together work independently', async () => {
        const doc  = await VisioDocument.create();
        const page = doc.pages[0];
        const shape = await page.addShape({ text: 'E', x: 1, y: 1, width: 2, height: 1 });

        await shape.setStyle({ fillColor: '#ffff00', lineColor: '#ff00ff' });

        const parsed = await getPageXml(doc);
        const shapes = normArray(parsed.PageContents?.Shapes?.Shape);
        const s = shapes.find((sh: any) => sh['@_ID'] === shape.id);
        const sections = normArray(s?.Section);

        const fillSec = sections.find((sec: any) => sec['@_N'] === 'Fill');
        const fillCells = normArray(fillSec?.Cell);
        expect(fillCells.find((c: any) => c['@_N'] === 'FillForegnd')?.['@_V']).toBe('#ffff00');

        const lineSec = sections.find((sec: any) => sec['@_N'] === 'Line');
        const lineCells = normArray(lineSec?.Cell);
        expect(lineCells.find((c: any) => c['@_N'] === 'LineColor')?.['@_V']).toBe('#ff00ff');
    });

    it('setStyle is chainable and returns the Shape', async () => {
        const doc  = await VisioDocument.create();
        const page = doc.pages[0];
        const shape = await page.addShape({ text: 'F', x: 1, y: 1, width: 2, height: 1 });
        const result = await shape.setStyle({ lineColor: '#aabbcc' });
        expect(result).toBe(shape);
    });

    it('second setStyle call replaces the Line section', async () => {
        const doc  = await VisioDocument.create();
        const page = doc.pages[0];
        const shape = await page.addShape({ text: 'G', x: 1, y: 1, width: 2, height: 1 });

        await shape.setStyle({ lineColor: '#111111' });
        await shape.setStyle({ lineColor: '#999999' });

        const parsed = await getPageXml(doc);
        const shapes = normArray(parsed.PageContents?.Shapes?.Shape);
        const s = shapes.find((sh: any) => sh['@_ID'] === shape.id);
        const sections = normArray(s?.Section);
        const lineSections = sections.filter((sec: any) => sec['@_N'] === 'Line');
        // Must be only one Line section
        expect(lineSections).toHaveLength(1);
        const cells = normArray(lineSections[0].Cell);
        expect(cells.find((c: any) => c['@_N'] === 'LineColor')?.['@_V']).toBe('#999999');
    });

    it('lineColor written to page XML has an RGB formula', async () => {
        const doc  = await VisioDocument.create();
        const page = doc.pages[0];
        const shape = await page.addShape({ text: 'H', x: 1, y: 1, width: 2, height: 1 });

        await shape.setStyle({ lineColor: '#ff0000' });

        const parsed = await getPageXml(doc);
        const shapes = normArray(parsed.PageContents?.Shapes?.Shape);
        const s = shapes.find((sh: any) => sh['@_ID'] === shape.id);
        const lineSec = normArray(s?.Section).find((sec: any) => sec['@_N'] === 'Line');
        const cells   = normArray(lineSec?.Cell);
        const lc = cells.find((c: any) => c['@_N'] === 'LineColor');
        expect(lc?.['@_F']).toMatch(/^RGB\(/);
    });
});

// ── round-trip ─────────────────────────────────────────────────────────────────

describe('shape.setStyle() line — round-trip save/reload', () => {
    it('lineColor survives save/reload', async () => {
        const doc1  = await VisioDocument.create();
        const shape = await doc1.pages[0].addShape({ text: 'X', x: 1, y: 1, width: 2, height: 1 });
        await shape.setStyle({ lineColor: '#abcdef' });

        const buf  = await doc1.save();
        const doc2 = await VisioDocument.load(buf);
        const parsed = await getPageXml(doc2);
        const shapes = normArray(parsed.PageContents?.Shapes?.Shape);
        const s = shapes.find((sh: any) => sh['@_ID'] === shape.id);
        const lineSec = normArray(s?.Section).find((sec: any) => sec['@_N'] === 'Line');
        const cells   = normArray(lineSec?.Cell);
        expect(cells.find((c: any) => c['@_N'] === 'LineColor')?.['@_V']).toBe('#abcdef');
    });

    it('lineWeight survives save/reload', async () => {
        const doc1  = await VisioDocument.create();
        const shape = await doc1.pages[0].addShape({ text: 'Y', x: 1, y: 1, width: 2, height: 1 });
        await shape.setStyle({ lineWeight: 4 });

        const buf  = await doc1.save();
        const doc2 = await VisioDocument.load(buf);
        const parsed = await getPageXml(doc2);
        const shapes = normArray(parsed.PageContents?.Shapes?.Shape);
        const s = shapes.find((sh: any) => sh['@_ID'] === shape.id);
        const lineSec = normArray(s?.Section).find((sec: any) => sec['@_N'] === 'Line');
        const cells   = normArray(lineSec?.Cell);
        expect(parseFloat(cells.find((c: any) => c['@_N'] === 'LineWeight')?.['@_V'])).toBeCloseTo(4 / 72, 6);
    });
});
