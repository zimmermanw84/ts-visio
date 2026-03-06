import { describe, it, expect } from 'vitest';
import JSZip from 'jszip';
import { VisioDocument } from '../src/VisioDocument';
import { Connector } from '../src/Connector';
import { ArrowHeads } from '../src/utils/StyleHelpers';

// ── helpers ────────────────────────────────────────────────────────────────────

async function twoShapePage() {
    const doc  = await VisioDocument.create();
    const page = doc.pages[0];
    const a    = await page.addShape({ text: 'A', x: 1, y: 1, width: 1, height: 1 });
    const b    = await page.addShape({ text: 'B', x: 5, y: 1, width: 1, height: 1 });
    return { doc, page, a, b };
}

// ── basic read-back ────────────────────────────────────────────────────────────

describe('page.getConnectors()', () => {
    it('returns empty array when no connectors exist', async () => {
        const { page } = await twoShapePage();
        expect(page.getConnectors()).toEqual([]);
    });

    it('returns one Connector after connectShapes', async () => {
        const { page, a, b } = await twoShapePage();
        await page.connectShapes(a, b);
        const connectors = page.getConnectors();
        expect(connectors).toHaveLength(1);
        expect(connectors[0]).toBeInstanceOf(Connector);
    });

    it('connector.id is a non-empty string', async () => {
        const { page, a, b } = await twoShapePage();
        await page.connectShapes(a, b);
        const [conn] = page.getConnectors();
        expect(typeof conn.id).toBe('string');
        expect(conn.id.length).toBeGreaterThan(0);
    });

    it('connector.fromShapeId and toShapeId match the connected shapes', async () => {
        const { page, a, b } = await twoShapePage();
        await page.connectShapes(a, b);
        const [conn] = page.getConnectors();
        expect(conn.fromShapeId).toBe(a.id);
        expect(conn.toShapeId).toBe(b.id);
    });

    it('returns multiple connectors when multiple connections exist', async () => {
        const { page, a, b } = await twoShapePage();
        const c = await page.addShape({ text: 'C', x: 3, y: 4, width: 1, height: 1 });
        await page.connectShapes(a, b);
        await page.connectShapes(b, c);
        expect(page.getConnectors()).toHaveLength(2);
    });
});

// ── port targets ───────────────────────────────────────────────────────────────

describe('Connector port targets', () => {
    it('fromPort and toPort are "center" for default connections', async () => {
        const { page, a, b } = await twoShapePage();
        await page.connectShapes(a, b);
        const [conn] = page.getConnectors();
        expect(conn.fromPort).toBe('center');
        expect(conn.toPort).toBe('center');
    });

    it('fromPort is "center" when connectShapes uses default port', async () => {
        const { page, a, b } = await twoShapePage();
        await page.connectShapes(a, b);
        const [conn] = page.getConnectors();
        expect(conn.fromPort).toBe('center');
    });
});

// ── line style ─────────────────────────────────────────────────────────────────

describe('Connector style properties', () => {
    it('default connector has a lineColor in style', async () => {
        const { page, a, b } = await twoShapePage();
        await page.connectShapes(a, b);
        const [conn] = page.getConnectors();
        // Default line color is #000000
        expect(conn.style.lineColor).toBeDefined();
    });

    it('custom lineColor is preserved', async () => {
        const { page, a, b } = await twoShapePage();
        await page.connectShapes(a, b, undefined, undefined, { lineColor: '#ff0000' });
        const [conn] = page.getConnectors();
        expect(conn.style.lineColor).toBe('#ff0000');
    });

    it('custom lineWeight is converted back to points', async () => {
        const { page, a, b } = await twoShapePage();
        await page.connectShapes(a, b, undefined, undefined, { lineWeight: 2 });
        const [conn] = page.getConnectors();
        // lineWeight stored as inches (2/72), read back and multiplied by 72 → ~2
        expect(conn.style.lineWeight).toBeCloseTo(2, 1);
    });

    it('linePattern is preserved', async () => {
        const { page, a, b } = await twoShapePage();
        await page.connectShapes(a, b, undefined, undefined, { linePattern: 2 });
        const [conn] = page.getConnectors();
        expect(conn.style.linePattern).toBe(2);
    });

    it('routing style "straight" is preserved', async () => {
        const { page, a, b } = await twoShapePage();
        await page.connectShapes(a, b, undefined, undefined, { routing: 'straight' });
        const [conn] = page.getConnectors();
        expect(conn.style.routing).toBe('straight');
    });

    it('routing style "orthogonal" is preserved (default)', async () => {
        const { page, a, b } = await twoShapePage();
        await page.connectShapes(a, b);
        const [conn] = page.getConnectors();
        // ShapeRouteStyle=1 → orthogonal
        expect(conn.style.routing).toBe('orthogonal');
    });

    it('routing style "curved" is preserved', async () => {
        const { page, a, b } = await twoShapePage();
        await page.connectShapes(a, b, undefined, undefined, { routing: 'curved' });
        const [conn] = page.getConnectors();
        expect(conn.style.routing).toBe('curved');
    });
});

// ── arrow heads ────────────────────────────────────────────────────────────────

describe('Connector arrow heads', () => {
    it('default connector has beginArrow "0" (none)', async () => {
        const { page, a, b } = await twoShapePage();
        await page.connectShapes(a, b);
        const [conn] = page.getConnectors();
        expect(conn.beginArrow).toBe('0');
    });

    it('default connector has endArrow "0" (none)', async () => {
        const { page, a, b } = await twoShapePage();
        await page.connectShapes(a, b);
        const [conn] = page.getConnectors();
        expect(conn.endArrow).toBe('0');
    });

    it('custom arrows are preserved', async () => {
        const { page, a, b } = await twoShapePage();
        await page.connectShapes(a, b, ArrowHeads.Standard, ArrowHeads.Open);
        const [conn] = page.getConnectors();
        expect(conn.beginArrow).toBe(ArrowHeads.Standard);
        expect(conn.endArrow).toBe(ArrowHeads.Open);
    });
});

// ── delete ─────────────────────────────────────────────────────────────────────

describe('Connector.delete()', () => {
    it('removes the connector from getConnectors()', async () => {
        const { page, a, b } = await twoShapePage();
        await page.connectShapes(a, b);
        const [conn] = page.getConnectors();
        await conn.delete();
        expect(page.getConnectors()).toHaveLength(0);
    });

    it('does not remove other connectors', async () => {
        const { page, a, b } = await twoShapePage();
        const c = await page.addShape({ text: 'C', x: 3, y: 4, width: 1, height: 1 });
        await page.connectShapes(a, b);
        await page.connectShapes(b, c);
        const [first] = page.getConnectors();
        await first.delete();
        expect(page.getConnectors()).toHaveLength(1);
    });

    it('does not remove the connected shapes themselves', async () => {
        const { page, a, b } = await twoShapePage();
        await page.connectShapes(a, b);
        const [conn] = page.getConnectors();
        await conn.delete();
        const shapes = page.getShapes();
        const ids = shapes.map(s => s.id);
        expect(ids).toContain(a.id);
        expect(ids).toContain(b.id);
    });
});

// ── partial connectors (Bug 17) ────────────────────────────────────────────────

describe('page.getConnectors() – partial connectors (Bug 17)', () => {
    it('includes a connector whose EndX endpoint is not connected (toShapeId is undefined)', async () => {
        const { doc, page, a, b } = await twoShapePage();
        await page.connectShapes(a, b);

        // Save, then strip the EndX <Connect> entry from the page XML
        const buf = await doc.save();
        const zip = await JSZip.loadAsync(buf);
        const pageFile = zip.file('visio/pages/page1.xml');
        let xml = await pageFile!.async('string');
        xml = xml.replace(/<Connect[^>]*FromCell="EndX"[^>]*>(\s*<\/Connect>)?/g, '');
        zip.file('visio/pages/page1.xml', xml);
        const buf2 = await zip.generateAsync({ type: 'nodebuffer' });

        const doc2 = await VisioDocument.load(buf2);
        const connectors = doc2.pages[0].getConnectors();
        expect(connectors).toHaveLength(1);
        expect(connectors[0].fromShapeId).toBe(a.id);
        expect(connectors[0].toShapeId).toBeUndefined();
    });

    it('includes a connector whose BeginX endpoint is not connected (fromShapeId is undefined)', async () => {
        const { doc, page, a, b } = await twoShapePage();
        await page.connectShapes(a, b);

        // Save, then strip the BeginX <Connect> entry from the page XML
        const buf = await doc.save();
        const zip = await JSZip.loadAsync(buf);
        const pageFile = zip.file('visio/pages/page1.xml');
        let xml = await pageFile!.async('string');
        xml = xml.replace(/<Connect[^>]*FromCell="BeginX"[^>]*>(\s*<\/Connect>)?/g, '');
        zip.file('visio/pages/page1.xml', xml);
        const buf2 = await zip.generateAsync({ type: 'nodebuffer' });

        const doc2 = await VisioDocument.load(buf2);
        const connectors = doc2.pages[0].getConnectors();
        expect(connectors).toHaveLength(1);
        expect(connectors[0].fromShapeId).toBeUndefined();
        expect(connectors[0].toShapeId).toBe(b.id);
    });

    it('returns a Connector instance for partial connectors', async () => {
        const { doc, page, a, b } = await twoShapePage();
        await page.connectShapes(a, b);

        const buf = await doc.save();
        const zip = await JSZip.loadAsync(buf);
        const pageFile = zip.file('visio/pages/page1.xml');
        let xml = await pageFile!.async('string');
        xml = xml.replace(/<Connect[^>]*FromCell="EndX"[^>]*>(\s*<\/Connect>)?/g, '');
        zip.file('visio/pages/page1.xml', xml);
        const buf2 = await zip.generateAsync({ type: 'nodebuffer' });

        const doc2 = await VisioDocument.load(buf2);
        const [conn] = doc2.pages[0].getConnectors();
        expect(conn).toBeInstanceOf(Connector);
        expect(conn.id).toBeTruthy();
    });
});

// ── round-trip: save & reload ──────────────────────────────────────────────────

describe('Connector round-trip save/reload', () => {
    it('connector survives save/reload with correct fromShapeId/toShapeId', async () => {
        const { doc, page, a, b } = await twoShapePage();
        await page.connectShapes(a, b, undefined, undefined, { lineColor: '#123456', routing: 'straight' });

        const buf  = await doc.save();
        const doc2 = await VisioDocument.load(buf);
        const pg2  = doc2.pages[0];
        const connectors = pg2.getConnectors();

        expect(connectors).toHaveLength(1);
        expect(connectors[0].fromShapeId).toBe(a.id);
        expect(connectors[0].toShapeId).toBe(b.id);
    });

    it('line style survives save/reload', async () => {
        const { doc, page, a, b } = await twoShapePage();
        await page.connectShapes(a, b, undefined, undefined, { lineColor: '#cc0000', lineWeight: 3 });

        const buf  = await doc.save();
        const doc2 = await VisioDocument.load(buf);
        const [conn] = doc2.pages[0].getConnectors();
        expect(conn.style.lineColor).toBe('#cc0000');
        expect(conn.style.lineWeight).toBeCloseTo(3, 1);
    });

    it('routing style survives save/reload', async () => {
        const { doc, page, a, b } = await twoShapePage();
        await page.connectShapes(a, b, undefined, undefined, { routing: 'curved' });

        const buf  = await doc.save();
        const doc2 = await VisioDocument.load(buf);
        const [conn] = doc2.pages[0].getConnectors();
        expect(conn.style.routing).toBe('curved');
    });

    it('arrows survive save/reload', async () => {
        const { doc, page, a, b } = await twoShapePage();
        await page.connectShapes(a, b, ArrowHeads.Standard, ArrowHeads.Standard);

        const buf  = await doc.save();
        const doc2 = await VisioDocument.load(buf);
        const [conn] = doc2.pages[0].getConnectors();
        expect(conn.beginArrow).toBe(ArrowHeads.Standard);
        expect(conn.endArrow).toBe(ArrowHeads.Standard);
    });

    it('multiple connectors survive save/reload', async () => {
        const doc  = await VisioDocument.create();
        const page = doc.pages[0];
        const a = await page.addShape({ text: 'A', x: 1, y: 1, width: 1, height: 1 });
        const b = await page.addShape({ text: 'B', x: 5, y: 1, width: 1, height: 1 });
        const c = await page.addShape({ text: 'C', x: 3, y: 4, width: 1, height: 1 });
        await page.connectShapes(a, b);
        await page.connectShapes(b, c);

        const buf  = await doc.save();
        const doc2 = await VisioDocument.load(buf);
        expect(doc2.pages[0].getConnectors()).toHaveLength(2);
    });
});
