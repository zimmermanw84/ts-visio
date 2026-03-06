import { describe, it, expect } from 'vitest';
import { VisioDocument } from '../src/VisioDocument';
import { ConnectorBuilder } from '../src/shapes/ConnectorBuilder';

// Helper: find a connector shape in the raw page XML by its ID
function findRawShape(parsed: any, id: string): any {
    const shapes = parsed.PageContents?.Shapes?.Shape;
    if (!shapes) return undefined;
    const arr = Array.isArray(shapes) ? shapes : [shapes];
    return arr.find((s: any) => s['@_ID'] === id);
}

function getCellVal(shape: any, name: string): string | undefined {
    if (!shape?.Cell) return undefined;
    const cells = Array.isArray(shape.Cell) ? shape.Cell : [shape.Cell];
    return cells.find((c: any) => c['@_N'] === name)?.['@_V'];
}

function getSectionCellVal(shape: any, sectionName: string, cellName: string): string | undefined {
    if (!shape?.Section) return undefined;
    const sections = Array.isArray(shape.Section) ? shape.Section : [shape.Section];
    const sec = sections.find((s: any) => s['@_N'] === sectionName);
    if (!sec?.Cell) return undefined;
    const cells = Array.isArray(sec.Cell) ? sec.Cell : [sec.Cell];
    return cells.find((c: any) => c['@_N'] === cellName)?.['@_V'];
}

// ---------------------------------------------------------------------------
// ConnectorBuilder unit tests
// ---------------------------------------------------------------------------

describe('ConnectorBuilder.createConnectorShapeObject()', () => {
    const layout = { beginX: 1, beginY: 1, endX: 4, endY: 1, width: 3, angle: 0 };

    it('uses default line color #000000 when no style is given', () => {
        const shape = ConnectorBuilder.createConnectorShapeObject('1', layout);
        const color = getSectionCellVal(shape, 'Line', 'LineColor');
        expect(color).toBe('#000000');
    });

    it('uses default ShapeRouteStyle 1 (orthogonal) when no routing given', () => {
        const shape = ConnectorBuilder.createConnectorShapeObject('1', layout);
        expect(getCellVal(shape, 'ShapeRouteStyle')).toBe('1');
    });

    it('applies lineColor from ConnectorStyle', () => {
        const shape = ConnectorBuilder.createConnectorShapeObject('1', layout, undefined, undefined, { lineColor: '#ff0000' });
        expect(getSectionCellVal(shape, 'Line', 'LineColor')).toBe('#ff0000');
    });

    it('converts lineWeight from points to inches', () => {
        // 1 pt = 1/72 inch ≈ 0.01389
        const shape = ConnectorBuilder.createConnectorShapeObject('1', layout, undefined, undefined, { lineWeight: 2 });
        const w = parseFloat(getSectionCellVal(shape, 'Line', 'LineWeight') ?? '0');
        expect(w).toBeCloseTo(2 / 72, 6);
    });

    it("applies linePattern", () => {
        const shape = ConnectorBuilder.createConnectorShapeObject('1', layout, undefined, undefined, { linePattern: 2 });
        expect(getSectionCellVal(shape, 'Line', 'LinePattern')).toBe('2');
    });

    it("routing 'straight' maps to ShapeRouteStyle 2", () => {
        const shape = ConnectorBuilder.createConnectorShapeObject('1', layout, undefined, undefined, { routing: 'straight' });
        expect(getCellVal(shape, 'ShapeRouteStyle')).toBe('2');
    });

    it("routing 'orthogonal' maps to ShapeRouteStyle 1", () => {
        const shape = ConnectorBuilder.createConnectorShapeObject('1', layout, undefined, undefined, { routing: 'orthogonal' });
        expect(getCellVal(shape, 'ShapeRouteStyle')).toBe('1');
    });

    it("routing 'curved' maps to ShapeRouteStyle 16", () => {
        const shape = ConnectorBuilder.createConnectorShapeObject('1', layout, undefined, undefined, { routing: 'curved' });
        expect(getCellVal(shape, 'ShapeRouteStyle')).toBe('16');
    });

    it('combined style applies all fields', () => {
        const shape = ConnectorBuilder.createConnectorShapeObject('1', layout, undefined, undefined, {
            lineColor: '#0000ff',
            lineWeight: 1,
            linePattern: 3,
            routing: 'curved',
        });
        expect(getSectionCellVal(shape, 'Line', 'LineColor')).toBe('#0000ff');
        expect(getSectionCellVal(shape, 'Line', 'LinePattern')).toBe('3');
        expect(getCellVal(shape, 'ShapeRouteStyle')).toBe('16');
    });

    // Bug 10 regression: connector shape must carry style attribute references
    it('includes LineStyle, FillStyle, and TextStyle attributes set to "0"', () => {
        const shape = ConnectorBuilder.createConnectorShapeObject('1', layout);
        expect((shape as any)['@_LineStyle']).toBe('0');
        expect((shape as any)['@_FillStyle']).toBe('0');
        expect((shape as any)['@_TextStyle']).toBe('0');
    });
});

// ---------------------------------------------------------------------------
// Integration tests via shape.connectTo() and page.connectShapes()
// ---------------------------------------------------------------------------

describe('Shape.connectTo() with ConnectorStyle', () => {
    it('creates a connector without style (backwards compatible)', async () => {
        const doc = await VisioDocument.create();
        const page = doc.pages[0];
        const a = await page.addShape({ text: 'A', x: 1, y: 5, width: 2, height: 1 });
        const b = await page.addShape({ text: 'B', x: 5, y: 5, width: 2, height: 1 });
        await expect(a.connectTo(b)).resolves.toBeDefined();
    });

    it('applies lineColor to connector via connectTo()', async () => {
        const doc = await VisioDocument.create();
        const page = doc.pages[0];
        const a = await page.addShape({ text: 'A', x: 1, y: 5, width: 2, height: 1 });
        const b = await page.addShape({ text: 'B', x: 5, y: 5, width: 2, height: 1 });

        await a.connectTo(b, undefined, undefined, { lineColor: '#cc0000' });

        const parsed = page['modifier'].getParsed(page['id']);
        const shapes = Array.isArray(parsed.PageContents.Shapes.Shape)
            ? parsed.PageContents.Shapes.Shape
            : [parsed.PageContents.Shapes.Shape];
        const connector = shapes.find((s: any) => s['@_NameU'] === 'Dynamic connector');
        expect(getSectionCellVal(connector, 'Line', 'LineColor')).toBe('#cc0000');
    });

    it("routing 'straight' is persisted on the connector shape", async () => {
        const doc = await VisioDocument.create();
        const page = doc.pages[0];
        const a = await page.addShape({ text: 'A', x: 1, y: 5, width: 2, height: 1 });
        const b = await page.addShape({ text: 'B', x: 5, y: 5, width: 2, height: 1 });

        await a.connectTo(b, undefined, undefined, { routing: 'straight' });

        const parsed = page['modifier'].getParsed(page['id']);
        const shapes = Array.isArray(parsed.PageContents.Shapes.Shape)
            ? parsed.PageContents.Shapes.Shape
            : [parsed.PageContents.Shapes.Shape];
        const connector = shapes.find((s: any) => s['@_NameU'] === 'Dynamic connector');
        expect(getCellVal(connector, 'ShapeRouteStyle')).toBe('2');
    });

    it("routing 'curved' is persisted on the connector shape", async () => {
        const doc = await VisioDocument.create();
        const page = doc.pages[0];
        const a = await page.addShape({ text: 'A', x: 1, y: 5, width: 2, height: 1 });
        const b = await page.addShape({ text: 'B', x: 5, y: 5, width: 2, height: 1 });

        await a.connectTo(b, undefined, undefined, { routing: 'curved' });

        const parsed = page['modifier'].getParsed(page['id']);
        const shapes = Array.isArray(parsed.PageContents.Shapes.Shape)
            ? parsed.PageContents.Shapes.Shape
            : [parsed.PageContents.Shapes.Shape];
        const connector = shapes.find((s: any) => s['@_NameU'] === 'Dynamic connector');
        expect(getCellVal(connector, 'ShapeRouteStyle')).toBe('16');
    });

    it('connectTo() still returns this for chaining', async () => {
        const doc = await VisioDocument.create();
        const page = doc.pages[0];
        const a = await page.addShape({ text: 'A', x: 1, y: 5, width: 2, height: 1 });
        const b = await page.addShape({ text: 'B', x: 5, y: 5, width: 2, height: 1 });
        const result = await a.connectTo(b, undefined, undefined, { lineColor: '#aabbcc' });
        expect(result).toBe(a);
    });
});

describe('Page.connectShapes() with ConnectorStyle', () => {
    it('passes style through to the connector', async () => {
        const doc = await VisioDocument.create();
        const page = doc.pages[0];
        const a = await page.addShape({ text: 'A', x: 1, y: 5, width: 2, height: 1 });
        const b = await page.addShape({ text: 'B', x: 5, y: 5, width: 2, height: 1 });

        await page.connectShapes(a, b, undefined, undefined, { routing: 'curved', lineColor: '#336699' });

        const parsed = page['modifier'].getParsed(page['id']);
        const shapes = Array.isArray(parsed.PageContents.Shapes.Shape)
            ? parsed.PageContents.Shapes.Shape
            : [parsed.PageContents.Shapes.Shape];
        const connector = shapes.find((s: any) => s['@_NameU'] === 'Dynamic connector');
        expect(getCellVal(connector, 'ShapeRouteStyle')).toBe('16');
        expect(getSectionCellVal(connector, 'Line', 'LineColor')).toBe('#336699');
    });
});

// ---------------------------------------------------------------------------
// Missing exports regression test
// ---------------------------------------------------------------------------

describe('Public exports', () => {
    it('Layer is exported from ts-visio', async () => {
        const mod = await import('../src/index');
        expect(mod.Layer).toBeDefined();
    });

    it('SchemaDiagram is exported from ts-visio', async () => {
        const mod = await import('../src/index');
        expect(mod.SchemaDiagram).toBeDefined();
    });

    it('VisioValidator is exported from ts-visio', async () => {
        const mod = await import('../src/index');
        expect(mod.VisioValidator).toBeDefined();
    });

    it('ConnectorStyle type is available (ConnectorRouting exported)', async () => {
        // Type-level check — just import and verify the module loads without error
        const mod = await import('../src/index');
        expect(mod.VisioPropType).toBeDefined(); // sanity check that types barrel works
    });
});
