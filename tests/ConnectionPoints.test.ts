import { describe, it, expect, beforeEach } from 'vitest';
import { VisioDocument } from '../src/VisioDocument';
import { ConnectionPointBuilder } from '../src/shapes/ConnectionPointBuilder';
import { ConnectorBuilder } from '../src/shapes/ConnectorBuilder';
import {
    ConnectionPointDef,
    ConnectionTarget,
    StandardConnectionPoints,
} from '../src/types/VisioTypes';

// ── Helpers ────────────────────────────────────────────────────────────────────

async function createDoc() {
    const doc = await VisioDocument.create();
    const page = await doc.addPage('Test');
    return { doc, page };
}

function makeShape(
    id: string,
    pinX: number,
    pinY: number,
    width: number,
    height: number,
    connectionPoints?: ConnectionPointDef[],
): any {
    const shape: any = {
        '@_ID': id,
        Cell: [
            { '@_N': 'PinX',    '@_V': pinX.toString()        },
            { '@_N': 'PinY',    '@_V': pinY.toString()        },
            { '@_N': 'Width',   '@_V': width.toString()       },
            { '@_N': 'Height',  '@_V': height.toString()      },
            { '@_N': 'LocPinX', '@_V': (width  / 2).toString() },
            { '@_N': 'LocPinY', '@_V': (height / 2).toString() },
        ],
    };
    if (connectionPoints && connectionPoints.length > 0) {
        shape.Section = [ConnectionPointBuilder.buildConnectionSection(connectionPoints)];
    }
    return shape;
}

// ── ConnectionPointBuilder unit tests ─────────────────────────────────────────

describe('ConnectionPointBuilder', () => {
    describe('buildConnectionSection', () => {
        it('creates a Section with N="Connection"', () => {
            const section = ConnectionPointBuilder.buildConnectionSection([
                { xFraction: 0.5, yFraction: 1.0 },
            ]);
            expect(section['@_N']).toBe('Connection');
        });

        it('creates one Row per point with correct IX', () => {
            const section = ConnectionPointBuilder.buildConnectionSection([
                { xFraction: 0.5, yFraction: 1.0 },
                { xFraction: 1.0, yFraction: 0.5 },
            ]);
            expect(section.Row).toHaveLength(2);
            expect(section.Row[0]['@_IX']).toBe('0');
            expect(section.Row[1]['@_IX']).toBe('1');
        });

        it('sets Row @_N when name is provided', () => {
            const section = ConnectionPointBuilder.buildConnectionSection([
                { name: 'Top', xFraction: 0.5, yFraction: 1.0 },
            ]);
            expect(section.Row[0]['@_N']).toBe('Top');
        });

        it('omits Row @_N when no name is provided', () => {
            const section = ConnectionPointBuilder.buildConnectionSection([
                { xFraction: 0.5, yFraction: 0.5 },
            ]);
            expect(section.Row[0]['@_N']).toBeUndefined();
        });

        it('stores X formula as Width*fraction', () => {
            const section = ConnectionPointBuilder.buildConnectionSection([
                { xFraction: 0.5, yFraction: 1.0 },
            ]);
            const cells: any[] = section.Row[0].Cell;
            const xCell = cells.find((c: any) => c['@_N'] === 'X');
            expect(xCell['@_F']).toBe('Width*0.5');
        });

        it('stores Y formula as Height*fraction', () => {
            const section = ConnectionPointBuilder.buildConnectionSection([
                { xFraction: 0.5, yFraction: 1.0 },
            ]);
            const cells: any[] = section.Row[0].Cell;
            const yCell = cells.find((c: any) => c['@_N'] === 'Y');
            expect(yCell['@_F']).toBe('Height*1');
        });

        it('stores DirX and DirY when direction provided', () => {
            const section = ConnectionPointBuilder.buildConnectionSection([
                { xFraction: 0.5, yFraction: 1.0, direction: { x: 0, y: 1 } },
            ]);
            const cells: any[] = section.Row[0].Cell;
            expect(cells.find((c: any) => c['@_N'] === 'DirX')['@_V']).toBe('0');
            expect(cells.find((c: any) => c['@_N'] === 'DirY')['@_V']).toBe('1');
        });

        it('stores Type=1 for outward connection', () => {
            const section = ConnectionPointBuilder.buildConnectionSection([
                { xFraction: 0.5, yFraction: 0.5, type: 'outward' },
            ]);
            const cells: any[] = section.Row[0].Cell;
            expect(cells.find((c: any) => c['@_N'] === 'Type')['@_V']).toBe('1');
        });

        it('stores Type=2 for "both" connection', () => {
            const section = ConnectionPointBuilder.buildConnectionSection([
                { xFraction: 0.5, yFraction: 0.5, type: 'both' },
            ]);
            const cells: any[] = section.Row[0].Cell;
            expect(cells.find((c: any) => c['@_N'] === 'Type')['@_V']).toBe('2');
        });

        it('adds Prompt cell when prompt is provided', () => {
            const section = ConnectionPointBuilder.buildConnectionSection([
                { xFraction: 0.5, yFraction: 0.5, prompt: 'Connect here' },
            ]);
            const cells: any[] = section.Row[0].Cell;
            const promptCell = cells.find((c: any) => c['@_N'] === 'Prompt');
            expect(promptCell).toBeDefined();
            expect(promptCell['@_V']).toBe('Connect here');
        });

        // Bug 4 regression: @_V must reflect actual coordinates, not always '0'
        it('Bug 4: X @_V equals width * xFraction when dimensions supplied', () => {
            const section = ConnectionPointBuilder.buildConnectionSection(
                [{ xFraction: 0.5, yFraction: 1.0 }],
                4, 2,
            );
            const cells: any[] = section.Row[0].Cell;
            const xCell = cells.find((c: any) => c['@_N'] === 'X');
            expect(xCell['@_V']).toBe('2'); // 4 * 0.5
        });

        it('Bug 4: Y @_V equals height * yFraction when dimensions supplied', () => {
            const section = ConnectionPointBuilder.buildConnectionSection(
                [{ xFraction: 0.5, yFraction: 1.0 }],
                4, 2,
            );
            const cells: any[] = section.Row[0].Cell;
            const yCell = cells.find((c: any) => c['@_N'] === 'Y');
            expect(yCell['@_V']).toBe('2'); // 2 * 1.0
        });
    });

    describe('resolveTarget', () => {
        const shape = makeShape('1', 2, 2, 2, 1, [
            { name: 'Top',   xFraction: 0.5, yFraction: 1.0 },
            { name: 'Right', xFraction: 1.0, yFraction: 0.5 },
        ]);

        it('"center" returns ToPart=3 and ToCell=PinX', () => {
            const r = ConnectionPointBuilder.resolveTarget('center', shape);
            expect(r.toPart).toBe('3');
            expect(r.toCell).toBe('PinX');
            expect(r.xFraction).toBeUndefined();
            expect(r.yFraction).toBeUndefined();
        });

        it('{ name } resolves to correct ToPart/ToCell', () => {
            const r = ConnectionPointBuilder.resolveTarget({ name: 'Top' }, shape);
            expect(r.toPart).toBe('100'); // 100 + IX=0
            expect(r.toCell).toBe('Connections.X1');
            expect(r.xFraction).toBeCloseTo(0.5);
            expect(r.yFraction).toBeCloseTo(1.0);
        });

        it('{ name } at IX=1 returns ToPart=101', () => {
            const r = ConnectionPointBuilder.resolveTarget({ name: 'Right' }, shape);
            expect(r.toPart).toBe('101');
            expect(r.toCell).toBe('Connections.X2');
            expect(r.xFraction).toBeCloseTo(1.0);
            expect(r.yFraction).toBeCloseTo(0.5);
        });

        it('{ index: 0 } resolves to correct ToPart/ToCell', () => {
            const r = ConnectionPointBuilder.resolveTarget({ index: 0 }, shape);
            expect(r.toPart).toBe('100');
            expect(r.toCell).toBe('Connections.X1');
        });

        it('{ index: 1 } resolves to ToPart=101', () => {
            const r = ConnectionPointBuilder.resolveTarget({ index: 1 }, shape);
            expect(r.toPart).toBe('101');
            expect(r.toCell).toBe('Connections.X2');
        });

        it('unknown name falls back to centre', () => {
            const r = ConnectionPointBuilder.resolveTarget({ name: 'NoSuchPoint' }, shape);
            expect(r.toPart).toBe('3');
            expect(r.toCell).toBe('PinX');
        });

        it('shape with no Connection section falls back to centre for name lookup', () => {
            const plain = makeShape('99', 1, 1, 2, 2);
            const r = ConnectionPointBuilder.resolveTarget({ name: 'Top' }, plain);
            expect(r.toPart).toBe('3');
            expect(r.toCell).toBe('PinX');
        });
    });
});

// ── StandardConnectionPoints preset ───────────────────────────────────────────

describe('StandardConnectionPoints', () => {
    it('cardinal has exactly 4 points', () => {
        expect(StandardConnectionPoints.cardinal).toHaveLength(4);
    });

    it('full has exactly 8 points', () => {
        expect(StandardConnectionPoints.full).toHaveLength(8);
    });

    it('cardinal contains Top, Right, Bottom, Left', () => {
        const names = StandardConnectionPoints.cardinal.map(p => p.name);
        expect(names).toContain('Top');
        expect(names).toContain('Right');
        expect(names).toContain('Bottom');
        expect(names).toContain('Left');
    });

    it('full contains all cardinal + corner names', () => {
        const names = StandardConnectionPoints.full.map(p => p.name);
        ['Top', 'Right', 'Bottom', 'Left', 'TopLeft', 'TopRight', 'BottomRight', 'BottomLeft']
            .forEach(n => expect(names).toContain(n));
    });
});

// ── ConnectorBuilder layout with ports ────────────────────────────────────────

describe('ConnectorBuilder.calculateConnectorLayout with ports', () => {
    // Shape A: centre (1, 1), 2×1.  Left edge x=0, right edge x=2, bottom y=0.5, top y=1.5
    // Shape B: centre (5, 1), 2×1.
    function buildHierarchy(
        aPoints?: ConnectionPointDef[],
        bPoints?: ConnectionPointDef[],
    ) {
        const shapeA = makeShape('A', 1, 1, 2, 1, aPoints);
        const shapeB = makeShape('B', 5, 1, 2, 1, bPoints);
        const hierarchy = new Map<string, { shape: any; parent: any }>();
        hierarchy.set('A', { shape: shapeA, parent: null });
        hierarchy.set('B', { shape: shapeB, parent: null });
        return hierarchy;
    }

    it('no ports — uses edge-intersection logic (existing behaviour)', () => {
        const h = buildHierarchy();
        const l = ConnectorBuilder.calculateConnectorLayout('A', 'B', h);
        // A is to the left of B, so BeginX should be ~2 (right edge of A)
        expect(l.beginX).toBeCloseTo(2);
        expect(l.endX).toBeCloseTo(4);   // left edge of B
    });

    it('fromPort "Top" → beginY at top edge of A', () => {
        const h = buildHierarchy(StandardConnectionPoints.cardinal);
        const l = ConnectorBuilder.calculateConnectorLayout('A', 'B', h, { name: 'Top' });
        // Top of A: x = 1 + 2*(0.5-0.5) = 1, y = 1 + 1*(1.0-0.5) = 1.5
        expect(l.beginX).toBeCloseTo(1);
        expect(l.beginY).toBeCloseTo(1.5);
    });

    it('toPort "Left" → endX at left edge of B', () => {
        const h = buildHierarchy(undefined, StandardConnectionPoints.cardinal);
        const l = ConnectorBuilder.calculateConnectorLayout('A', 'B', h, undefined, { name: 'Left' });
        // Left of B: x = 5 + 2*(0-0.5) = 5 - 1 = 4, y = 1 + 1*(0.5-0.5) = 1
        expect(l.endX).toBeCloseTo(4);
        expect(l.endY).toBeCloseTo(1);
    });

    it('{ index: 1 } resolves Right point of A', () => {
        // A has cardinal points; Right is IX=1 (xFraction=1.0, yFraction=0.5)
        const h = buildHierarchy(StandardConnectionPoints.cardinal);
        const l = ConnectorBuilder.calculateConnectorLayout('A', 'B', h, { index: 1 });
        // Right of A: x = 1 + 2*(1.0-0.5) = 2, y = 1 + 1*(0.5-0.5) = 1
        expect(l.beginX).toBeCloseTo(2);
        expect(l.beginY).toBeCloseTo(1);
    });

    it('unknown port name falls back to edge-intersection', () => {
        const h = buildHierarchy(StandardConnectionPoints.cardinal);
        const l = ConnectorBuilder.calculateConnectorLayout('A', 'B', h, { name: 'NonExistent' });
        // Falls back to edge point — same as no-port result
        expect(l.beginX).toBeCloseTo(2);
    });

    it('"center" fromPort uses edge-intersection, not literal centre', () => {
        const h = buildHierarchy(StandardConnectionPoints.cardinal);
        const lNoPort  = ConnectorBuilder.calculateConnectorLayout('A', 'B', h);
        const lCenter  = ConnectorBuilder.calculateConnectorLayout('A', 'B', h, 'center');
        expect(lCenter.beginX).toBeCloseTo(lNoPort.beginX);
        expect(lCenter.beginY).toBeCloseTo(lNoPort.beginY);
    });
});

// ── ConnectorBuilder.addConnectorToConnects with ports ────────────────────────

describe('ConnectorBuilder.addConnectorToConnects with ports', () => {
    function buildParsedWithShapes(aPoints?: ConnectionPointDef[]) {
        const shapeA = makeShape('1', 2, 2, 2, 1, aPoints);
        const shapeB = makeShape('2', 5, 2, 2, 1);
        const parsed = {
            PageContents: {
                Shapes: { Shape: [shapeA, shapeB] },
            },
        };
        const hierarchy = new Map<string, { shape: any; parent: any }>();
        hierarchy.set('1', { shape: shapeA, parent: null });
        hierarchy.set('2', { shape: shapeB, parent: null });
        return { parsed, hierarchy };
    }

    it('default (no ports) — ToPart=3 for both endpoints', () => {
        const { parsed, hierarchy } = buildParsedWithShapes();
        ConnectorBuilder.addConnectorToConnects(parsed, '99', '1', '2', hierarchy);
        const connects = parsed.PageContents.Connects.Connect;
        expect(connects[0]['@_ToPart']).toBe('3');
        expect(connects[1]['@_ToPart']).toBe('3');
    });

    it('fromPort by name → ToPart=100 for Begin connect', () => {
        const { parsed, hierarchy } = buildParsedWithShapes(StandardConnectionPoints.cardinal);
        ConnectorBuilder.addConnectorToConnects(parsed, '99', '1', '2', hierarchy, { name: 'Top' });
        const connects = parsed.PageContents.Connects.Connect;
        const begin = connects.find((c: any) => c['@_FromCell'] === 'BeginX');
        expect(begin['@_ToPart']).toBe('100');
        expect(begin['@_ToCell']).toBe('Connections.X1');
    });

    it('toPort by index → ToPart=100 for End connect', () => {
        const { parsed, hierarchy } = buildParsedWithShapes();
        // Give shape B connection points (Right at IX=1)
        const shapeB = hierarchy.get('2')!.shape;
        shapeB.Section = [ConnectionPointBuilder.buildConnectionSection(StandardConnectionPoints.cardinal)];
        ConnectorBuilder.addConnectorToConnects(parsed, '99', '1', '2', hierarchy, undefined, { index: 0 });
        const connects = parsed.PageContents.Connects.Connect;
        const end = connects.find((c: any) => c['@_FromCell'] === 'EndX');
        expect(end['@_ToPart']).toBe('100');
        expect(end['@_ToCell']).toBe('Connections.X1');
    });
});

// ── Integration: adding shapes with connectionPoints ──────────────────────────

describe('Shape with connectionPoints integration', () => {
    it('connectionPoints rendered in shape XML via addShape', async () => {
        const { page } = await createDoc();
        const shape = await page.addShape({
            text: 'Node',
            x: 3, y: 3, width: 2, height: 1,
            connectionPoints: StandardConnectionPoints.cardinal,
        });
        expect(shape).toBeDefined();
        expect(shape.id).toBeTruthy();
    });

    it('addConnectionPoint adds a point to an existing shape', async () => {
        const { page } = await createDoc();
        const shape = await page.addShape({
            text: 'A', x: 2, y: 2, width: 2, height: 1,
        });
        const ix = shape.addConnectionPoint({ name: 'Top', xFraction: 0.5, yFraction: 1.0 });
        expect(ix).toBe(0);
    });

    it('addConnectionPoint returns sequential IX for multiple points', async () => {
        const { page } = await createDoc();
        const shape = await page.addShape({
            text: 'B', x: 2, y: 2, width: 2, height: 1,
        });
        const ix0 = shape.addConnectionPoint({ name: 'Top',   xFraction: 0.5, yFraction: 1.0 });
        const ix1 = shape.addConnectionPoint({ name: 'Right', xFraction: 1.0, yFraction: 0.5 });
        expect(ix0).toBe(0);
        expect(ix1).toBe(1);
    });

    // Bug 4 regression: @_V must not be '0' after addConnectionPoint on a 4×2 shape
    it('Bug 4: buildRow V values match shape dimensions', () => {
        const row = ConnectionPointBuilder.buildRow({ name: 'Top', xFraction: 0.5, yFraction: 1.0 }, 0, 4, 2);
        const cells: any[] = row.Cell;
        const xCell = cells.find((c: any) => c['@_N'] === 'X');
        const yCell = cells.find((c: any) => c['@_N'] === 'Y');
        expect(xCell['@_V']).toBe('2');   // 4 * 0.5
        expect(yCell['@_V']).toBe('2');   // 2 * 1.0
        expect(xCell['@_F']).toBe('Width*0.5');
        expect(yCell['@_F']).toBe('Height*1');
    });
});

// ── Integration: connecting via named ports ────────────────────────────────────

describe('connectShapes with named ports', () => {
    it('connects two shapes using named connection points', async () => {
        const { page } = await createDoc();
        const a = await page.addShape({
            text: 'A', x: 2, y: 2, width: 2, height: 1,
            connectionPoints: StandardConnectionPoints.cardinal,
        });
        const b = await page.addShape({
            text: 'B', x: 6, y: 2, width: 2, height: 1,
            connectionPoints: StandardConnectionPoints.cardinal,
        });
        // Should not throw
        await expect(
            page.connectShapes(a, b, undefined, undefined, undefined, { name: 'Right' }, { name: 'Left' }),
        ).resolves.toBeUndefined();
    });

    it('connectTo on Shape API forwards ports', async () => {
        const { page } = await createDoc();
        const a = await page.addShape({
            text: 'A', x: 2, y: 2, width: 2, height: 1,
            connectionPoints: StandardConnectionPoints.cardinal,
        });
        const b = await page.addShape({
            text: 'B', x: 6, y: 2, width: 2, height: 1,
            connectionPoints: StandardConnectionPoints.cardinal,
        });
        await expect(
            a.connectTo(b, undefined, undefined, undefined, { name: 'Right' }, { name: 'Left' }),
        ).resolves.toBeDefined();
    });

    it('falls back gracefully when port name does not exist', async () => {
        const { page } = await createDoc();
        const a = await page.addShape({ text: 'A', x: 2, y: 2, width: 2, height: 1 });
        const b = await page.addShape({ text: 'B', x: 6, y: 2, width: 2, height: 1 });
        // No connection points defined — should fall back to edge-intersection, not throw
        await expect(
            page.connectShapes(a, b, undefined, undefined, undefined, { name: 'NonExistent' }),
        ).resolves.toBeUndefined();
    });

    it('center target behaves same as no port', async () => {
        const { page } = await createDoc();
        const a = await page.addShape({ text: 'A', x: 2, y: 2, width: 2, height: 1 });
        const b = await page.addShape({ text: 'B', x: 6, y: 2, width: 2, height: 1 });
        await expect(
            page.connectShapes(a, b, undefined, undefined, undefined, 'center', 'center'),
        ).resolves.toBeUndefined();
    });
});
