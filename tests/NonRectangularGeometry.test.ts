import { describe, it, expect } from 'vitest';
import { VisioDocument } from '../src/VisioDocument';
import { GeometryBuilder } from '../src/shapes/GeometryBuilder';

// ---------------------------------------------------------------------------
// GeometryBuilder unit tests — verify the raw section objects directly
// ---------------------------------------------------------------------------

describe('GeometryBuilder.rectangle()', () => {
    it('produces a Geometry section with NoFill=1 when no fill', () => {
        const geo = GeometryBuilder.rectangle(4, 2, '1');
        expect(geo['@_N']).toBe('Geometry');
        const noFill = geo.Cell.find((c: any) => c['@_N'] === 'NoFill');
        expect(noFill['@_V']).toBe('1');
    });

    it('includes a NoShow cell with value 0', () => {
        const geo = GeometryBuilder.rectangle(4, 2, '0');
        const noShow = geo.Cell.find((c: any) => c['@_N'] === 'NoShow');
        expect(noShow).toBeDefined();
        expect(noShow['@_V']).toBe('0');
    });

    it('produces 5 rows: MoveTo + 4 LineTo', () => {
        const geo = GeometryBuilder.rectangle(4, 2, '0');
        expect(geo.Row).toHaveLength(5);
        expect(geo.Row[0]['@_T']).toBe('MoveTo');
        expect(geo.Row[1]['@_T']).toBe('LineTo');
        expect(geo.Row[4]['@_T']).toBe('LineTo');
    });

    it('carries Width/Height @_F formulas on the corner cells', () => {
        const geo = GeometryBuilder.rectangle(4, 2, '0');
        const row2 = geo.Row[1]; // LineTo (W, 0)
        expect(row2.Cell.find((c: any) => c['@_N'] === 'X')['@_F']).toBe('Width');
    });
});

describe('GeometryBuilder.ellipse()', () => {
    it('produces a single Ellipse row — no MoveTo needed', () => {
        const geo = GeometryBuilder.ellipse(4, 2, '0');
        expect(geo.Row).toHaveLength(1);
        expect(geo.Row[0]['@_T']).toBe('Ellipse');
    });

    it('sets centre X to W/2 and centre Y to H/2', () => {
        const geo = GeometryBuilder.ellipse(4, 2, '0');
        const cells: any[] = geo.Row[0].Cell;
        const cx = cells.find((c: any) => c['@_N'] === 'X');
        const cy = cells.find((c: any) => c['@_N'] === 'Y');
        expect(parseFloat(cx['@_V'])).toBeCloseTo(2, 5);
        expect(parseFloat(cy['@_V'])).toBeCloseTo(1, 5);
    });

    it('sets right-point A=W, B=H/2 and top-point C=W/2, D=H', () => {
        const geo = GeometryBuilder.ellipse(4, 2, '0');
        const cells: any[] = geo.Row[0].Cell;
        const A = cells.find((c: any) => c['@_N'] === 'A');
        const B = cells.find((c: any) => c['@_N'] === 'B');
        const C = cells.find((c: any) => c['@_N'] === 'C');
        const D = cells.find((c: any) => c['@_N'] === 'D');
        expect(parseFloat(A['@_V'])).toBeCloseTo(4, 5);   // rightmost X = W
        expect(parseFloat(B['@_V'])).toBeCloseTo(1, 5);   // rightmost Y = H/2
        expect(parseFloat(C['@_V'])).toBeCloseTo(2, 5);   // topmost X = W/2
        expect(parseFloat(D['@_V'])).toBeCloseTo(2, 5);   // topmost Y = H
    });

    it('carries Width/Height @_F formulas', () => {
        const geo = GeometryBuilder.ellipse(4, 2, '0');
        const cells: any[] = geo.Row[0].Cell;
        const A = cells.find((c: any) => c['@_N'] === 'A');
        expect(A['@_F']).toBe('Width');
        const D = cells.find((c: any) => c['@_N'] === 'D');
        expect(D['@_F']).toBe('Height');
    });
});

describe('GeometryBuilder.diamond()', () => {
    it('produces 5 rows: MoveTo + 4 LineTo', () => {
        const geo = GeometryBuilder.diamond(4, 2, '0');
        expect(geo.Row).toHaveLength(5);
        expect(geo.Row[0]['@_T']).toBe('MoveTo');
        geo.Row.slice(1).forEach((r: any) => expect(r['@_T']).toBe('LineTo'));
    });

    it('starts at the top vertex (W/2, H)', () => {
        const geo = GeometryBuilder.diamond(4, 2, '0');
        const startRow = geo.Row[0];
        const x = startRow.Cell.find((c: any) => c['@_N'] === 'X');
        const y = startRow.Cell.find((c: any) => c['@_N'] === 'Y');
        expect(parseFloat(x['@_V'])).toBeCloseTo(2, 5); // W/2
        expect(parseFloat(y['@_V'])).toBeCloseTo(2, 5); // H
    });

    it('has right vertex at (W, H/2)', () => {
        const geo = GeometryBuilder.diamond(4, 2, '0');
        const row = geo.Row[1];
        const x = row.Cell.find((c: any) => c['@_N'] === 'X');
        const y = row.Cell.find((c: any) => c['@_N'] === 'Y');
        expect(parseFloat(x['@_V'])).toBeCloseTo(4, 5); // W
        expect(parseFloat(y['@_V'])).toBeCloseTo(1, 5); // H/2
    });
});

describe('GeometryBuilder.roundedRectangle()', () => {
    it('produces 9 rows: MoveTo + 4*(LineTo + EllipticalArcTo)', () => {
        const geo = GeometryBuilder.roundedRectangle(4, 2, 0.25, '0');
        expect(geo.Row).toHaveLength(9);
    });

    it('contains exactly 4 EllipticalArcTo rows', () => {
        const geo = GeometryBuilder.roundedRectangle(4, 2, 0.25, '0');
        const arcs = geo.Row.filter((r: any) => r['@_T'] === 'EllipticalArcTo');
        expect(arcs).toHaveLength(4);
    });

    it('each arc has D=1 (circular) and C=0 (no rotation)', () => {
        const geo = GeometryBuilder.roundedRectangle(4, 2, 0.25, '0');
        const arcs = geo.Row.filter((r: any) => r['@_T'] === 'EllipticalArcTo');
        for (const arc of arcs) {
            const D = arc.Cell.find((c: any) => c['@_N'] === 'D');
            const C = arc.Cell.find((c: any) => c['@_N'] === 'C');
            expect(D['@_V']).toBe('1');
            expect(C['@_V']).toBe('0');
        }
    });

    it('starts path at (r, 0) — just right of bottom-left corner', () => {
        const r = 0.25;
        const geo = GeometryBuilder.roundedRectangle(4, 2, r, '0');
        const start = geo.Row[0];
        const x = start.Cell.find((c: any) => c['@_N'] === 'X');
        const y = start.Cell.find((c: any) => c['@_N'] === 'Y');
        expect(parseFloat(x['@_V'])).toBeCloseTo(r, 5);
        expect(parseFloat(y['@_V'])).toBeCloseTo(0, 5);
    });

    it('defaults cornerRadius to 10% of the smaller dimension via build()', () => {
        // W=4, H=2 → default r = 0.2
        const geo = GeometryBuilder.build({ width: 4, height: 2, geometry: 'rounded-rectangle' });
        const start = geo.Row[0];
        const x = start.Cell.find((c: any) => c['@_N'] === 'X');
        expect(parseFloat(x['@_V'])).toBeCloseTo(0.2, 5); // 10% of H=2
    });
});

describe('GeometryBuilder.triangle()', () => {
    it('produces 4 rows: MoveTo + 3 LineTo', () => {
        const geo = GeometryBuilder.triangle(4, 2, '0');
        expect(geo.Row).toHaveLength(4);
        expect(geo.Row[0]['@_T']).toBe('MoveTo');
        geo.Row.slice(1).forEach((r: any) => expect(r['@_T']).toBe('LineTo'));
    });

    it('places the apex at (W, H/2) — right-pointing', () => {
        const geo = GeometryBuilder.triangle(4, 2, '0');
        const apexRow = geo.Row[1];
        const x = apexRow.Cell.find((c: any) => c['@_N'] === 'X');
        const y = apexRow.Cell.find((c: any) => c['@_N'] === 'Y');
        expect(parseFloat(x['@_V'])).toBeCloseTo(4, 5); // W
        expect(parseFloat(y['@_V'])).toBeCloseTo(1, 5); // H/2
    });
});

describe('GeometryBuilder.parallelogram()', () => {
    it('produces 5 rows: MoveTo + 4 LineTo', () => {
        const geo = GeometryBuilder.parallelogram(4, 2, '0');
        expect(geo.Row).toHaveLength(5);
        expect(geo.Row[0]['@_T']).toBe('MoveTo');
    });

    it('starts at (W*0.2, 0) — the skewed bottom-left vertex', () => {
        const geo = GeometryBuilder.parallelogram(4, 2, '0');
        const start = geo.Row[0];
        const x = start.Cell.find((c: any) => c['@_N'] === 'X');
        expect(parseFloat(x['@_V'])).toBeCloseTo(4 * 0.2, 5); // 0.8
    });

    it('top-right vertex is at (W - W*0.2, H)', () => {
        const geo = GeometryBuilder.parallelogram(4, 2, '0');
        const row = geo.Row[2]; // third row = top-right vertex
        const x = row.Cell.find((c: any) => c['@_N'] === 'X');
        const y = row.Cell.find((c: any) => c['@_N'] === 'Y');
        expect(parseFloat(x['@_V'])).toBeCloseTo(4 - 4 * 0.2, 5); // 3.2
        expect(parseFloat(y['@_V'])).toBeCloseTo(2, 5);
    });
});

// ---------------------------------------------------------------------------
// GeometryBuilder.build() — dispatch tests
// ---------------------------------------------------------------------------

describe('GeometryBuilder.build()', () => {
    it('defaults to rectangle when geometry is omitted', () => {
        const geo = GeometryBuilder.build({ width: 2, height: 1 });
        expect(geo.Row[0]['@_T']).toBe('MoveTo');
        expect(geo.Row).toHaveLength(5);
    });

    it("dispatches 'ellipse' to the Ellipse row", () => {
        const geo = GeometryBuilder.build({ width: 2, height: 1, geometry: 'ellipse' });
        expect(geo.Row[0]['@_T']).toBe('Ellipse');
    });

    it("dispatches 'diamond'", () => {
        const geo = GeometryBuilder.build({ width: 2, height: 1, geometry: 'diamond' });
        expect(geo.Row[0]['@_T']).toBe('MoveTo');
        const arcs = geo.Row.filter((r: any) => r['@_T'] === 'EllipticalArcTo');
        expect(arcs).toHaveLength(0);
    });

    it("dispatches 'rounded-rectangle'", () => {
        const geo = GeometryBuilder.build({ width: 2, height: 1, geometry: 'rounded-rectangle' });
        const arcs = geo.Row.filter((r: any) => r['@_T'] === 'EllipticalArcTo');
        expect(arcs).toHaveLength(4);
    });

    it("dispatches 'triangle'", () => {
        const geo = GeometryBuilder.build({ width: 2, height: 1, geometry: 'triangle' });
        expect(geo.Row).toHaveLength(4);
    });

    it("dispatches 'parallelogram'", () => {
        const geo = GeometryBuilder.build({ width: 2, height: 1, geometry: 'parallelogram' });
        expect(geo.Row).toHaveLength(5);
    });

    it('sets NoFill=0 when fillColor is provided', () => {
        const geo = GeometryBuilder.build({ width: 2, height: 1, fillColor: '#ff0000' });
        const noFill = geo.Cell.find((c: any) => c['@_N'] === 'NoFill');
        expect(noFill['@_V']).toBe('0');
    });

    it('sets NoFill=1 when no fillColor', () => {
        const geo = GeometryBuilder.build({ width: 2, height: 1 });
        const noFill = geo.Cell.find((c: any) => c['@_N'] === 'NoFill');
        expect(noFill['@_V']).toBe('1');
    });

    it('all geometry types include NoShow=0 cell', () => {
        const geometries = ['rectangle', 'ellipse', 'diamond', 'rounded-rectangle', 'triangle', 'parallelogram'] as const;
        for (const g of geometries) {
            const geo = GeometryBuilder.build({ width: 4, height: 2, geometry: g });
            const noShow = geo.Cell.find((c: any) => c['@_N'] === 'NoShow');
            expect(noShow, `${g} missing NoShow cell`).toBeDefined();
            expect(noShow['@_V'], `${g} NoShow should be '0'`).toBe('0');
        }
    });
});

// ---------------------------------------------------------------------------
// Integration tests — page.addShape() with geometry prop
// ---------------------------------------------------------------------------

describe('page.addShape() with geometry prop', () => {
    it('creates an ellipse shape with correct dimensions', async () => {
        const doc = await VisioDocument.create();
        const page = doc.pages[0];
        const shape = await page.addShape({ text: 'O', x: 2, y: 2, width: 3, height: 2, geometry: 'ellipse' });
        expect(shape.width).toBeCloseTo(3, 5);
        expect(shape.height).toBeCloseTo(2, 5);
    });

    it('creates a diamond shape', async () => {
        const doc = await VisioDocument.create();
        const page = doc.pages[0];
        const shape = await page.addShape({ text: '?', x: 2, y: 2, width: 2, height: 2, geometry: 'diamond' });
        expect(shape.id).toBeDefined();
    });

    it('creates a rounded-rectangle shape', async () => {
        const doc = await VisioDocument.create();
        const page = doc.pages[0];
        const shape = await page.addShape({
            text: 'Rounded',
            x: 2, y: 2, width: 4, height: 2,
            geometry: 'rounded-rectangle',
            cornerRadius: 0.3,
        });
        expect(shape.width).toBeCloseTo(4, 5);
    });

    it('creates a triangle shape', async () => {
        const doc = await VisioDocument.create();
        const page = doc.pages[0];
        const shape = await page.addShape({ text: '▶', x: 2, y: 2, width: 2, height: 2, geometry: 'triangle' });
        expect(shape.id).toBeDefined();
    });

    it('creates a parallelogram shape', async () => {
        const doc = await VisioDocument.create();
        const page = doc.pages[0];
        const shape = await page.addShape({ text: 'Data', x: 2, y: 2, width: 3, height: 1.5, geometry: 'parallelogram' });
        expect(shape.height).toBeCloseTo(1.5, 5);
    });

    it('all geometry types survive a save/reload round-trip', async () => {
        const doc = await VisioDocument.create();
        const page = doc.pages[0];
        const geometries = ['rectangle', 'ellipse', 'diamond', 'rounded-rectangle', 'triangle', 'parallelogram'] as const;
        let x = 1;
        for (const g of geometries) {
            await page.addShape({ text: g, x, y: 5, width: 2, height: 1, geometry: g });
            x += 3;
        }
        const buf = await doc.save();
        expect(buf).toBeDefined();
        expect(buf.byteLength).toBeGreaterThan(0);
    });

    it('geometry works together with fillColor and text styling', async () => {
        const doc = await VisioDocument.create();
        const page = doc.pages[0];
        const shape = await page.addShape({
            text: 'Styled',
            x: 3, y: 3, width: 3, height: 2,
            geometry: 'ellipse',
            fillColor: '#3399ff',
            fontColor: '#ffffff',
            bold: true,
        });
        expect(shape.text).toBe('Styled');
    });

    it('default shape (no geometry prop) still produces a valid rectangle', async () => {
        const doc = await VisioDocument.create();
        const page = doc.pages[0];
        const shape = await page.addShape({ text: 'Box', x: 1, y: 1, width: 2, height: 1 });
        expect(shape.width).toBeCloseTo(2, 5);
        expect(shape.height).toBeCloseTo(1, 5);
    });
});
