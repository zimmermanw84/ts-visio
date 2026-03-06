import { describe, it, expect } from 'vitest';
import { VisioDocument } from '../src/VisioDocument';

// ---------------------------------------------------------------------------
// shape.rotate()
// ---------------------------------------------------------------------------

describe('Shape.rotate()', () => {
    it('angle getter returns 0 when no rotation has been applied', async () => {
        const doc = await VisioDocument.create();
        const page = doc.pages[0];
        const shape = await page.addShape({ text: 'A', x: 1, y: 1, width: 2, height: 1 });

        expect(shape.angle).toBe(0);
    });

    it('rotate() sets angle and getter reflects the new value', async () => {
        const doc = await VisioDocument.create();
        const page = doc.pages[0];
        const shape = await page.addShape({ text: 'B', x: 1, y: 1, width: 2, height: 1 });

        await shape.rotate(90);

        expect(shape.angle).toBeCloseTo(90, 5);
    });

    it('rotate() replaces previous angle on subsequent calls', async () => {
        const doc = await VisioDocument.create();
        const page = doc.pages[0];
        const shape = await page.addShape({ text: 'C', x: 1, y: 1, width: 2, height: 1 });

        await shape.rotate(45);
        await shape.rotate(180);

        expect(shape.angle).toBeCloseTo(180, 5);
    });

    it('rotate(0) resets angle to 0', async () => {
        const doc = await VisioDocument.create();
        const page = doc.pages[0];
        const shape = await page.addShape({ text: 'D', x: 1, y: 1, width: 2, height: 1 });

        await shape.rotate(60);
        await shape.rotate(0);

        expect(shape.angle).toBeCloseTo(0, 5);
    });

    it('rotate() returns the shape for chaining', async () => {
        const doc = await VisioDocument.create();
        const page = doc.pages[0];
        const shape = await page.addShape({ text: 'E', x: 1, y: 1, width: 2, height: 1 });

        const result = await shape.rotate(30);
        expect(result).toBe(shape);
    });

    it('Angle cell is stored as radians in the underlying XML', async () => {
        const doc = await VisioDocument.create();
        const page = doc.pages[0];
        const shape = await page.addShape({ text: 'F', x: 1, y: 1, width: 2, height: 1 });

        await shape.rotate(90);

        const raw = shape['internalShape'].Cells['Angle']?.V;
        expect(raw).toBeDefined();
        const radians = parseFloat(raw!);
        expect(radians).toBeCloseTo(Math.PI / 2, 5);
    });
});

// ---------------------------------------------------------------------------
// shape.resize()
// ---------------------------------------------------------------------------

describe('Shape.resize()', () => {
    it('width and height getters reflect the new size after resize()', async () => {
        const doc = await VisioDocument.create();
        const page = doc.pages[0];
        const shape = await page.addShape({ text: 'G', x: 2, y: 2, width: 2, height: 1 });

        await shape.resize(4, 3);

        expect(shape.width).toBeCloseTo(4, 5);
        expect(shape.height).toBeCloseTo(3, 5);
    });

    it('resize() returns the shape for chaining', async () => {
        const doc = await VisioDocument.create();
        const page = doc.pages[0];
        const shape = await page.addShape({ text: 'H', x: 1, y: 1, width: 1, height: 1 });

        const result = await shape.resize(2, 2);
        expect(result).toBe(shape);
    });

    it('resize() throws when width is not positive', async () => {
        const doc = await VisioDocument.create();
        const page = doc.pages[0];
        const shape = await page.addShape({ text: 'I', x: 1, y: 1, width: 1, height: 1 });

        await expect(shape.resize(0, 1)).rejects.toThrow();
    });

    it('resize() throws when height is not positive', async () => {
        const doc = await VisioDocument.create();
        const page = doc.pages[0];
        const shape = await page.addShape({ text: 'J', x: 1, y: 1, width: 1, height: 1 });

        await expect(shape.resize(1, -1)).rejects.toThrow();
    });

    it('LocPinX is updated to width/2 after resize()', async () => {
        const doc = await VisioDocument.create();
        const page = doc.pages[0];
        const shape = await page.addShape({ text: 'K', x: 2, y: 2, width: 2, height: 1 });

        await shape.resize(6, 2);

        const locPinX = shape['internalShape'].Cells['LocPinX']?.V;
        expect(parseFloat(locPinX!)).toBeCloseTo(3, 5);
    });

    it('multiple resize() calls converge correctly', async () => {
        const doc = await VisioDocument.create();
        const page = doc.pages[0];
        const shape = await page.addShape({ text: 'L', x: 2, y: 2, width: 2, height: 2 });

        await shape.resize(3, 3);
        await shape.resize(5, 1);

        expect(shape.width).toBeCloseTo(5, 5);
        expect(shape.height).toBeCloseTo(1, 5);
    });

    it('updateShapeDimensions updates LocPinX and LocPinY', async () => {
        const doc  = await VisioDocument.create();
        const page = doc.pages[0];
        const shape = await page.addShape({ text: 'M2', x: 2, y: 2, width: 2, height: 1 });

        // Call the lower-level modifier path directly (used by ContainerEditor)
        await (page as any).modifier.updateShapeDimensions(page.id, shape.id, 6, 4);

        const parsed = await (page as any).modifier.cache.getParsed(page.id);
        const shapeMap = (page as any).modifier.cache.getShapeMap(parsed);
        const cells: any[] = shapeMap.get(shape.id).Cell;

        const locPinX = cells.find((c: any) => c['@_N'] === 'LocPinX');
        const locPinY = cells.find((c: any) => c['@_N'] === 'LocPinY');

        expect(parseFloat(locPinX['@_V'])).toBeCloseTo(3, 5); // 6 / 2
        expect(parseFloat(locPinY['@_V'])).toBeCloseTo(2, 5); // 4 / 2
    });

    it('resizeContainerToFit updates container LocPinX and LocPinY', async () => {
        const doc  = await VisioDocument.create();
        const page = doc.pages[0];
        const container = await page.addContainer({ text: 'C', x: 3, y: 3, width: 4, height: 3 });
        const member    = await page.addShape({ text: 'M', x: 3, y: 3, width: 1, height: 1 });
        await container.addMember(member);

        await container.resizeToFit();

        // After resize, LocPinX must equal Width/2 and LocPinY must equal Height/2
        const parsed   = await (page as any).modifier.cache.getParsed(page.id);
        const shapeMap = (page as any).modifier.cache.getShapeMap(parsed);
        const cells: any[] = shapeMap.get(container.id).Cell;

        const width    = parseFloat(cells.find((c: any) => c['@_N'] === 'Width')['@_V']);
        const height   = parseFloat(cells.find((c: any) => c['@_N'] === 'Height')['@_V']);
        const locPinX  = parseFloat(cells.find((c: any) => c['@_N'] === 'LocPinX')['@_V']);
        const locPinY  = parseFloat(cells.find((c: any) => c['@_N'] === 'LocPinY')['@_V']);

        expect(locPinX).toBeCloseTo(width  / 2, 5);
        expect(locPinY).toBeCloseTo(height / 2, 5);
    });

    // --- Bug 5 regression tests ---

    it('resize() updates geometry @_V for exact Width/Height formula references', async () => {
        // rectangle geometry uses 'Width' and 'Height' formulas exactly
        const doc  = await VisioDocument.create();
        const page = doc.pages[0];
        const shape = await page.addShape({ text: 'B5a', x: 1, y: 1, width: 2, height: 1, geometry: 'rectangle' });

        await shape.resize(6, 4);

        const parsed   = (page as any).modifier.cache.getParsed(page.id);
        const shapeMap = (page as any).modifier.cache.getShapeMap(parsed);
        const raw      = shapeMap.get(shape.id);
        const sections: any[] = Array.isArray(raw.Section) ? raw.Section : [raw.Section];
        const geoSection = sections.find((s: any) => s['@_N'] === 'Geometry');
        const rows: any[] = Array.isArray(geoSection.Row) ? geoSection.Row : [geoSection.Row];

        // Row 2 X cell has @_F='Width', should be updated to 6
        const row2cells: any[] = Array.isArray(rows[1].Cell) ? rows[1].Cell : [rows[1].Cell];
        const xCell = row2cells.find((c: any) => c['@_N'] === 'X' && c['@_F'] === 'Width');
        expect(parseFloat(xCell['@_V'])).toBeCloseTo(6, 5);

        // Row 4 Y cell has @_F='Height', should be updated to 4
        const row4cells: any[] = Array.isArray(rows[3].Cell) ? rows[3].Cell : [rows[3].Cell];
        const yCell = row4cells.find((c: any) => c['@_N'] === 'Y' && c['@_F'] === 'Height');
        expect(parseFloat(yCell['@_V'])).toBeCloseTo(4, 5);
    });

    it('resize() updates geometry @_V for compound formulas like Width*0.5 and Height*0.5', async () => {
        // ellipse geometry uses 'Width*0.5', 'Height*0.5', 'Width', 'Height' formulas
        const doc  = await VisioDocument.create();
        const page = doc.pages[0];
        const shape = await page.addShape({ text: 'B5b', x: 1, y: 1, width: 2, height: 2, geometry: 'ellipse' });

        await shape.resize(8, 4);

        const parsed   = (page as any).modifier.cache.getParsed(page.id);
        const shapeMap = (page as any).modifier.cache.getShapeMap(parsed);
        const raw      = shapeMap.get(shape.id);
        const sections: any[] = Array.isArray(raw.Section) ? raw.Section : [raw.Section];
        const geoSection = sections.find((s: any) => s['@_N'] === 'Geometry');
        const rows: any[] = Array.isArray(geoSection.Row) ? geoSection.Row : [geoSection.Row];
        const ellipseRow = rows[0];
        const cells: any[] = Array.isArray(ellipseRow.Cell) ? ellipseRow.Cell : [ellipseRow.Cell];

        // X cell @_F='Width*0.5' → 8*0.5 = 4
        const xCell = cells.find((c: any) => c['@_N'] === 'X');
        expect(parseFloat(xCell['@_V'])).toBeCloseTo(4, 5);

        // A cell @_F='Width' → 8
        const aCell = cells.find((c: any) => c['@_N'] === 'A');
        expect(parseFloat(aCell['@_V'])).toBeCloseTo(8, 5);

        // D cell @_F='Height' → 4
        const dCell = cells.find((c: any) => c['@_N'] === 'D');
        expect(parseFloat(dCell['@_V'])).toBeCloseTo(4, 5);
    });

    it('resize() updates geometry @_V for diamond shape with Width*0.5 and Height*0.5 formulas', async () => {
        const doc  = await VisioDocument.create();
        const page = doc.pages[0];
        const shape = await page.addShape({ text: 'B5c', x: 2, y: 2, width: 4, height: 2, geometry: 'diamond' });

        await shape.resize(10, 6);

        const parsed   = (page as any).modifier.cache.getParsed(page.id);
        const shapeMap = (page as any).modifier.cache.getShapeMap(parsed);
        const raw      = shapeMap.get(shape.id);
        const sections: any[] = Array.isArray(raw.Section) ? raw.Section : [raw.Section];
        const geoSection = sections.find((s: any) => s['@_N'] === 'Geometry');
        const rows: any[] = Array.isArray(geoSection.Row) ? geoSection.Row : [geoSection.Row];

        // Row 1 (MoveTo, top vertex): X=Width*0.5 → 5, Y=Height → 6
        const row1cells: any[] = Array.isArray(rows[0].Cell) ? rows[0].Cell : [rows[0].Cell];
        const topX = row1cells.find((c: any) => c['@_N'] === 'X' && c['@_F'] === 'Width*0.5');
        expect(parseFloat(topX['@_V'])).toBeCloseTo(5, 5);

        // Row 2 (right vertex): X=Width → 10, Y=Height*0.5 → 3
        const row2cells: any[] = Array.isArray(rows[1].Cell) ? rows[1].Cell : [rows[1].Cell];
        const rightY = row2cells.find((c: any) => c['@_N'] === 'Y' && c['@_F'] === 'Height*0.5');
        expect(parseFloat(rightY['@_V'])).toBeCloseTo(3, 5);
    });
});

// ---------------------------------------------------------------------------
// shape.flipX() / shape.flipY()
// ---------------------------------------------------------------------------

describe('Shape.flipX() / Shape.flipY()', () => {
    it('flipX() sets FlipX cell to 1', async () => {
        const doc = await VisioDocument.create();
        const page = doc.pages[0];
        const shape = await page.addShape({ text: 'M', x: 1, y: 1, width: 2, height: 1 });

        await shape.flipX();

        const raw = shape['internalShape'].Cells['FlipX']?.V;
        expect(raw).toBe('1');
    });

    it('flipX(false) sets FlipX cell to 0', async () => {
        const doc = await VisioDocument.create();
        const page = doc.pages[0];
        const shape = await page.addShape({ text: 'N', x: 1, y: 1, width: 2, height: 1 });

        await shape.flipX();
        await shape.flipX(false);

        const raw = shape['internalShape'].Cells['FlipX']?.V;
        expect(raw).toBe('0');
    });

    it('flipY() sets FlipY cell to 1', async () => {
        const doc = await VisioDocument.create();
        const page = doc.pages[0];
        const shape = await page.addShape({ text: 'O', x: 1, y: 1, width: 2, height: 1 });

        await shape.flipY();

        const raw = shape['internalShape'].Cells['FlipY']?.V;
        expect(raw).toBe('1');
    });

    it('flipY(false) sets FlipY cell to 0', async () => {
        const doc = await VisioDocument.create();
        const page = doc.pages[0];
        const shape = await page.addShape({ text: 'P', x: 1, y: 1, width: 2, height: 1 });

        await shape.flipY();
        await shape.flipY(false);

        const raw = shape['internalShape'].Cells['FlipY']?.V;
        expect(raw).toBe('0');
    });

    it('flipX() and flipY() can be combined', async () => {
        const doc = await VisioDocument.create();
        const page = doc.pages[0];
        const shape = await page.addShape({ text: 'Q', x: 1, y: 1, width: 2, height: 1 });

        await shape.flipX();
        await shape.flipY();

        expect(shape['internalShape'].Cells['FlipX']?.V).toBe('1');
        expect(shape['internalShape'].Cells['FlipY']?.V).toBe('1');
    });

    it('flipX() returns the shape for chaining', async () => {
        const doc = await VisioDocument.create();
        const page = doc.pages[0];
        const shape = await page.addShape({ text: 'R', x: 1, y: 1, width: 2, height: 1 });

        const result = await shape.flipX();
        expect(result).toBe(shape);
    });

    it('flipY() returns the shape for chaining', async () => {
        const doc = await VisioDocument.create();
        const page = doc.pages[0];
        const shape = await page.addShape({ text: 'S', x: 1, y: 1, width: 2, height: 1 });

        const result = await shape.flipY();
        expect(result).toBe(shape);
    });
});
