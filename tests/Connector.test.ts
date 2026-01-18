import { describe, it, expect, afterEach } from 'vitest';
import { VisioDocument } from '../src/VisioDocument';
import { VisioPackage } from '../src/VisioPackage';
import { ShapeReader } from '../src/ShapeReader';
import fs from 'fs';
import path from 'path';

describe('Connectors', () => {
    const testFile = path.resolve(__dirname, 'connector_test.vsdx');

    afterEach(() => {
        if (fs.existsSync(testFile)) {
            fs.unlinkSync(testFile);
        }
    });

    it('should create a connector between two shapes with correct visibility properties', async () => {
        const doc = await VisioDocument.create();
        const page = doc.pages[0];

        // Create two shapes
        const box1 = await page.addShape({ text: 'Box 1', x: 2, y: 4, width: 2, height: 1 });
        const box2 = await page.addShape({ text: 'Box 2', x: 6, y: 4, width: 2, height: 1 });

        await page.connectShapes(box1, box2);

        await doc.save(testFile);

        // Read back to verify low-level properties
        const buffer = fs.readFileSync(testFile);
        const pkg = new VisioPackage();
        await pkg.load(buffer);
        const reader = new ShapeReader(pkg);
        const shapes = reader.readShapes('visio/pages/page1.xml');

        const connector = shapes.find(s => s.NameU === 'Dynamic connector');
        expect(connector).toBeDefined();

        // 1. Check ObjType = 2 (1D Shape)
        expect(connector?.Cells['ObjType']?.V).toBe('2');

        // Check coordinates match connected shapes (Edge Routing)
        // Box 1 (Center 2, Width 2) -> Right Edge is 3
        // Box 2 (Center 6, Width 2) -> Left Edge is 5
        expect(connector?.Cells['BeginX']?.V).toBe('3');
        expect(connector?.Cells['BeginY']?.V).toBe('4');
        expect(connector?.Cells['EndX']?.V).toBe('5');
        expect(connector?.Cells['EndY']?.V).toBe('4');

        // Check pre-calculated Width/Angle for visibility
        // dx = 2 (5-3), dy = 0 -> Width = 2, Angle = 0
        expect(connector?.Cells['Width']?.V).toBe('2');
        expect(connector?.Cells['Angle']?.V).toBe('0');

        // Check PinX/PinY (Midpoint)
        // Begin(3,4), End(5,4). Mid = (4, 4)
        expect(connector?.Cells['PinX']?.V).toBe('4');
        expect(connector?.Cells['PinY']?.V).toBe('4');

        // 2. Check Line Section exists (Visibility & Arrows)
        const lineSection = connector?.Sections['Line'];
        if (!lineSection) throw new Error('Line section missing');
        if (!lineSection.Cells) throw new Error('Line section cells missing');

        expect(lineSection.Cells['LineColor']).toBeDefined();

        // Check Arrow Defaults (should be '0' if not specified, but verify they EXIST)
        expect(lineSection.Cells['BeginArrow']?.V).toBe('0');
        expect(lineSection.Cells['EndArrow']?.V).toBe('0');
        expect(lineSection.Cells['BeginArrowSize']?.V).toBe('2'); // Default Medium
        expect(lineSection.Cells['EndArrowSize']?.V).toBe('2'); // Default Medium

        // 3. Check Geometry formulas (Dynamic sizing)
        const geometry = connector?.Sections['Geometry'];
        if (!geometry) throw new Error('Geometry section missing');
        const geomRows = geometry.Rows;
        expect(geomRows).toBeDefined();

        const lineToRow = geomRows.find((r: any) => r.T === 'LineTo');
        expect(lineToRow).toBeDefined();

        // Check Cell X formula
        const xCell = lineToRow?.Cells['X'];
        expect(xCell).toBeDefined();
        expect(xCell?.F).toBe('Width');
        // Ensure static value is also set for immediate visibility
        expect(xCell?.V).toBe('2');
    });
});
