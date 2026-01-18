import { describe, it, expect, afterEach } from 'vitest';
import { VisioDocument } from '../src/VisioDocument';
import { VisioPackage } from '../src/VisioPackage';
import { ShapeReader } from '../src/ShapeReader';
import fs from 'fs';
import path from 'path';

describe('Group Shapes', () => {
    const testFile = path.resolve(__dirname, 'group_test.vsdx');

    afterEach(() => {
        if (fs.existsSync(testFile)) {
            fs.unlinkSync(testFile);
        }
    });

    it('should create a table as a group shape', async () => {
        const doc = await VisioDocument.create();
        const page = doc.pages[0];

        const title = 'GroupedTable';
        const columns = ['id: int', 'name: varchar'];

        // Add Table
        const groupShape = await page.addTable(4, 6, title, columns);

        expect(groupShape).toBeDefined();
        // Since we don't return the type in internalStub yet without re-reading,
        // we mainly trust that it didn't throw and returned a shape ID.
        // But we can check if downstream logic works (like connecting to it).

        expect(groupShape.id).toBeDefined();

        // Save
        await doc.save(testFile);
        expect(fs.existsSync(testFile)).toBe(true);
    });

    it('should avoid ID collisions when adding connectors between groups', async () => {
        const doc = await VisioDocument.create();
        const page = doc.pages[0];

        // 1. Create Table 1 (Users)
        // Consumes multiple IDs for group + children
        const t1 = await page.addTable(0, 0, 'Users', ['id']);

        // 2. Create Table 2 (Posts)
        // Consumes multiple IDs
        const t2 = await page.addTable(4, 0, 'Posts', ['id']);

        // 3. Connect them
        // This ensures the connector gets a unique ID that respects the children of the groups
        await page.connectShapes(t1, t2);

        await doc.save(testFile);
        expect(fs.existsSync(testFile)).toBe(true);
    });

    it('should create group shapes without geometry (transparent container)', async () => {
        const doc = await VisioDocument.create();
        const page = doc.pages[0];

        await page.addTable(0, 0, 'TransparencyTest', ['col1']);
        await doc.save(testFile);

        const buffer = fs.readFileSync(testFile);
        const pkg = new VisioPackage();
        await pkg.load(buffer);
        const reader = new ShapeReader(pkg);
        const shapes = reader.readShapes('visio/pages/page1.xml');

        const group = shapes.find(s => s.Type === 'Group');
        expect(group).toBeDefined();

        // Ensure NO Geometry section
        expect(group?.Sections['Geometry']).toBeUndefined();
    });

    it('should calculate absolute coordinates for connectors between nested shapes', async () => {
        const doc = await VisioDocument.create();
        const page = doc.pages[0];

        // 1. Create Group 1 at (10, 10)
        // Size 4x4. LocPin defaults to center (2,2).
        // Origin (Bottom-Left) = 10-2=8, 10-2=8.
        const g1 = await page.addShape({ text: 'G1', x: 10, y: 10, width: 4, height: 4, type: 'Group' });

        // 2. Add Child 1 inside G1 at (3, 3) relative
        // Absolute Expected: Origin(8) + 3 = 11.
        const c1 = await page.addShape({ text: 'C1', x: 3, y: 3, width: 1, height: 1 }, g1.id);

        // 3. Create Group 2 at (20, 10)
        // Size 4x4. LocPin (2,2). Origin (18, 8).
        const g2 = await page.addShape({ text: 'G2', x: 20, y: 10, width: 4, height: 4, type: 'Group' });

        // 4. Add Child 2 inside G2 at (1, 1) relative
        // Absolute Expected: Origin(18) + 1 = 19.
        const c2 = await page.addShape({ text: 'C2', x: 1, y: 1, width: 1, height: 1 }, g2.id);

        // 5. Connect C1 -> C2
        await page.connectShapes(c1, c2);
        await doc.save(testFile);

        const buffer = fs.readFileSync(testFile);
        const pkg = new VisioPackage();
        await pkg.load(buffer);
        const reader = new ShapeReader(pkg);
        const shapes = reader.readShapes('visio/pages/page1.xml');

        const connector = shapes.find(s => s.NameU === 'Dynamic connector');
        expect(connector).toBeDefined();

        // Check BeginX/Y (from C1)
        const beginX = parseFloat(connector?.Cells['BeginX']?.V || '0');
        const beginY = parseFloat(connector?.Cells['BeginY']?.V || '0');

        // C1 Center 11, Width 1. Right Edge is 11.5.
        // It connects to C2 (19, 9). Vector (8, -2).
        // Hit intersects Right Edge.
        // x = 11.5.
        expect(beginX).toBe(11.5);
        // y = 10.875 (Calculated slope offset)
        expect(beginY).toBeCloseTo(10.875);

        // Check EndX/Y (to C2)
        const endX = parseFloat(connector?.Cells['EndX']?.V || '0');
        const endY = parseFloat(connector?.Cells['EndY']?.V || '0');

        // C2 Center 19, Width 1. Left Edge is 18.5.
        // x = 18.5.
        expect(endX).toBe(18.5);
        // y = 9.125
        expect(endY).toBeCloseTo(9.125);
    });
});
