import { describe, it, expect, afterEach } from 'vitest';
import { VisioDocument } from '../src/VisioDocument';
import { VisioPackage } from '../src/VisioPackage';
import { ShapeReader } from '../src/ShapeReader';
import fs from 'fs';
import path from 'path';

describe('Container Shapes', () => {
    const testFile = path.resolve(__dirname, 'container_test.vsdx');

    afterEach(() => {
        if (fs.existsSync(testFile)) {
            fs.unlinkSync(testFile);
        }
    });

    it('should convert a shape to a container', async () => {
        const doc = await VisioDocument.create();
        const page = doc.pages[0];

        const shape = await page.addShape({ text: 'Container', x: 2, y: 2, width: 2, height: 2 });
        shape.convertToContainer();

        await doc.save(testFile);

        // Verify XML
        const buffer = fs.readFileSync(testFile);
        const pkg = new VisioPackage();
        await pkg.load(buffer);
        const reader = new ShapeReader(pkg);
        const shapes = reader.readShapes('visio/pages/page1.xml');

        const container = shapes.find(s => s.ID === shape.id);
        expect(container).toBeDefined();

        // Check for User Section
        const userSection = container?.Sections['User'];
        expect(userSection).toBeDefined();

        // Check for msvStructureType row
        // ShapeReader parses Rows into array.
        const structureRow = userSection?.Rows.find(r => r.N === 'User.msvStructureType');
        expect(structureRow).toBeDefined();

        // Check Value cell
        const valueCell = structureRow?.Cells['Value'];
        expect(valueCell).toBeDefined();
        expect(valueCell?.V).toBe('Container');
    });

    it('should create a table that is a container', async () => {
        const doc = await VisioDocument.create();
        const page = doc.pages[0];

        const table = await page.addTable(5, 5, 'MyTable', ['id']);
        expect(table).toBeDefined();

        await doc.save(testFile);

        const buffer = fs.readFileSync(testFile);
        const pkg = new VisioPackage();
        await pkg.load(buffer);
        const reader = new ShapeReader(pkg);
        const shapes = reader.readShapes('visio/pages/page1.xml');

        const group = shapes.find(s => s.ID === table.id);
        expect(group).toBeDefined();

        const userSection = group?.Sections['User'];
        expect(userSection).toBeDefined();

        const structureRow = userSection?.Rows.find(r => r.N === 'User.msvStructureType');
        expect(structureRow).toBeDefined();
        expect(structureRow?.Cells['Value'].V).toBe('Container');
    });
});
