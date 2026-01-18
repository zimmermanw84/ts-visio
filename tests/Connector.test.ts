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
        // Note: ShapeReader parses Cells into an object map
        expect(connector?.Cells['ObjType']?.V).toBe('2');

        // 2. Check Line Section exists (Visibility)
        // Note: ShapeReader parses Sections into an object map by Name
        expect(connector?.Sections['Line']).toBeDefined();
        // Section structure is { '@_N': 'Line', Cell: [...] }
        // ShapeReader parses rows/cells.
        // Let's assume standard parsing: Section -> Row (generic) or Cell (if property section)?
        // ShapeReader.ts: shape.Sections[section['@_N']] = parseSection(section);
        // StyleHelpers creates Line section with Cell array directly under Section, no Row?
        // Wait, ShapeModifier injected: Section: [ createLineSection... ]
        // createLineSection returns: { '@_N': 'Line', Cell: [...] }
        // So it has Cells directly.
        // Let's check how ShapeReader parses sections.

        // ShapeReader.ts:
        // parseSection just returns the object?
        // import { asArray, parseCells, parseSection } from './utils/VisioParsers';
        // I need to check VisioParsers to be sure how it handles it.
        // But likely it preserves structure.

        const lineSection = connector?.Sections['Line'];
        // It should have Cells
        // The test code I previously wrote assumed Row[0], but Line section usually has Cell directly if it's a Property section.
        // Let's check createLineSection again.
        // Cell: [ { LineColor... }, ... ]
        // So expectation: lineSection.Cell is array.

        expect(lineSection).toBeDefined();
        // Check Cells map
        expect(lineSection.Cells).toBeDefined();
        expect(lineSection.Cells['LineColor']).toBeDefined();

        // 3. Check Geometry formulas (Dynamic sizing)
        const geometry = connector?.Sections['Geometry'];
        const geomRows = geometry.Rows;
        expect(geomRows).toBeDefined();

        const lineToRow = geomRows.find((r: any) => r.T === 'LineTo');
        expect(lineToRow).toBeDefined();

        // Check Cell X formula
        // lineToRow.Cells is a map { Name: Cell }
        const xCell = lineToRow?.Cells['X'];
        expect(xCell).toBeDefined();
        expect(xCell?.F).toBe('Width');
    });
});
