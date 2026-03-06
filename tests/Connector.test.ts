import { describe, it, expect, afterEach, vi } from 'vitest';
import { VisioDocument } from '../src/VisioDocument';
import { VisioPackage } from '../src/VisioPackage';
import { ShapeReader } from '../src/ShapeReader';
import { ConnectorEditor } from '../src/core/ConnectorEditor';
import { RelsManager } from '../src/core/RelsManager';
import { PageXmlCache } from '../src/core/PageXmlCache';
import fs from 'fs';
import path from 'path';

describe('Connectors', () => {
    const testFile = path.resolve(__dirname, 'connector_test.vsdx');

    afterEach(() => {
        if (fs.existsSync(testFile)) {
            fs.unlinkSync(testFile);
        }
    });

    it('should connect shapes on a page that starts empty (BUG 20)', async () => {
        // addConnector normalised <Shapes/> (Shape=undefined) as [undefined] before the fix,
        // producing malformed XML with an empty phantom shape before the connector.
        const doc = await VisioDocument.create();
        const page = doc.pages[0];

        // Add both shapes — first addShape call hits the empty-page normalisation path
        const a = await page.addShape({ text: 'A', x: 1, y: 1, width: 1, height: 1 });
        const b = await page.addShape({ text: 'B', x: 4, y: 1, width: 1, height: 1 });
        await page.connectShapes(a, b);

        const buffer = await doc.save();
        const pkg = new VisioPackage();
        await pkg.load(buffer);
        const reader = new ShapeReader(pkg);
        const shapes = reader.readShapes('visio/pages/page1.xml');

        // There must be exactly 3 shapes: A, B, and the connector — no phantom undefined shape
        expect(shapes).toHaveLength(3);
        expect(shapes.every(s => s.ID !== undefined)).toBe(true);
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

        // 2. Check Line Section exists (Visibility)
        const lineSection = connector?.Sections['Line'];
        if (!lineSection) throw new Error('Line section missing');
        if (!lineSection.Cells) throw new Error('Line section cells missing');

        expect(lineSection.Cells['LineColor']).toBeDefined();

        // Check Arrow Defaults as direct shape cells (Visio reads arrows from Shape cells, not Line section)
        expect(connector?.Cells['BeginArrow']?.V).toBe('0');
        expect(connector?.Cells['EndArrow']?.V).toBe('0');
        expect(connector?.Cells['BeginArrowSize']?.V).toBe('2'); // Default Medium
        expect(connector?.Cells['EndArrowSize']?.V).toBe('2'); // Default Medium

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

    it('regression bug-28: addConnector does NOT add MASTERS relationship when connector has no @_Master', async () => {
        const mockPkg = {
            getFileText: vi.fn().mockImplementation(() => { throw new Error('not found'); }),
            updateFile: vi.fn(),
        } as unknown as VisioPackage;

        const cache = new PageXmlCache(mockPkg);
        // Seed the cache with a minimal page containing two shapes
        const pageXml = `<?xml version="1.0" encoding="UTF-8"?>
<PageContents><Shapes>
  <Shape ID="1" Type="Shape"><Cell N="PinX" V="1"/><Cell N="PinY" V="1"/><Cell N="Width" V="1"/><Cell N="Height" V="1"/></Shape>
  <Shape ID="2" Type="Shape"><Cell N="PinX" V="4"/><Cell N="PinY" V="1"/><Cell N="Width" V="1"/><Cell N="Height" V="1"/></Shape>
</Shapes></PageContents>`;
        mockPkg.getFileText = vi.fn().mockReturnValue(pageXml);
        cache.registerPage('1', 'visio/pages/page1.xml');

        const relsManager = new RelsManager(mockPkg);
        const ensureSpy = vi.spyOn(relsManager, 'ensureRelationship');

        const editor = new ConnectorEditor(cache, relsManager);
        await editor.addConnector('1', '1', '2');

        // ensureRelationship must NOT have been called because the connector has no @_Master
        expect(ensureSpy).not.toHaveBeenCalled();
    });
});
