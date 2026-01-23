import { describe, it, expect, beforeEach, vi } from 'vitest';
import { ShapeModifier } from '../src/ShapeModifier';
import { VisioPackage } from '../src/VisioPackage';
import { VisioDocument } from '../src/VisioDocument';
import { XMLParser } from 'fast-xml-parser';

describe('Shape Data Schema (ShapeModifier)', () => {
    let mockPkg: VisioPackage;
    let modifier: ShapeModifier;
    let parser: XMLParser;

    const pageId = '1';
    const shapeId = '5';
    const pagePath = 'visio/pages/page1.xml';

    // Minimal Page XML with one shape
    const initialXml = `<PageContents xmlns="http://schemas.microsoft.com/office/visio/2012/main">
    <Shapes>
        <Shape ID="5" Name="Sheet.5" Type="Shape">
            <Cell N="PinX" V="1"/>
            <Cell N="PinY" V="1"/>
        </Shape>
    </Shapes>
</PageContents>`;

    beforeEach(() => {
        mockPkg = {
            getFileText: vi.fn().mockReturnValue(initialXml),
            updateFile: vi.fn(),
        } as unknown as VisioPackage;

        modifier = new ShapeModifier(mockPkg);
        parser = new XMLParser({ ignoreAttributes: false, attributeNamePrefix: '@_' });
    });

    it('should create Property section and add a definition row', async () => {
        await modifier.addPropertyDefinition(pageId, shapeId, 'Cost', 2, { label: 'Project Cost' }); // 2 = Number

        const calls = vi.mocked(mockPkg.updateFile).mock.calls;
        expect(calls).toHaveLength(1);
        expect(calls[0][0]).toBe(pagePath);

        const newXml = calls[0][1];
        const parsed = parser.parse(newXml);
        const shape = parsed.PageContents.Shapes.Shape; // Should be object, not array for single shape

        expect(shape['@_ID']).toBe('5');

        // Verify Section
        const sections = Array.isArray(shape.Section) ? shape.Section : [shape.Section];
        const propSection = sections.find((s: any) => s['@_N'] === 'Property');
        expect(propSection).toBeDefined();

        // Verify Row
        const rows = Array.isArray(propSection.Row) ? propSection.Row : [propSection.Row];
        expect(rows).toHaveLength(1);
        const row = rows[0];
        expect(row['@_N']).toBe('Prop.Cost');

        // Verify Cells
        const cells = row.Cell;
        expect(cells).toBeDefined();
        // Helpers
        const getCell = (n: string) => cells.find((c: any) => c['@_N'] === n);

        expect(getCell('Label')['@_V']).toBe('Project Cost');
        expect(getCell('Type')['@_V']).toBe('2'); // Number
        expect(getCell('Invisible')['@_V']).toBe('0'); // Default Visible
    });

    it('should support adding invisible properties', async () => {
        await modifier.addPropertyDefinition(pageId, shapeId, 'SecretID', 0, { label: 'ID', invisible: true }); // visible=false

        const calls = vi.mocked(mockPkg.updateFile).mock.calls;
        const newXml = calls[0][1];
        const parsed = parser.parse(newXml);
        const shape = parsed.PageContents.Shapes.Shape;

        // Simplified access (assuming structure from previous test works)
        const sections = Array.isArray(shape.Section) ? shape.Section : [shape.Section];
        const propSection = sections.find((s: any) => s['@_N'] === 'Property');
        // actually, simpler to force array check:
        const rows = Array.isArray(propSection.Row) ? propSection.Row : [propSection.Row];
        const row0 = rows[0];

        const invisibleCell = row0.Cell.find((c: any) => c['@_N'] === 'Invisible');
        expect(invisibleCell['@_V']).toBe('1');
    });

    it('Hidden Metadata', async () => {
        // Requirement: Add 'EmployeeID' field with invisible: true. Assert XML has Invisible set to 1.
        await modifier.addPropertyDefinition(pageId, shapeId, 'EmployeeID', 0, { invisible: true });

        const calls = vi.mocked(mockPkg.updateFile).mock.calls;
        const newXml = calls[0][1];
        const parsed = parser.parse(newXml);
        const shape = parsed.PageContents.Shapes.Shape;

        const sections = Array.isArray(shape.Section) ? shape.Section : [shape.Section];
        const propSection = sections.find((s: any) => s['@_N'] === 'Property');
        const rows = Array.isArray(propSection.Row) ? propSection.Row : [propSection.Row];

        const employeeRow = rows.find((r: any) => r['@_N'] === 'Prop.EmployeeID');
        expect(employeeRow).toBeDefined();

        const invisibleCell = employeeRow.Cell.find((c: any) => c['@_N'] === 'Invisible');
        expect(invisibleCell).toBeDefined();
        expect(invisibleCell['@_V']).toBe('1');
    });

    it('should update existing definition if called again', async () => {
        // 1. Create first
        await modifier.addPropertyDefinition(pageId, shapeId, 'Cost', 2, { label: 'Cost' });

        // Mock getFileText to return the UPDATED xml for the second call
        const firstUpdateXml = vi.mocked(mockPkg.updateFile).mock.calls[0][1];
        vi.mocked(mockPkg.getFileText).mockReturnValue(firstUpdateXml);

        // 2. Update Label
        await modifier.addPropertyDefinition(pageId, shapeId, 'Cost', 2, { label: 'Updated Label' });

        const calls = vi.mocked(mockPkg.updateFile).mock.calls;
        expect(calls).toHaveLength(2);

        const finalXml = calls[1][1];
        const parsed = parser.parse(finalXml);
        const shape = parsed.PageContents.Shapes.Shape;
        const propSection = (Array.isArray(shape.Section) ? shape.Section : [shape.Section]).find((s: any) => s['@_N'] === 'Property');
        const rows = Array.isArray(propSection.Row) ? propSection.Row : [propSection.Row];

        expect(rows).toHaveLength(1); // Should not duplicate
        const row = rows[0];
        const labelCell = row.Cell.find((c: any) => c['@_N'] === 'Label');
        expect(labelCell['@_V']).toBe('Updated Label');
    });

    it('should set property values correctly', async () => {
        // Setup: Create definition first
        await modifier.addPropertyDefinition(pageId, shapeId, 'Cost', 2);

        // Mock state for next call
        const defXml = vi.mocked(mockPkg.updateFile).mock.calls[0][1];
        vi.mocked(mockPkg.getFileText).mockReturnValue(defXml);
        vi.mocked(mockPkg.updateFile).mockClear();

        // 1. Set Value
        await modifier.setPropertyValue(pageId, shapeId, 'Cost', 100);

        const calls = vi.mocked(mockPkg.updateFile).mock.calls;
        expect(calls).toHaveLength(1);

        const newXml = calls[0][1];
        const parsed = parser.parse(newXml);
        const shape = parsed.PageContents.Shapes.Shape;
        const propSection = (Array.isArray(shape.Section) ? shape.Section : [shape.Section]).find((s: any) => s['@_N'] === 'Property');

        const rows = Array.isArray(propSection.Row) ? propSection.Row : [propSection.Row];
        const costRow = rows.find((r: any) => r['@_N'] === 'Prop.Cost');

        const valueCell = costRow.Cell.find((c: any) => c['@_N'] === 'Value');
        expect(valueCell['@_V']).toBe('100');
    });

    it('should serialize Date objects to ISO string', async () => {
        // Setup definition
        await modifier.addPropertyDefinition(pageId, shapeId, 'Deadline', 5); // 5=Date
        const defXml = vi.mocked(mockPkg.updateFile).mock.calls[0][1];
        vi.mocked(mockPkg.getFileText).mockReturnValue(defXml);
        vi.mocked(mockPkg.updateFile).mockClear();

        const date = new Date('2023-12-25T12:00:00Z');
        await modifier.setPropertyValue(pageId, shapeId, 'Deadline', date);

        const calls = vi.mocked(mockPkg.updateFile).mock.calls;
        const newXml = calls[0][1];
        const parsed = parser.parse(newXml);

        const shape = parsed.PageContents.Shapes.Shape;
        const propSection = (Array.isArray(shape.Section) ? shape.Section : [shape.Section]).find((s: any) => s['@_N'] === 'Property');
        const rows = Array.isArray(propSection.Row) ? propSection.Row : [propSection.Row];

        const row = rows[0];
        const valueCell = row.Cell.find((c: any) => c['@_N'] === 'Value');
        // Check standard ISO string format that Visio accepts
        expect(valueCell['@_V']).toContain('2023-12-25T12:00:00');
    });

    it('should handle String and Boolean values', async () => {
        // Setup IsActive
        await modifier.addPropertyDefinition(pageId, shapeId, 'IsActive', 3);
        const xmlWithIfActive = vi.mocked(mockPkg.updateFile).mock.calls[0][1];

        // Test Boolean (IsActive)
        vi.mocked(mockPkg.getFileText).mockReturnValue(xmlWithIfActive);
        vi.mocked(mockPkg.updateFile).mockClear();

        await modifier.setPropertyValue(pageId, shapeId, 'IsActive', true);

        const boolXml = vi.mocked(mockPkg.updateFile).mock.calls[0][1];
        const parsedBool = parser.parse(boolXml);
        const boolShape = parsedBool.PageContents.Shapes.Shape;
        const boolRows = (Array.isArray(boolShape.Section) ? boolShape.Section : [boolShape.Section])
            .find((s: any) => s['@_N'] === 'Property').Row;
        const boolRow = (Array.isArray(boolRows) ? boolRows : [boolRows]).find((r: any) => r['@_N'] === 'Prop.IsActive');
        expect(boolRow.Cell.find((c: any) => c['@_N'] === 'Value')['@_V']).toBe('1');

        // Test String (Role) - Clean slate to avoid dependency on previous mock state
        vi.mocked(mockPkg.getFileText).mockReturnValue(initialXml);
        vi.mocked(mockPkg.updateFile).mockClear();

        await modifier.addPropertyDefinition(pageId, shapeId, 'Role', 0);
        const xmlWithRole = vi.mocked(mockPkg.updateFile).mock.calls[0][1];

        vi.mocked(mockPkg.getFileText).mockReturnValue(xmlWithRole);
        vi.mocked(mockPkg.updateFile).mockClear();

        await modifier.setPropertyValue(pageId, shapeId, 'Role', 'Manager');

        const strXml = vi.mocked(mockPkg.updateFile).mock.calls[0][1];
        const parsedStr = parser.parse(strXml);
        const strShape = parsedStr.PageContents.Shapes.Shape;
        const strRows = (Array.isArray(strShape.Section) ? strShape.Section : [strShape.Section])
            .find((s: any) => s['@_N'] === 'Property').Row;
        const strRow = (Array.isArray(strRows) ? strRows : [strRows]).find((r: any) => r['@_N'] === 'Prop.Role');
        expect(strRow.Cell.find((c: any) => c['@_N'] === 'Value')['@_V']).toBe('Manager');
    });

    // bugfix
    it('should update NextShapeID in PageSheet when adding a shape', async () => {
        const doc = await VisioDocument.create();
        const page = doc.pages[0];

        // 1. Add a shape (ID should be 1, usually)
        const shape = await page.addShape({ text: 'Test', x: 1, y: 1, width: 1, height: 1 });
        const shapeId = parseInt(shape.id);

        // 2. Save and inspect XML
        const buffer = await doc.save();
        const pkg = (doc as any).pkg; // Access internal package
        const pageXml = pkg.getFileText('visio/pages/page1.xml');

        const parser = new XMLParser({ ignoreAttributes: false, attributeNamePrefix: "@_" });
        const parsed = parser.parse(pageXml);

        // 3. Find PageSheet
        const pageSheet = parsed.PageContents.PageSheet;
        expect(pageSheet).toBeDefined();

        // 4. Find NextShapeID cell
        const cells = Array.isArray(pageSheet.Cell) ? pageSheet.Cell : [pageSheet.Cell];
        const nextShapeIdCell = cells.find((c: any) => c['@_N'] === 'NextShapeID');

        // 5. Assert existence and value
        // NextShapeID must be greater than the current shape ID to prevent ID conflicts
        expect(nextShapeIdCell).toBeDefined();
        const nextIdValue = parseInt(nextShapeIdCell['@_V']);

        expect(nextIdValue).toBeGreaterThan(shapeId);
    });
});
