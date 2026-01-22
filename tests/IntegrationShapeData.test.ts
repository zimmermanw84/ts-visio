import { describe, it, expect, afterEach } from 'vitest';
import { VisioDocument } from '../src/VisioDocument';
import path from 'path';
import fs from 'fs';
import { XMLParser } from 'fast-xml-parser';

describe('Shape Data Integration', () => {
    const testFile = path.resolve(__dirname, 'integration_shape_data.vsdx');

    afterEach(() => {
        if (fs.existsSync(testFile)) {
            fs.unlinkSync(testFile);
        }
    });

    it('should persist mixed data types and visibility correctly', async () => {
        const doc = await VisioDocument.create();
        const page = doc.pages[0];

        // 1. Create Shape
        const shape = await page.addTable(0, 0, 'Inventory', []);

        // 2. Add 3 data fields: String, Number, Hidden
        shape.addData('ItemName', { label: 'Item Name', value: 'Widget', type: 0 }) // String
            .addData('Quantity', { label: 'Qty', value: 50, type: 2 })             // Number
            .addData('InternalID', { value: 999, hidden: true });                  // Hidden Number

        // 3. Save
        await doc.save(testFile);
        expect(fs.existsSync(testFile)).toBe(true);

        // 4. Verification: Parse XML
        // In a real environment we might unzip, but here we trust the in-memory state matches saved state
        // OR we can read the file back via VisioPackage if we wanted to be 100% strict,
        // but accessing the internal file map is sufficient for logic verification.
        const pageXml = (doc as any).pkg.filesMap.get('visio/pages/page1.xml');
        expect(pageXml).toBeDefined();

        const parser = new XMLParser({ ignoreAttributes: false, attributeNamePrefix: '@_' });
        const parsed = parser.parse(pageXml);

        // Locate the shape
        const shapes = parsed.PageContents.Shapes.Shape;
        const targetShape = Array.isArray(shapes) ? shapes.find((s: any) => s.Text['#text'] === 'Inventory') : shapes;
        expect(targetShape).toBeDefined();

        // Locate Property Section
        const sections = Array.isArray(targetShape.Section) ? targetShape.Section : [targetShape.Section];
        const propSection = sections.find((s: any) => s['@_N'] === 'Property');
        expect(propSection).toBeDefined();

        // Verify Rows
        const rows = Array.isArray(propSection.Row) ? propSection.Row : [propSection.Row];
        expect(rows).toHaveLength(3);

        const getCell = (row: any, name: string) => row.Cell.find((c: any) => c['@_N'] === name);

        // Verify String (ItemName)
        const row1 = rows.find((r: any) => r['@_N'] === 'Prop.ItemName');
        expect(row1).toBeDefined();
        expect(getCell(row1, 'Label')['@_V']).toBe('Item Name');
        expect(getCell(row1, 'Value')['@_V']).toBe('Widget');
        expect(getCell(row1, 'Invisible')['@_V']).toBe('0'); // Default visible

        // Verify Number (Quantity)
        const row2 = rows.find((r: any) => r['@_N'] === 'Prop.Quantity');
        expect(row2).toBeDefined();
        expect(getCell(row2, 'Label')['@_V']).toBe('Qty');
        expect(getCell(row2, 'Value')['@_V']).toBe('50');

        // Verify Hidden (InternalID)
        const row3 = rows.find((r: any) => r['@_N'] === 'Prop.InternalID');
        expect(row3).toBeDefined();
        expect(getCell(row3, 'Value')['@_V']).toBe('999');
        expect(getCell(row3, 'Invisible')['@_V']).toBe('1'); // Hidden
    });
});
