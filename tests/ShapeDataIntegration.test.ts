import { describe, it, expect } from 'vitest';
import { VisioDocument } from '../src/VisioDocument';
import { VisioPropType } from '../src/types/VisioTypes';

describe('Shape Data Integration', () => {
    it('should allow adding shape data via fluent API and save correctly', async () => {
        const doc = await VisioDocument.create();
        const page = doc.pages[0];

        // 1. Create Shape
        const shape = await page.addShape({ text: 'Server', x: 2, y: 2, width: 1, height: 1 });

        // 2. Add Data via Fluent API
        const now = new Date('2024-01-01T10:00:00Z');

        // Split chained calls to sequential awaits because addData returns a Promise
        await shape.addData('IP', { value: '192.168.1.5', label: 'IP Address' });
        await shape.addData('Cost', { value: 2500, type: VisioPropType.Currency }); // Explicit type override
        await shape.addData('InstallDate', { value: now }); // Auto-detect Date
        await shape.addData('IsActive', { value: true, hidden: true }); // Auto-detect Bool, Hidden

        // 3. Save and Reload
        const saved = await doc.save();
        const reloaded = await VisioDocument.load(saved);
        const reloadedPage = reloaded.pages[0];
        const reloadedShapes = await reloadedPage.getShapes();
        const serverShape = reloadedShapes[0];

        expect(serverShape).toBeDefined();

        // Verify internal mapped structure (ShapeReader transforms raw XML)
        const internal = (serverShape as any).internalShape;
        const propSection = internal.Sections['Property'];
        expect(propSection).toBeDefined();
        const rows = propSection.Rows || [];

        // 1. Check IP (String)
        const ipRow = rows.find((r: any) => r.N === 'Prop.IP'); // Using added 'N' property
        expect(ipRow).toBeDefined();
        expect(ipRow.Cells['Value'].V).toBe('192.168.1.5');
        expect(ipRow.Cells['Label'].V).toBe('IP Address');

        // 2. Check Cost (Currency/Number)
        const costRow = rows.find((r: any) => r.N === 'Prop.Cost');
        expect(costRow.Cells['Type'].V).toBe('7'); // Currency=7
        expect(costRow.Cells['Value'].V).toBe('2500');

        // 3. Check InstallDate (Date)
        const dateRow = rows.find((r: any) => r.N === 'Prop.InstallDate');
        expect(dateRow.Cells['Value'].V).toContain('2024-01-01T10:00:00');

        // 4. Check IsActive (Hidden Boolean)
        const activeRow = rows.find((r: any) => r.N === 'Prop.IsActive');
        expect(activeRow.Cells['Invisible'].V).toBe('1');
        expect(activeRow.Cells['Value'].V).toBe('1');
    });
});
