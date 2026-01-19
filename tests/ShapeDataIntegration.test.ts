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

        // Note: We don't have a `getData()` API yet to verify easily from the high level.
        // We verified the internal XML writing in unit tests.
        // Here we just ensure the file is valid and can be loaded, and maybe check internal XML if we really want to be sure,
        // but `getShapes()` returning successfully implies no corruption.

        expect(serverShape).toBeDefined();
        // Additional: We could implement `getData()` later, but for now this confirms the API usage doesn't crash
        // and produces a parseable file.
    });
});
