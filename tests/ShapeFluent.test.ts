import { describe, it, expect, afterEach } from 'vitest';
import { VisioDocument } from '../src/VisioDocument';
import path from 'path';
import fs from 'fs';

describe('Shape Fluent API', () => {
    const testFile = path.resolve(__dirname, 'fluent_test.vsdx');

    afterEach(() => {
        if (fs.existsSync(testFile)) {
            fs.unlinkSync(testFile);
        }
    });

    it('should allow chaining addData calls', async () => {
        const doc = await VisioDocument.create();
        const page = doc.pages[0];

        // Use addTable to get a wrapper shape quickly
        const shape = await page.addTable(0, 0, 'Server', []);

        // Fluent Chaining
        shape.addData('ip', { label: 'IP Address', value: '192.168.1.1' })
            .addData('id', { value: 12345, hidden: true })
            .addData('active', { value: true });

        // Save and verify existence
        await doc.save(testFile);
        expect(fs.existsSync(testFile)).toBe(true);

        // We could inspect internals, but the fact it didn't crash and returned 'this' suggests chaining worked.
        // Let's inspect the last added property to be sure.
        // We can use proper XML inspection or just trust the 'not throwing' for chaining + previous tests covering logic.
        // Let's do a quick verify of the XML content to be robust.

        const pageXml = (doc as any).pkg.filesMap.get('visio/pages/page1.xml');
        expect(pageXml).toBeDefined();

        expect(pageXml).toMatch(/Prop\.ip/);
        expect(pageXml).toMatch(/192\.168\.1\.1/);

        expect(pageXml).toMatch(/Prop\.id/);
        expect(pageXml).toMatch(/12345/);
        expect(pageXml).toMatch(/N=["']Invisible["'][^>]*V=["']1["']/); // Hidden

        expect(pageXml).toMatch(/Prop\.active/);
    });
});
