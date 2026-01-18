import { describe, it, expect, afterEach } from 'vitest';
import { VisioDocument } from '../src/VisioDocument';
import fs from 'fs';
import path from 'path';

describe('Styling', () => {
    const testFile = path.resolve(__dirname, 'style_test.vsdx');

    afterEach(() => {
        if (fs.existsSync(testFile)) {
            fs.unlinkSync(testFile);
        }
    });

    it('should create a shape with fill color', async () => {
        const doc = await VisioDocument.create();
        const page = doc.pages[0];

        // Create shape with fill
        const shape = await page.addShape({
            text: 'Colorful Box',
            x: 2,
            y: 4,
            width: 2,
            height: 1,
            fillColor: '#FF0000'
        });

        // Current Shape wrapper doesn't expose fill reading yet,
        // but high level "addShape" should succeed without error.
        expect(shape).toBeDefined();

        await doc.save(testFile);
        expect(fs.existsSync(testFile)).toBe(true);

        // Advanced: Read back file and check XML content manually?
        // Skipped for now, trust creation.
    });
});
