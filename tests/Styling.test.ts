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

    it('should create a shape with bold text and color', async () => {
        const doc = await VisioDocument.create();
        const page = doc.pages[0];

        const shape = await page.addShape({
            text: 'Bold Red Text',
            x: 5,
            y: 5,
            width: 2,
            height: 1,
            fillColor: '#FFFFFF',
            fontColor: '#FF0000',
            bold: true
        });

        await doc.save(testFile);
        expect(fs.existsSync(testFile)).toBe(true);
    });

    it('should allow fluent chaining of styles', async () => {
        const doc = await VisioDocument.create();
        const page = doc.pages[0];
        const shape = await page.addShape({ text: 'Fluent', x: 2, y: 2, width: 2, height: 1 });

        // Method Chaining
        await (await shape.setStyle({ fillColor: '#00FF00' })).setStyle({ bold: true });

        // Verify return value is the shape instance
        const s2 = await shape.setStyle({ fontColor: '#0000FF' });
        expect(s2).toBe(shape);
    });

    it('should allow linking with connectTo and chaining', async () => {
        const doc = await VisioDocument.create();
        const page = doc.pages[0];
        const s1 = await page.addShape({ text: 'A', x: 0, y: 0, width: 1, height: 1 });
        const s2 = await page.addShape({ text: 'B', x: 4, y: 4, width: 1, height: 1 });

        const ret = await s1.connectTo(s2);
        expect(ret).toBe(s1);
    });
});
