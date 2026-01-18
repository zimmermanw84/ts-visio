import { describe, it, expect, afterEach } from 'vitest';
import { VisioDocument } from '../src/VisioDocument';
import fs from 'fs';
import path from 'path';

describe('High Level API', () => {
    const testFile = path.resolve(__dirname, 'hl_test.vsdx');

    afterEach(() => {
        if (fs.existsSync(testFile)) {
            fs.unlinkSync(testFile);
        }
    });

    it('should create a document, add a shape, and modify it', async () => {
        const doc = await VisioDocument.create();
        const pages = doc.pages;
        expect(pages).toHaveLength(1);

        const page1 = pages[0];
        expect(page1.name).toBe('Page-1');

        // Add Shape
        const shape = await page1.addShape({
            text: 'Original Text',
            x: 1,
            y: 1,
            width: 2,
            height: 1
        });

        expect(shape.text).toBe('Original Text');
        expect(shape.width).toBe(2);

        // Modify Shape
        await shape.setText('Modified Text');
        expect(shape.text).toBe('Modified Text');

        // Verify with getShapes
        const shapes = page1.getShapes();
        expect(shapes).toHaveLength(1);
        expect(shapes[0].text).toBe('Modified Text');

        // Save
        await doc.save(testFile);
        expect(fs.existsSync(testFile)).toBe(true);
    });

    it('should load an existing document', async () => {
        // Create one first
        const initDoc = await VisioDocument.create();
        await initDoc.pages[0].addShape({ text: 'Loaded', x: 0, y: 0, width: 1, height: 1 });
        await initDoc.save(testFile);

        // Load it back
        const loadedDoc = await VisioDocument.load(testFile);
        const shapes = loadedDoc.pages[0].getShapes();
        expect(shapes).toHaveLength(1);
        expect(shapes[0].text).toBe('Loaded');
    });
});
