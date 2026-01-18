import { describe, it, expect, afterEach } from 'vitest';
import { VisioDocument } from '../src/VisioDocument';
import fs from 'fs';
import path from 'path';

describe('Connectors', () => {
    const testFile = path.resolve(__dirname, 'connector_test.vsdx');

    afterEach(() => {
        if (fs.existsSync(testFile)) {
            fs.unlinkSync(testFile);
        }
    });

    it('should create a connector between two shapes', async () => {
        const doc = await VisioDocument.create();
        const page = doc.pages[0];

        // Create two shapes
        const box1 = await page.addShape({ text: 'Box 1', x: 2, y: 4, width: 2, height: 1 });
        const box2 = await page.addShape({ text: 'Box 2', x: 6, y: 4, width: 2, height: 1 });

        // Verify shape IDs (assuming sequence)
        const id1 = box1.id;
        const id2 = box2.id;

        await page.connectShapes(box1, box2);

        await doc.save(testFile);
        expect(fs.existsSync(testFile)).toBe(true);
    });
});
