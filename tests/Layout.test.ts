import { describe, it, expect, afterEach } from 'vitest';
import { VisioDocument } from '../src/VisioDocument';
import fs from 'fs';
import path from 'path';

describe('Layout', () => {
    const testFile = path.resolve(__dirname, 'layout_test.vsdx');

    afterEach(() => {
        if (fs.existsSync(testFile)) {
            fs.unlinkSync(testFile);
        }
    });

    it('should place a shape to the right of another', async () => {
        const doc = await VisioDocument.create();
        const page = doc.pages[0];

        // Shape A at (2, 2) width 2
        const s1 = await page.addShape({ text: 'A', x: 2, y: 2, width: 2, height: 1 });

        // Shape B at arbitrary position
        const s2 = await page.addShape({ text: 'B', x: 0, y: 0, width: 1, height: 1 });

        // Place B right of A with gap 1 (default)
        // Target x = 2 + 2 + 1 = 5
        // Target y = 2
        await s2.placeRightOf(s1);

        expect(s2.x).toBe(5);
        expect(s2.y).toBe(2);
    });

    it('should place shape with custom gap', async () => {
        const doc = await VisioDocument.create();
        const page = doc.pages[0];
        const s1 = await page.addShape({ text: 'A', x: 2, y: 2, width: 2, height: 1 });
        const s2 = await page.addShape({ text: 'B', x: 0, y: 0, width: 1, height: 1 });

        await s2.placeRightOf(s1, { gap: 0.5 });

        // 2 + 2 + 0.5 = 4.5
        expect(s2.x).toBe(4.5);
    });
    it('should place a shape below another', async () => {
        const doc = await VisioDocument.create();
        const page = doc.pages[0];
        const s1 = await page.addShape({ text: 'A', x: 2, y: 5, width: 2, height: 1 });
        const s2 = await page.addShape({ text: 'B', x: 0, y: 0, width: 2, height: 1 });

        await s2.placeBelow(s1, { gap: 1 });

        // X aligned: 2
        // Y: 5 - (1+1)/2 - 1 = 3
        expect(s2.x).toBe(2);
        expect(s2.y).toBe(3);
    });
});
