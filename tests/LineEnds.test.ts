import { describe, it, expect, afterEach } from 'vitest';
import { VisioDocument } from '../src/VisioDocument';
import { ArrowHeads } from '../src/utils/StyleHelpers';
import fs from 'fs';
import path from 'path';

describe('Line Ends', () => {
    const testFile = path.resolve(__dirname, 'lines_test.vsdx');

    afterEach(() => {
        if (fs.existsSync(testFile)) {
            fs.unlinkSync(testFile);
        }
    });

    it('should create a connector with crows foot notation', async () => {
        const doc = await VisioDocument.create();
        const page = doc.pages[0];

        const box1 = await page.addShape({ text: 'User', x: 2, y: 4, width: 2, height: 1 });
        const box2 = await page.addShape({ text: 'List', x: 6, y: 4, width: 2, height: 1 });

        // Connect 1 (One) to N (Many)
        await page.connectShapes(box1, box2, ArrowHeads.One, ArrowHeads.CrowsFoot);

        await doc.save(testFile);
        expect(fs.existsSync(testFile)).toBe(true);
    });
});
