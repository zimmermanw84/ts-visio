import { describe, it, expect, afterEach } from 'vitest';
import { VisioDocument } from '../src/VisioDocument';
import fs from 'fs';
import path from 'path';

describe('Compound Shapes (Table)', () => {
    const testFile = path.resolve(__dirname, 'table_test.vsdx');

    afterEach(() => {
        if (fs.existsSync(testFile)) {
            fs.unlinkSync(testFile);
        }
    });

    it('should create a table with header and body', async () => {
        const doc = await VisioDocument.create();
        const page = doc.pages[0];

        const title = 'Users';
        const columns = ['id: int', 'username: varchar', 'email: varchar', 'created_at: timestamp'];

        // Add Table at (4, 6)
        // Add Table at (4, 6)
        const mainShape = await page.addTable(4, 6, title, columns);

        expect(mainShape).toBeDefined();
        expect(mainShape.id).toBeDefined(); // Check ID exists on object
        expect(mainShape.id).toBeDefined(); // Check ID exists on object
        // expect(mainShape.text).toBe(title); // OLD: returned header. NEW: returns Group (empty text)
        // Group shape itself has no text, text is on children.
        expect(mainShape).toBeDefined();

        await doc.save(testFile);
        expect(fs.existsSync(testFile)).toBe(true);
    });
});
