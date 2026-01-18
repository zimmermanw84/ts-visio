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
        const mainId = await page.addTable(4, 6, title, columns);

        expect(mainId).toBeDefined();

        await doc.save(testFile);
        expect(fs.existsSync(testFile)).toBe(true);
    });
});
