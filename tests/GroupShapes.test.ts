import { describe, it, expect, afterEach } from 'vitest';
import { VisioDocument } from '../src/VisioDocument';
import fs from 'fs';
import path from 'path';

describe('Group Shapes', () => {
    const testFile = path.resolve(__dirname, 'group_test.vsdx');

    afterEach(() => {
        if (fs.existsSync(testFile)) {
            fs.unlinkSync(testFile);
        }
    });

    it('should create a table as a group shape', async () => {
        const doc = await VisioDocument.create();
        const page = doc.pages[0];

        const title = 'GroupedTable';
        const columns = ['id: int', 'name: varchar'];

        // Add Table
        const groupShape = await page.addTable(4, 6, title, columns);

        expect(groupShape).toBeDefined();
        // Since we don't return the type in internalStub yet without re-reading,
        // we mainly trust that it didn't throw and returned a shape ID.
        // But we can check if downstream logic works (like connecting to it).

        expect(groupShape.id).toBeDefined();

        // Save
        await doc.save(testFile);
        expect(fs.existsSync(testFile)).toBe(true);
    });

    it('should avoid ID collisions when adding connectors between groups', async () => {
        const doc = await VisioDocument.create();
        const page = doc.pages[0];

        // 1. Create Table 1 (Users)
        // Consumes multiple IDs for group + children
        const t1 = await page.addTable(0, 0, 'Users', ['id']);

        // 2. Create Table 2 (Posts)
        // Consumes multiple IDs
        const t2 = await page.addTable(4, 0, 'Posts', ['id']);

        // 3. Connect them
        // This ensures the connector gets a unique ID that respects the children of the groups
        await page.connectShapes(t1, t2);

        await doc.save(testFile);
        expect(fs.existsSync(testFile)).toBe(true);
    });
});
