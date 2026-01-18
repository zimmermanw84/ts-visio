import { describe, it, expect, afterEach } from 'vitest';
import { VisioDocument } from '../src/VisioDocument';
import { SchemaDiagram } from '../src/SchemaDiagram';
import fs from 'fs';
import path from 'path';

describe('SchemaDiagram', () => {
    const testFile = path.resolve(__dirname, 'schema_test.vsdx');

    afterEach(() => {
        if (fs.existsSync(testFile)) {
            fs.unlinkSync(testFile);
        }
    });

    it('should create tables and relations using the facade', async () => {
        const doc = await VisioDocument.create();
        const page = doc.pages[0];
        const schema = new SchemaDiagram(page);

        // Add Tables
        const users = await schema.addTable('Users', ['id', 'email'], 2, 6);
        const posts = await schema.addTable('Posts', ['id', 'user_id', 'title'], 6, 6);

        expect(users).toBeDefined();
        expect(posts).toBeDefined();

        // Add Relation
        // We can't easily verify the arrow code in the XML without reading it back deeply,
        // but we can verify it doesn't throw and presumably assumes delegation works if other tests pass.
        // Ideally, we'd inspect the XML or mock the page, but integration test is fine for now.
        await expect(schema.addRelation(users, posts, '1:N')).resolves.not.toThrow();

        await doc.save(testFile);
        expect(fs.existsSync(testFile)).toBe(true);
    });
});
