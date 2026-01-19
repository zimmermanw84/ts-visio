import { describe, it, expect } from 'vitest';
import { VisioDocument } from '../src/VisioDocument';

describe('Multi-Page High Level API', () => {
    it('should allow creating pages and adding shapes to them', async () => {
        const doc = await VisioDocument.create();

        // 1. Default Page
        const page1 = doc.pages[0];
        await page1.addShape({ text: 'Start', x: 1, y: 1, width: 1, height: 1 });

        // 2. Add Architecture Page
        const archPage = await doc.addPage('Architecture');
        expect(archPage.name).toBe('Architecture');
        expect(archPage.id).not.toBe(page1.id);

        await archPage.addShape({ text: 'Server', x: 2, y: 2, width: 1, height: 1 });

        // 3. Add Database Page
        const dbPage = await doc.addPage('Database');
        expect(dbPage.name).toBe('Database');

        await dbPage.addShape({ text: 'DB', x: 3, y: 3, width: 1, height: 1 });

        // 4. Save and Reload
        const saved = await doc.save();
        const reloaded = await VisioDocument.load(saved);

        expect(reloaded.pages).toHaveLength(3);

        const p1 = reloaded.pages.find(p => p.id === page1.id);
        expect(p1).toBeDefined();
        const shapes1 = await p1!.getShapes();
        expect(shapes1).toHaveLength(1);
        expect(shapes1[0].text).toBe('Start');

        const pArch = reloaded.pages.find(p => p.name === 'Architecture');
        expect(pArch).toBeDefined();
        const shapesArch = await pArch!.getShapes();
        expect(shapesArch).toHaveLength(1);
        expect(shapesArch[0].text).toBe('Server');

        const pDb = reloaded.pages.find(p => p.name === 'Database');
        expect(pDb).toBeDefined();
        const shapesDb = await pDb!.getShapes();
        expect(shapesDb).toHaveLength(1);
        expect(shapesDb[0].text).toBe('DB');
    });
});
