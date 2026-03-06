/**
 * Regression: Bug 24 — Page read methods returned stale data when autoSave=false.
 *
 * getShapes() / getShapeById() / findShapes() / getConnectors() all used
 * ShapeReader(pkg) which read directly from pkg. When autoSave=false, mutations
 * live only in PageXmlCache.dirtyPages until flush(). The fix passes a
 * cache-aware XML override so reads reflect in-flight mutations immediately.
 */
import { describe, it, expect } from 'vitest';
import { VisioDocument } from '../src/VisioDocument';

describe('regression bug-24: Page read methods honour dirty cache when autoSave=false', () => {
    it('getShapes() returns newly added shape before flush', async () => {
        const doc = await VisioDocument.create();
        const page = doc.pages[0];

        // Disable auto-save so mutations stay in cache, not written to pkg
        doc.autoSave = false;

        await page.addShape({ text: 'Dirty Shape', x: 1, y: 1, width: 1, height: 1 });

        // Without the fix, getShapes() would return 0 shapes (stale pkg read)
        const shapes = page.getShapes();
        expect(shapes.length).toBeGreaterThan(0);
        expect(shapes.some(s => s.text === 'Dirty Shape')).toBe(true);
    });

    it('getShapeById() finds a shape added before flush', async () => {
        const doc = await VisioDocument.create();
        const page = doc.pages[0];
        doc.autoSave = false;

        const added = await page.addShape({ text: 'FindMe', x: 2, y: 2, width: 1, height: 1 });

        const found = page.getShapeById(added.id);
        expect(found).toBeDefined();
        expect(found?.id).toBe(added.id);
    });

    it('findShapes() locates shapes added before flush', async () => {
        const doc = await VisioDocument.create();
        const page = doc.pages[0];
        doc.autoSave = false;

        await page.addShape({ text: 'Alpha', x: 1, y: 1, width: 1, height: 1 });
        await page.addShape({ text: 'Beta',  x: 2, y: 2, width: 1, height: 1 });

        const results = page.findShapes(s => s.text === 'Alpha');
        expect(results.length).toBe(1);
        expect(results[0].text).toBe('Alpha');
    });
});
