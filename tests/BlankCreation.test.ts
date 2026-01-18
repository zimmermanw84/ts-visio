import { describe, it, expect } from 'vitest';
import { VisioPackage } from '../src/VisioPackage';
import { PageManager } from '../src/PageManager';
import { ShapeReader } from '../src/ShapeReader';

describe('VisioPackage Creation', () => {
    it('should create a valid blank package', async () => {
        const pkg = await VisioPackage.create();

        // Verify files exist in memory
        expect(() => pkg.getFileText('[Content_Types].xml')).not.toThrow();
        expect(() => pkg.getFileText('visio/pages/pages.xml')).not.toThrow();

        // Verify PageManager can read it
        const pm = new PageManager(pkg);
        const pages = pm.getPages();
        expect(pages).toHaveLength(1);
        expect(pages[0].Name).toBe('Page-1');

        // Verify Page is empty
        const reader = new ShapeReader(pkg);
        const shapes = reader.readShapes('visio/pages/page1.xml');
        expect(shapes).toHaveLength(0);
    });

    it('should be saveable', async () => {
        const pkg = await VisioPackage.create();
        const buffer = await pkg.save();
        expect(buffer).toBeDefined();
        expect(buffer.length).toBeGreaterThan(0);

        // Reload to verify validity
        const newPkg = new VisioPackage();
        await newPkg.load(buffer);
        expect(() => newPkg.getFileText('visio/document.xml')).not.toThrow();
    });
});
