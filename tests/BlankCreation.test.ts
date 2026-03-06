import { describe, it, expect } from 'vitest';
import { VisioPackage } from '../src/VisioPackage';
import { PageManager } from '../src/core/PageManager';
import { ShapeReader } from '../src/ShapeReader';
import { createXmlParser } from '../src/utils/XmlHelper';

describe('VisioPackage Creation', () => {
    it('should create a valid blank package', async () => {
        const pkg = await VisioPackage.create();

        // Verify files exist in memory
        expect(() => pkg.getFileText('[Content_Types].xml')).not.toThrow();
        expect(() => pkg.getFileText('visio/pages/pages.xml')).not.toThrow();

        // Verify PageManager can read it
        const pm = new PageManager(pkg);
        const pages = pm.load();
        expect(pages).toHaveLength(1);
        expect(pages[0].name).toBe('Page-1');

        // Verify Page is empty
        const reader = new ShapeReader(pkg);
        const shapes = reader.readShapes('visio/pages/page1.xml');
        expect(shapes).toHaveLength(0);
    });

    it('initial page1.xml should have a PageSheet with US Letter dimensions', async () => {
        const pkg = await VisioPackage.create();
        const page1Xml = pkg.getFileText('visio/pages/page1.xml');
        const parser = createXmlParser();
        const parsed = parser.parse(page1Xml);

        const ps = parsed.PageContents?.PageSheet;
        expect(ps).toBeDefined();
        expect(ps['@_LineStyle']).toBe('0');
        expect(ps['@_FillStyle']).toBe('0');
        expect(ps['@_TextStyle']).toBe('0');

        const cells: any[] = Array.isArray(ps.Cell) ? ps.Cell : [ps.Cell];
        const pageWidth = cells.find((c: any) => c['@_N'] === 'PageWidth');
        const pageHeight = cells.find((c: any) => c['@_N'] === 'PageHeight');
        expect(pageWidth?.['@_V']).toBe('8.5');
        expect(pageHeight?.['@_V']).toBe('11');
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
