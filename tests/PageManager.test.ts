import { describe, it, expect, beforeEach } from 'vitest';
import JSZip from 'jszip';
import { VisioPackage } from '../src/VisioPackage';
import { PageManager } from '../src/PageManager';

describe('PageManager', () => {
    let pkg: VisioPackage;
    let pm: PageManager;

    beforeEach(async () => {
        pkg = new VisioPackage();
        const zip = new JSZip();
        zip.file('visio/pages/pages.xml', `
            <Pages>
                <Page ID="0" Name="Page-1" NameU="Page-1" />
                <Page ID="4" Name="Another Page" NameU="Another Page" />
            </Pages>
        `);
        // Mock a missing file case by not creating it in another test if needed,
        // but here we set it up.

        const buffer = await zip.generateAsync({ type: 'nodebuffer' });
        await pkg.load(buffer);
        pm = new PageManager(pkg);
    });

    it('should list pages correctly', () => {
        const pages = pm.getPages();
        expect(pages).toHaveLength(2);
        expect(pages[0].Name).toBe('Page-1');
        expect(pages[0].ID).toBe('0');
        expect(pages[1].Name).toBe('Another Page');
        expect(pages[1].ID).toBe('4');
    });

    it('should return empty array if pages.xml missing', async () => {
        const emptyPkg = new VisioPackage();
        const zip = new JSZip(); // empty
        const buffer = await zip.generateAsync({ type: 'nodebuffer' });
        await emptyPkg.load(buffer);

        const pmEmpty = new PageManager(emptyPkg);
        expect(pmEmpty.getPages()).toEqual([]);
    });
});
