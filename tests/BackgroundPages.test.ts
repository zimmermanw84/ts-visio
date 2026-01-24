import { describe, it, expect } from 'vitest';
import { VisioDocument } from '../src/index';
import { XMLParser } from 'fast-xml-parser';

describe('Background Pages', () => {
    const parser = new XMLParser({ ignoreAttributes: false, attributeNamePrefix: '@_' });

    it('should create a background page with Background=true', async () => {
        const doc = await VisioDocument.create();

        const bgPage = await doc.addBackgroundPage('Background');

        // Verify page is marked as background
        const pagesXml = (doc as any).pkg.getFileText('visio/pages/pages.xml');
        const parsed = parser.parse(pagesXml);

        const pages = Array.isArray(parsed.Pages.Page) ? parsed.Pages.Page : [parsed.Pages.Page];
        const bgNode = pages.find((p: any) => p['@_ID'] === bgPage.id);

        expect(bgNode['@_Background'].toString()).toMatch(/1|true/);
    });

    it('should assign a background page to a foreground page', async () => {
        const doc = await VisioDocument.create();

        // Create background page
        const bgPage = await doc.addBackgroundPage('Grid Background');

        // Get foreground page
        const fgPage = doc.pages[0];

        // Assign background
        await doc.setBackgroundPage(fgPage, bgPage);

        // Verify BackPage attribute is set
        const pagesXml = (doc as any).pkg.getFileText('visio/pages/pages.xml');
        const parsed = parser.parse(pagesXml);

        const pages = Array.isArray(parsed.Pages.Page) ? parsed.Pages.Page : [parsed.Pages.Page];
        const fgNode = pages.find((p: any) => p['@_ID'] === fgPage.id);

        expect(fgNode['@_BackPage'].toString()).toBe(bgPage.id.toString());
    });

    it('should add shapes to a background page', async () => {
        const doc = await VisioDocument.create();

        // Create background page with grid lines
        const bgPage = await doc.addBackgroundPage('Grid');
        await bgPage.addShape({ text: 'Grid', x: 4, y: 5, width: 8, height: 10 });

        // Verify shape was added to background page
        const pageXml = (doc as any).pkg.getFileText(`visio/pages/page${bgPage.id}.xml`);
        expect(pageXml).toContain('Grid');
    });

    it('should throw error when setting non-background page as background', async () => {
        const doc = await VisioDocument.create();

        // Create a regular page (not background)
        const page2 = await doc.addPage('Page 2');
        const page1 = doc.pages[0];

        // Should throw because page2 is not a background page
        await expect(
            doc.setBackgroundPage(page1, page2)
        ).rejects.toThrow('is not a background page');
    });
});
