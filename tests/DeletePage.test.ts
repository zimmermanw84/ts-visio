import { describe, it, expect } from 'vitest';
import { VisioDocument } from '../src/VisioDocument';
import { XMLParser } from 'fast-xml-parser';

const parser = new XMLParser({ ignoreAttributes: false, attributeNamePrefix: '@_' });

async function getZipEntries(doc: VisioDocument): Promise<Set<string>> {
    const buf = await doc.save();
    const JSZip = (await import('jszip')).default;
    const zip = await JSZip.loadAsync(buf);
    const names = new Set<string>();
    zip.forEach((path) => names.add(path));
    return names;
}

async function getPagesXml(doc: VisioDocument): Promise<any> {
    const buf = await doc.save();
    const JSZip = (await import('jszip')).default;
    const zip = await JSZip.loadAsync(buf);
    const xml = await zip.file('visio/pages/pages.xml')!.async('string');
    return parser.parse(xml);
}

async function getPagesRelsXml(doc: VisioDocument): Promise<any> {
    const buf = await doc.save();
    const JSZip = (await import('jszip')).default;
    const zip = await JSZip.loadAsync(buf);
    const xml = await zip.file('visio/pages/_rels/pages.xml.rels')!.async('string');
    return parser.parse(xml);
}

async function getContentTypesXml(doc: VisioDocument): Promise<any> {
    const buf = await doc.save();
    const JSZip = (await import('jszip')).default;
    const zip = await JSZip.loadAsync(buf);
    const xml = await zip.file('[Content_Types].xml')!.async('string');
    return parser.parse(xml);
}

describe('deletePage', () => {
    it('removes the page from doc.pages', async () => {
        const doc = await VisioDocument.create();
        const p2 = await doc.addPage('Page 2');
        expect(doc.pages).toHaveLength(2);

        await doc.deletePage(p2);
        expect(doc.pages).toHaveLength(1);
        expect(doc.pages[0].name).toBe('Page-1');
    });

    it('removes the page XML file from the ZIP', async () => {
        const doc = await VisioDocument.create();
        const p2 = await doc.addPage('Page 2');

        await doc.deletePage(p2);

        const entries = await getZipEntries(doc);
        expect(entries.has('visio/pages/page2.xml')).toBe(false);
        expect(entries.has('visio/pages/page1.xml')).toBe(true);
    });

    it('removes the page entry from pages.xml', async () => {
        const doc = await VisioDocument.create();
        const p2 = await doc.addPage('Page 2');
        await doc.deletePage(p2);

        const parsed = await getPagesXml(doc);
        const pages = parsed.Pages.Page;
        const arr = Array.isArray(pages) ? pages : pages ? [pages] : [];
        expect(arr).toHaveLength(1);
        expect(arr[0]['@_ID']).toBe('1');
    });

    it('removes the relationship from pages.xml.rels', async () => {
        const doc = await VisioDocument.create();
        const p2 = await doc.addPage('Page 2');
        const p2RelId = doc.pages.find(p => p.id === p2.id)?.id; // just to ensure it was added

        await doc.deletePage(p2);

        const parsed = await getPagesRelsXml(doc);
        const rels = parsed.Relationships.Relationship;
        const arr = Array.isArray(rels) ? rels : rels ? [rels] : [];
        // Only the original rId1 (page1) relationship should remain
        expect(arr).toHaveLength(1);
        expect(arr[0]['@_Target']).toBe('page1.xml');
    });

    it('removes the Content Types override for the deleted page', async () => {
        const doc = await VisioDocument.create();
        const p2 = await doc.addPage('Page 2');
        await doc.deletePage(p2);

        const parsed = await getContentTypesXml(doc);
        const overrides = parsed.Types.Override;
        const arr = Array.isArray(overrides) ? overrides : overrides ? [overrides] : [];
        const pageOverrides = arr.filter((o: any) => o['@_PartName']?.includes('page2.xml'));
        expect(pageOverrides).toHaveLength(0);
    });

    it('throws when deleting a page that does not exist', async () => {
        const doc = await VisioDocument.create();
        const p2 = await doc.addPage('Page 2');
        await doc.deletePage(p2);

        // Deleting the same page again should throw
        await expect(doc.deletePage(p2)).rejects.toThrow();
    });

    it('can delete the first page and leave others intact', async () => {
        const doc = await VisioDocument.create();
        await doc.addPage('Page 2');
        const p1 = doc.pages[0];

        await doc.deletePage(p1);

        expect(doc.pages).toHaveLength(1);
        expect(doc.pages[0].name).toBe('Page 2');
    });

    it('removes BackPage reference from foreground pages when background is deleted', async () => {
        const doc = await VisioDocument.create();
        const bg = await doc.addBackgroundPage('Background');
        const fg = doc.pages[0];
        await doc.setBackgroundPage(fg, bg);

        await doc.deletePage(bg);

        const parsed = await getPagesXml(doc);
        const pages = parsed.Pages.Page;
        const arr = Array.isArray(pages) ? pages : pages ? [pages] : [];
        const fgNode = arr.find((n: any) => n['@_ID'] === fg.id);
        expect(fgNode?.['@_BackPage']).toBeUndefined();
    });

    it('can delete multiple pages in sequence', async () => {
        const doc = await VisioDocument.create();
        const p2 = await doc.addPage('Page 2');
        const p3 = await doc.addPage('Page 3');

        await doc.deletePage(p2);
        expect(doc.pages).toHaveLength(2);

        await doc.deletePage(p3);
        expect(doc.pages).toHaveLength(1);

        const entries = await getZipEntries(doc);
        expect(entries.has('visio/pages/page2.xml')).toBe(false);
        expect(entries.has('visio/pages/page3.xml')).toBe(false);
        expect(entries.has('visio/pages/page1.xml')).toBe(true);
    });

    it('does not remove page .rels file if it never existed', async () => {
        // A page with no images/relationships won't have a .rels file;
        // deletePage should succeed without throwing on the missing file.
        const doc = await VisioDocument.create();
        const p2 = await doc.addPage('Clean Page');
        await expect(doc.deletePage(p2)).resolves.not.toThrow();
    });
});
