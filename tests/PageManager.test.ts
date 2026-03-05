import { describe, it, expect, beforeEach, vi } from 'vitest';
import { PageManager } from '../src/core/PageManager';
import { VisioPackage } from '../src/VisioPackage';

describe('PageManager', () => {
    let mockPkg: VisioPackage;
    let manager: PageManager;

    beforeEach(() => {
        mockPkg = {
            getFileText: vi.fn(),
        } as unknown as VisioPackage;
        manager = new PageManager(mockPkg);
    });

    it('should return empty array if pages.xml missing', () => {
        vi.mocked(mockPkg.getFileText).mockImplementation(() => { throw new Error('Not found'); });
        const pages = manager.load();
        expect(pages).toEqual([]);
    });

    it('should parse single page and resolve path from rels', () => {
        const pagesXml = `
            <Pages>
                <Page ID="0" Name="Page-1" r:id="rId1"/>
            </Pages>
        `;
        const relsXml = `
            <Relationships>
                <Relationship Id="rId1" Type="http://..." Target="page1.xml"/>
            </Relationships>
        `;

        vi.mocked(mockPkg.getFileText).mockImplementation((path) => {
            if (path === 'visio/pages/pages.xml') return pagesXml;
            if (path === 'visio/pages/_rels/pages.xml.rels') return relsXml;
            throw new Error(`Unexpected path: ${path}`);
        });

        const pages = manager.load();

        expect(pages).toHaveLength(1);
        expect(pages[0]).toEqual({
            id: '0',
            name: 'Page-1',
            relId: 'rId1',
            xmlPath: 'visio/pages/page1.xml',
            isBackground: false,
            backPageId: undefined
        });
    });

    it('should parse backPageId as string, not number (BUG 6)', () => {
        const pagesXml = `
            <Pages>
                <Page ID="1" Name="Page-1" BackPage="2" r:id="rId1"/>
                <Page ID="2" Name="BG" Background="1" r:id="rId2"/>
            </Pages>
        `;
        const relsXml = `
            <Relationships>
                <Relationship Id="rId1" Type="http://..." Target="page1.xml"/>
                <Relationship Id="rId2" Type="http://..." Target="page2.xml"/>
            </Relationships>
        `;

        vi.mocked(mockPkg.getFileText).mockImplementation((path) => {
            if (path === 'visio/pages/pages.xml') return pagesXml;
            if (path === 'visio/pages/_rels/pages.xml.rels') return relsXml;
            throw new Error(`Unexpected path: ${path}`);
        });

        const pages = manager.load();
        const fgPage = pages.find(p => p.name === 'Page-1')!;

        // Must be a string, not a number, to match VisioPage.backPageId type
        expect(fgPage.backPageId).toBe('2');
        expect(typeof fgPage.backPageId).toBe('string');
    });

    it('Bug 2: page id is always a string, never a number', () => {
        const pagesXml = `
            <Pages>
                <Page ID="1" Name="First" r:id="rId1"/>
                <Page ID="2" Name="Second" r:id="rId2"/>
            </Pages>
        `;
        const relsXml = `
            <Relationships>
                <Relationship Id="rId1" Type="http://..." Target="page1.xml"/>
                <Relationship Id="rId2" Type="http://..." Target="page2.xml"/>
            </Relationships>
        `;

        vi.mocked(mockPkg.getFileText).mockImplementation((path) => {
            if (path === 'visio/pages/pages.xml') return pagesXml;
            if (path === 'visio/pages/_rels/pages.xml.rels') return relsXml;
            throw new Error(`Unexpected path: ${path}`);
        });

        const pages = manager.load();

        // IDs must be strings, not numbers, so strict equality comparisons work
        expect(typeof pages[0].id).toBe('string');
        expect(typeof pages[1].id).toBe('string');
        expect(pages[0].id).toBe('1');
        expect(pages[1].id).toBe('2');
        // Strict equality must hold (not just == coercion)
        expect(pages[0].id === '1').toBe(true);
        expect(pages[0].id === 1 as unknown as string).toBe(false);
    });

    it('Bug 2: deletePage uses strict string equality for lookup', async () => {
        const pagesXml = `
            <Pages>
                <Page ID="3" Name="Only" r:id="rId3"/>
            </Pages>
        `;
        const relsXml = `
            <Relationships>
                <Relationship Id="rId3" Type="http://..." Target="page3.xml"/>
            </Relationships>
        `;
        const ctXml = `<Types><Override PartName="/visio/pages/page3.xml" ContentType="x"/></Types>`;

        vi.mocked(mockPkg.getFileText).mockImplementation((path) => {
            if (path === 'visio/pages/pages.xml') return pagesXml;
            if (path === 'visio/pages/_rels/pages.xml.rels') return relsXml;
            if (path === '[Content_Types].xml') return ctXml;
            throw new Error(`Unexpected path: ${path}`);
        });
        const removedFiles: string[] = [];
        const updatedFiles: string[] = [];
        (mockPkg as any).removeFile = (p: string) => removedFiles.push(p);
        (mockPkg as any).updateFile = (p: string) => updatedFiles.push(p);

        // Passing a string ID that matches '3' must succeed (no type coercion needed)
        await expect(manager.deletePage('3')).resolves.toBeUndefined();
        expect(removedFiles).toContain('visio/pages/page3.xml');
    });

    it('should parse multiple pages', () => {
        const pagesXml = `
            <Pages>
                <Page ID="0" Name="Page-1" r:id="rId1"/>
                <Page ID="4" Name="Details" r:id="rId2"/>
            </Pages>
        `;
        const relsXml = `
            <Relationships>
                <Relationship Id="rId1" Type="http://..." Target="page1.xml"/>
                <Relationship Id="rId2" Type="http://..." Target="page2.xml"/>
            </Relationships>
        `;

        vi.mocked(mockPkg.getFileText).mockImplementation((path) => {
            if (path.includes('pages.xml.rels')) return relsXml;
            return pagesXml;
        });

        const pages = manager.load();

        expect(pages).toHaveLength(2);
        expect(pages[1].name).toBe('Details');
        expect(pages[1].xmlPath).toBe('visio/pages/page2.xml');
    });
});
