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
            id: 0,
            name: 'Page-1',
            relId: 'rId1',
            xmlPath: 'visio/pages/page1.xml',
            isBackground: false,
            backPageId: undefined
        });
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
