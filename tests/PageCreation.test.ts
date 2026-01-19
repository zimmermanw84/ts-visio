import { describe, it, expect, beforeEach, vi } from 'vitest';
import { PageManager } from '../src/core/PageManager';
import { VisioPackage } from '../src/VisioPackage';
import { XMLParser } from 'fast-xml-parser';

describe('PageManager Creation', () => {
    let mockPkg: VisioPackage;
    let manager: PageManager;
    let parser: XMLParser;

    // Initial State Mocks
    const initialPagesXml = `<Pages><Page ID="0" Name="Page-1" r:id="rId1"/></Pages>`;
    const initialRelsXml = `<Relationships><Relationship Id="rId1" Type="http://..." Target="page1.xml"/></Relationships>`;
    const initialContentTypes = `<Types><Override PartName="/visio/pages/page1.xml" ContentType="application/vnd.ms-visio.page+xml"/></Types>`;

    beforeEach(() => {
        mockPkg = {
            getFileText: vi.fn(),
            updateFile: vi.fn(),
        } as unknown as VisioPackage;
        manager = new PageManager(mockPkg);
        parser = new XMLParser({ ignoreAttributes: false, attributeNamePrefix: "@_" });

        vi.mocked(mockPkg.getFileText).mockImplementation((path) => {
            if (path === 'visio/pages/pages.xml') return initialPagesXml;
            if (path === 'visio/pages/_rels/pages.xml.rels') return initialRelsXml;
            if (path === '[Content_Types].xml') return initialContentTypes;
            throw new Error(`Unexpected path: ${path}`);
        });
    });

    it('should create new page and update all 4 files', async () => {
        const newId = await manager.createPage('New Page');

        expect(newId).toBe('1'); // 0 exists, so 1 is next

        const calls = vi.mocked(mockPkg.updateFile).mock.calls;

        // 1. Check physical file creation
        const pageFileCall = calls.find(c => c[0] === 'visio/pages/page1.xml');
        expect(pageFileCall).toBeDefined();
        // content should contain PageContents and PageSheet
        expect(pageFileCall![1]).toContain('<PageContents');
        expect(pageFileCall![1]).toContain('<PageSheet');
        expect(pageFileCall![1]).toContain('PageWidth');

        // 2. Check Content Types update
        const ctCall = calls.find(c => c[0] === '[Content_Types].xml');
        expect(ctCall).toBeDefined();
        expect(ctCall![1]).toContain('PartName="/visio/pages/page1.xml"');

        // 3. Check Rels update
        const relsCall = calls.find(c => c[0] === 'visio/pages/_rels/pages.xml.rels');
        expect(relsCall).toBeDefined();
        const parsedRels = parser.parse(relsCall![1]);
        const rels = parsedRels.Relationships.Relationship;
        expect(rels).toHaveLength(2); // rId1 + new one
        const newRel = rels.find((r: any) => r['@_Target'] === 'page1.xml'); // Wait, ID 1 -> page1.xml?
        // Original was ID 0 -> page1.xml?
        // Actually, internal ID is separate from filename usually.
        // But assuming ID 0 is page1.xml, ID 1 should be page2.xml or similar?
        // Let's see implementation. ID 0 is standard Visio default.
    });
});
