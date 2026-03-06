import { describe, it, expect, beforeEach, vi } from 'vitest';
import { RelsManager } from '../src/core/RelsManager';
import { VisioPackage } from '../src/VisioPackage';
import { XMLParser } from 'fast-xml-parser';

describe('RelsManager', () => {
    let mockPkg: VisioPackage;
    let manager: RelsManager;
    let parser: XMLParser;

    beforeEach(() => {
        mockPkg = {
            getFileText: vi.fn(),
            updateFile: vi.fn(),
        } as unknown as VisioPackage;
        manager = new RelsManager(mockPkg);
        parser = new XMLParser({ ignoreAttributes: false, attributeNamePrefix: "@_" });
    });

    it('should calculate correct _rels path', () => {
        // Access private method via casting if needed, or imply via ensureRelationship behavior
        // We'll test behavior via ensureRelationship calls
    });

    it('should create new .rels file if missing and add relationship', async () => {
        vi.mocked(mockPkg.getFileText).mockImplementation(() => { throw new Error('Not found'); });

        const rId = await manager.ensureRelationship('visio/pages/page1.xml', 'masters/masters.xml', 'http://schemas.microsoft.com/visio/2010/relationships/masters');

        expect(rId).toBe('rId1');

        const updateCall = vi.mocked(mockPkg.updateFile).mock.calls[0];
        expect(updateCall[0]).toBe('visio/pages/_rels/page1.xml.rels');

        const newXml = updateCall[1];
        expect(newXml).toMatch(/^<\?xml version="1\.0"/);
        const parsed = parser.parse(newXml);
        expect(parsed.Relationships.Relationship['@_Target']).toBe('masters/masters.xml');
    });

    it('should reuse existing relationship if target and type match', async () => {
        const mockRels = `
            <Relationships>
                <Relationship Id="rId5" Type="http://schemas.microsoft.com/visio/2010/relationships/masters" Target="masters/masters.xml"/>
            </Relationships>
        `;
        vi.mocked(mockPkg.getFileText).mockReturnValue(mockRels);

        const rId = await manager.ensureRelationship('visio/pages/page1.xml', 'masters/masters.xml', 'http://schemas.microsoft.com/visio/2010/relationships/masters');

        expect(rId).toBe('rId5');
        expect(mockPkg.updateFile).not.toHaveBeenCalled();
    });

    it('should generate new rId if existing relationships but different target', async () => {
        const mockRels = `
            <Relationships>
                <Relationship Id="rId1" Type="http://other" Target="other.xml"/>
            </Relationships>
        `;
        vi.mocked(mockPkg.getFileText).mockReturnValue(mockRels);

        const rId = await manager.ensureRelationship('visio/pages/page1.xml', 'masters/masters.xml', 'http://schemas.microsoft.com/visio/2010/relationships/masters');

        expect(rId).toBe('rId2');

        const updateCall = vi.mocked(mockPkg.updateFile).mock.calls[0];
        const newXml = updateCall[1];
        expect(newXml).toMatch(/^<\?xml version="1\.0"/);
        const parsed = parser.parse(newXml);
        const rels = parsed.Relationships.Relationship;
        expect(rels).toHaveLength(2);
    });

    // regression bug-29: fragile rId prefix assumption
    it('should handle non-rId prefixed relationship IDs without collision', async () => {
        const mockRels = `
            <Relationships>
                <Relationship Id="R3" Type="http://other" Target="other.xml"/>
            </Relationships>
        `;
        vi.mocked(mockPkg.getFileText).mockReturnValue(mockRels);

        const rId = await manager.ensureRelationship('visio/pages/page1.xml', 'new.xml', 'http://some/type');

        // Should extract numeric suffix 3 from "R3" and generate rId4, not rId1
        expect(rId).toBe('rId4');
    });

    it('should handle IDs with no numeric suffix by treating them as 0', async () => {
        const mockRels = `
            <Relationships>
                <Relationship Id="rid" Type="http://other" Target="other.xml"/>
            </Relationships>
        `;
        vi.mocked(mockPkg.getFileText).mockReturnValue(mockRels);

        const rId = await manager.ensureRelationship('visio/pages/page1.xml', 'new.xml', 'http://some/type');

        // No numeric suffix → maxId stays 0 → new ID is rId1
        expect(rId).toBe('rId1');
    });
});
