import { describe, it, expect, beforeEach, vi, type MockedFunction } from 'vitest';
import { MasterManager } from '../src/core/MasterManager';
import { VisioPackage } from '../src/VisioPackage';
import { XMLParser } from 'fast-xml-parser';

describe('MasterManager.load()', () => {
    let mockPkg: VisioPackage;
    let manager: MasterManager;

    beforeEach(() => {
        mockPkg = {
            getFileText: vi.fn(),
        } as unknown as VisioPackage;
        manager = new MasterManager(mockPkg);
    });

    it('should return empty array if masters.xml does not exist', () => {
        vi.mocked(mockPkg.getFileText).mockImplementation(() => {
            throw new Error('File not found');
        });
        expect(manager.load()).toEqual([]);
    });

    it('should parse masters correctly from XML', () => {
        vi.mocked(mockPkg.getFileText).mockImplementation((path: string) => {
            if (path.includes('.rels')) return '<Relationships></Relationships>';
            return `
                <Masters>
                    <Master ID="1" Name="Rectangle" NameU="Rectangle"/>
                    <Master ID="2" Name="Router" NameU="Router_U"/>
                </Masters>`;
        });

        const masters = manager.load();
        expect(masters).toHaveLength(2);
        expect(masters[0]).toMatchObject({ id: '1', name: 'Rectangle', nameU: 'Rectangle' });
        expect(masters[1]).toMatchObject({ id: '2', name: 'Router',    nameU: 'Router_U' });
    });

    it('should return empty array for empty <Masters/> element (BUG 18)', () => {
        vi.mocked(mockPkg.getFileText).mockReturnValue('<Masters/>');
        expect(() => manager.load()).not.toThrow();
        expect(manager.load()).toEqual([]);
    });

    it('should handle single master entry (not array)', () => {
        vi.mocked(mockPkg.getFileText).mockImplementation((path: string) => {
            if (path.includes('.rels')) return '<Relationships></Relationships>';
            return `
                <Masters>
                    <Master ID="5" Name="Single" NameU="SingleU"/>
                </Masters>`;
        });

        const masters = manager.load();
        expect(masters).toHaveLength(1);
        expect(masters[0].name).toBe('Single');
    });

    it('should populate xmlPath from masters.xml.rels', () => {
        vi.mocked(mockPkg.getFileText).mockImplementation((path: string) => {
            if (path.includes('.rels')) {
                return `<Relationships>
                    <Relationship Id="rId1" Type="master" Target="master1.xml"/>
                </Relationships>`;
            }
            return `<Masters>
                <Master ID="1" Name="Box" NameU="Box">
                    <Rel r:id="rId1"/>
                </Master>
            </Masters>`;
        });

        const masters = manager.load();
        expect(masters[0].xmlPath).toBe('visio/masters/master1.xml');
    });
});

// regression bug-23
describe('MasterManager.importFromStencil (regression bug-23)', () => {
    it('should call load() only once before the loop, not once per master', async () => {
        // Build a minimal stencil ZIP with two master entries
        const JSZip = (await import('jszip')).default;
        const stencilZip = new JSZip();
        stencilZip.file('visio/masters/masters.xml', `<?xml version="1.0" encoding="UTF-8"?>
<Masters xmlns="http://schemas.microsoft.com/office/visio/2012/main">
    <Master ID="1" Name="Alpha" NameU="Alpha"><Rel r:id="rId1"/></Master>
    <Master ID="2" Name="Beta"  NameU="Beta" ><Rel r:id="rId2"/></Master>
</Masters>`);
        stencilZip.file('visio/masters/_rels/masters.xml.rels', `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="master" Target="master1.xml"/>
    <Relationship Id="rId2" Type="master" Target="master2.xml"/>
</Relationships>`);
        stencilZip.file('visio/masters/master1.xml', `<?xml version="1.0" encoding="UTF-8"?><MasterContents/>`);
        stencilZip.file('visio/masters/master2.xml', `<?xml version="1.0" encoding="UTF-8"?><MasterContents/>`);
        const stencilBuf = await stencilZip.generateAsync({ type: 'nodebuffer' });

        const files: Record<string, string> = {
            '[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/></Types>`,
            'visio/_rels/document.xml.rels': `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>`,
            'visio/masters/masters.xml': `<?xml version="1.0" encoding="UTF-8"?><Masters xmlns="http://schemas.microsoft.com/office/visio/2012/main"></Masters>`,
            'visio/masters/_rels/masters.xml.rels': `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>`,
        };

        const getFileText = vi.fn((path: string): string => {
            if (path in files) return files[path];
            throw new Error(`File not found: ${path}`);
        });
        const mockPkg = {
            getFileText,
            updateFile: vi.fn((path: string, content: string) => { files[path] = content; }),
        } as unknown as VisioPackage;

        const manager = new MasterManager(mockPkg);
        const loadSpy = vi.spyOn(manager, 'load');

        const imported = await manager.importFromStencil(stencilBuf);

        // Two masters imported with sequential IDs
        expect(imported).toHaveLength(2);
        expect(imported[0].id).toBe('1');
        expect(imported[1].id).toBe('2');

        // With the O(n²) bug: load() called once per master (2 times for 2 masters)
        // With the fix: load() called exactly once before the loop
        expect(loadSpy).toHaveBeenCalledTimes(1);
    });
});

// regression bug-22
describe('MasterManager.addMasterEntry (regression bug-22)', () => {
    it('should not write inline <Shapes> into masters.xml Master entry', () => {
        const parser = new XMLParser({ ignoreAttributes: false, attributeNamePrefix: '@_' });

        // Track what was written to masters.xml
        let writtenMastersXml = '';
        const files: Record<string, string> = {
            '[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/></Types>`,
            'visio/_rels/document.xml.rels': `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>`,
            'visio/masters/masters.xml': `<?xml version="1.0" encoding="UTF-8"?><Masters xmlns="http://schemas.microsoft.com/office/visio/2012/main"></Masters>`,
        };

        const mockPkg = {
            getFileText: vi.fn((path: string): string => {
                if (path in files) return files[path];
                throw new Error(`File not found: ${path}`);
            }),
            updateFile: vi.fn((path: string, content: string) => {
                files[path] = content;
                if (path === 'visio/masters/masters.xml') {
                    writtenMastersXml = content;
                }
            }),
        } as unknown as VisioPackage;

        const manager = new MasterManager(mockPkg);
        // Directly invoke private method via cast
        (manager as any).addMasterEntry(1, 'TestShape', 'TestShape', 'rId1');

        expect(writtenMastersXml).toBeTruthy();
        const parsed = parser.parse(writtenMastersXml);
        const masterEntry = parsed.Masters?.Master;
        const entry = Array.isArray(masterEntry) ? masterEntry[0] : masterEntry;
        expect(entry).toBeDefined();
        // Must NOT contain a Shapes element
        expect(entry.Shapes).toBeUndefined();
        // Must still contain PageSheet and Rel
        expect(entry.PageSheet).toBeDefined();
        expect(entry.Rel).toBeDefined();
    });
});
