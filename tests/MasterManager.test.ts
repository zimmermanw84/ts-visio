import { describe, it, expect, beforeEach, vi } from 'vitest';
import { MasterManager } from '../src/core/MasterManager';
import { VisioPackage } from '../src/VisioPackage';

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
