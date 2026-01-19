import { describe, it, expect, beforeEach, vi } from 'vitest';
import { MasterManager } from '../src/core/MasterManager';
import { VisioPackage } from '../src/VisioPackage';

describe('MasterManager', () => {
    let mockPkg: VisioPackage;
    let manager: MasterManager;

    beforeEach(() => {
        // Mock VisioPackage
        mockPkg = {
            getFileText: vi.fn(),
        } as unknown as VisioPackage;
        manager = new MasterManager(mockPkg);
    });

    it('should return empty array if masters.xml does not exist', () => {
        // Setup mock to throw error simulating missing file
        vi.mocked(mockPkg.getFileText).mockImplementation(() => {
            throw new Error('File not found');
        });

        const masters = manager.load();
        expect(masters).toEqual([]);
    });

    it('should parse masters correctly from XML', () => {
        const mockXml = `
            <Masters>
                <Master ID="1" Name="Rectangle" NameU="Rectangle" Type="Shape" />
                <Master ID="2" Name="Router" NameU="Router_U" Type="Shape" />
            </Masters>
        `;
        vi.mocked(mockPkg.getFileText).mockReturnValue(mockXml);

        const masters = manager.load();

        expect(masters).toHaveLength(2);

        expect(masters[0]).toEqual(expect.objectContaining({
            id: '1',
            name: 'Rectangle',
            nameU: 'Rectangle',
            type: 'Shape'
        }));

        expect(masters[1]).toEqual(expect.objectContaining({
            id: '2',
            name: 'Router',
            nameU: 'Router_U',
            type: 'Shape'
        }));
    });

    it('should handle single master entry (not array)', () => {
        const mockXml = `
            <Masters>
                <Master ID="5" Name="Single" NameU="SingleU" Type="Shape" />
            </Masters>
        `;
        vi.mocked(mockPkg.getFileText).mockReturnValue(mockXml);

        const masters = manager.load();

        expect(masters).toHaveLength(1);
        expect(masters[0].name).toBe('Single');
    });
});
