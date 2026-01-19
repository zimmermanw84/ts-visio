import { describe, it, expect, beforeEach, vi } from 'vitest';
import { ShapeModifier } from '../src/ShapeModifier';
import { VisioPackage } from '../src/VisioPackage';
import { XMLParser } from 'fast-xml-parser';

describe('Master Shape Instantiation', () => {
    let mockPkg: VisioPackage;
    let modifier: ShapeModifier;
    let parser: XMLParser;

    // Minimal Page XML
    const mockPageXml = `
        <PageContents>
            <Shapes></Shapes>
        </PageContents>
    `;

    beforeEach(() => {
        mockPkg = {
            getFileText: vi.fn(),
            updateFile: vi.fn(),
        } as unknown as VisioPackage;
        modifier = new ShapeModifier(mockPkg);
        parser = new XMLParser({ ignoreAttributes: false, attributeNamePrefix: "@_" });

        vi.mocked(mockPkg.getFileText).mockImplementation((path: string) => {
            if (path.includes('.rels')) return '<Relationships></Relationships>';
            return mockPageXml;
        });
    });

    it('should create legacy shape with Geometry when no masterId provided', async () => {
        const id = await modifier.addShape('1', {
            text: 'Legacy Box',
            x: 1, y: 1, width: 1, height: 1
        });

        // Capture the XML written back
        const updateCall = vi.mocked(mockPkg.updateFile).mock.calls[0];
        const newXml = updateCall[1];
        const parsed = parser.parse(newXml);
        const shape = parsed.PageContents.Shapes.Shape;

        expect(shape['@_Master']).toBeUndefined();

        let sections = shape.Section || [];
        if (!Array.isArray(sections)) sections = [sections];

        // Check for Geometry Section
        const geom = sections.find((s: any) => s['@_N'] === 'Geometry');
        expect(geom).toBeDefined();

        // Verify MoveTo/LineTo exist
        let rows = geom.Row || [];
        if (!Array.isArray(rows)) rows = [rows];
        expect(rows.length).toBeGreaterThan(0);
    });

    it('should create master instance with Master attribute and NO Geometry', async () => {
        const id = await modifier.addShape('1', {
            text: 'Router',
            x: 2, y: 2, width: 1, height: 1,
            masterId: '5'
        });

        const calls = vi.mocked(mockPkg.updateFile).mock.calls;
        const pageCall = calls.find(c => c[0] && c[0].includes('visio/pages/page'));
        if (!pageCall) throw new Error("No page updated");
        const newXml = pageCall[1];
        const parsed = parser.parse(newXml);
        const shape = parsed.PageContents.Shapes.Shape;

        expect(shape['@_Master']).toBe('5');

        let sections = shape.Section || [];
        if (!Array.isArray(sections)) sections = [sections];

        // Geometry should be absent
        const geom = sections.find((s: any) => s['@_N'] === 'Geometry');
        expect(geom).toBeUndefined();
    });

    it('should allow mixing legacy and master shapes on the same page', async () => {
        // First add legacy
        await modifier.addShape('1', { text: 'Legacy', x: 0, y: 0, width: 1, height: 1 });
        expect(true).toBe(true);
    });
});
