
import { describe, it, expect } from 'vitest';
import { VisioPackage } from '../src/VisioPackage';
import { MediaManager } from '../src/core/MediaManager';
import { RelsManager } from '../src/core/RelsManager';
import { ShapeModifier } from '../src/ShapeModifier';
import { XMLParser } from 'fast-xml-parser';

describe('Image Shape Structure (Phase 2)', () => {
    it('should generate correct Foreign shape XML', async () => {
        const pkg = await VisioPackage.create();
        const media = new MediaManager(pkg);
        const rels = new RelsManager(pkg);
        const shapes = new ShapeModifier(pkg);

        // 1. Add Media
        const dummyBuffer = Buffer.from([0x89, 0x50, 0x4E, 0x47]);
        const mediaPath = media.addMedia('test.png', dummyBuffer); // returns "../media/test.png"

        // 2. Add Relationship
        // Target should be relative to page folder "visio/pages/" -> "../media/test.png" matches what we got
        const rId = await rels.addPageImageRel('1', mediaPath);

        // 3. Add Shape
        const shapeId = await shapes.addShape('1', { // Page 1
            text: 'Image Shape',
            x: 5, y: 5, width: 2, height: 1,
            type: 'Foreign',
            imgRelId: rId
        });

        // 4. Verify XML
        const pageXml = pkg.getFileText('visio/pages/page1.xml');
        const parser = new XMLParser({ ignoreAttributes: false, attributeNamePrefix: "@_" });
        const parsed = parser.parse(pageXml);

        const allShapes = parsed.PageContents.Shapes.Shape;
        const shp = Array.isArray(allShapes) ? allShapes.find((s: any) => s['@_ID'] == shapeId) : allShapes;

        expect(shp).toBeDefined();
        expect(shp['@_Type']).toBe('Foreign');
        expect(shp.ForeignData).toBeDefined();
        expect(shp.ForeignData['@_r:id']).toBe(rId);

        // 5. Verify Rel Exists
        const relsXml = pkg.getFileText('visio/pages/_rels/page1.xml.rels');
        const relsParsed = parser.parse(relsXml);
        const relations = Array.isArray(relsParsed.Relationships.Relationship)
            ? relsParsed.Relationships.Relationship
            : [relsParsed.Relationships.Relationship];

        const rel = relations.find((r: any) => r['@_Id'] === rId);
        expect(rel).toBeDefined();
        expect(rel['@_Target']).toBe(mediaPath);
        expect(rel['@_Type']).toContain('relationships/image');

        // 6. Verify Default Styling (LinePattern=0, Geometry NoFill=1)
        const lineSection = shp.Section.find((s: any) => s['@_N'] === 'Line');
        expect(lineSection).toBeDefined();
        const linePattern = lineSection.Cell.find((c: any) => c['@_N'] === 'LinePattern');
        expect(linePattern['@_V']).toBe('0'); // Invisible Border

        const geomSection = shp.Section.find((s: any) => s['@_N'] === 'Geometry');
        expect(geomSection).toBeDefined();
        const geomCells = Array.isArray(geomSection.Cell) ? geomSection.Cell : [geomSection.Cell];
        const noFill = geomCells.find((c: any) => c['@_N'] === 'NoFill');
        expect(noFill['@_V']).toBe('1'); // No Fill
    });
});
