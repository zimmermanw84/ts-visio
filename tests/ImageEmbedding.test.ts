
import { describe, it, expect } from 'vitest';
import { VisioDocument } from '../src/VisioDocument';
import { XMLParser } from 'fast-xml-parser';

describe('Image Embedding (E2E)', () => {
    it('should embedding an image via public API', async () => {
        const doc = await VisioDocument.create();
        const page = doc.pages[0]; // Access first page directly

        const dummyBuffer = Buffer.from([0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A]); // PNG Header
        const x = 5, y = 5, w = 3, h = 2;

        // Act: Add Image
        const imageShape = await page.addImage(dummyBuffer, 'logo.png', x, y, w, h);

        // Assert: Shape Object is returned correctly
        expect(imageShape.id).toBeDefined();

        // Assert: Save and Verify Internals
        const buffer = await doc.save();

        // We can inspect the internal package of the doc directly for verification
        // (Accessing private pkg via cast for testing)
        const pkg = (doc as any).pkg;

        // 1. Verify Media File
        // We can inspect the internal package of the doc directly for verification

        // Let's use the internal files map to check persistence
        const filesMap = (pkg as any).files as Map<string, string | Buffer>;
        const storedBuffer = filesMap.get('visio/media/logo.png');
        expect(Buffer.isBuffer(storedBuffer)).toBe(true);
        expect((storedBuffer as Buffer).equals(dummyBuffer)).toBe(true);

        // 2. Verify Relationship
        const relsXml = (pkg as any).getFileText('visio/pages/_rels/page1.xml.rels');
        const parser = new XMLParser({ ignoreAttributes: false, attributeNamePrefix: "@_" });
        const relsParsed = parser.parse(relsXml);
        const rels = Array.isArray(relsParsed.Relationships.Relationship)
            ? relsParsed.Relationships.Relationship
            : [relsParsed.Relationships.Relationship];

        const imageRel = rels.find((r: any) => r['@_Target'] === '../media/logo.png');
        expect(imageRel).toBeDefined();
        const rId = imageRel['@_Id'];

        // 3. Verify Shape XML
        const pageXml = (pkg as any).getFileText('visio/pages/page1.xml');
        const pageParsed = parser.parse(pageXml);
        const shapes = pageParsed.PageContents.Shapes.Shape;
        const targetShape = Array.isArray(shapes) ? shapes.find((s: any) => s['@_ID'] == imageShape.id) : shapes;

        expect(targetShape).toBeDefined();
        expect(targetShape['@_Type']).toBe('Foreign');
        expect(targetShape.ForeignData['@_r:id']).toBe(rId);

        // 4. Verify Content Types
        const ctXml = (pkg as any).getFileText('[Content_Types].xml');
        expect(ctXml).toContain('Extension="png"');
        expect(ctXml).toContain('ContentType="image/png"');
    });
});
