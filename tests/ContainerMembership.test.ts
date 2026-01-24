
import { describe, it, expect } from 'vitest';
import { VisioDocument } from '../src/VisioDocument';
import { XMLParser } from 'fast-xml-parser';

describe('Visio Container Membership', () => {
    it('should link shapes via Relationship tag', async () => {
        const doc = await VisioDocument.create();
        const page = doc.pages[0];

        // 1. Create Container
        const container = await page.addContainer({
            text: 'Container',
            x: 1, y: 1, width: 5, height: 5
        });

        // 2. Create Member
        const box = await page.addShape({
            text: 'Member Box',
            x: 2, y: 2, width: 1, height: 1
        });

        // 3. Add Member
        await container.addMember(box);

        // 4. Verify XML
        const buffer = await doc.save();
        const pkg = (doc as any).pkg;
        const pageXml = pkg.getFileText('visio/pages/page1.xml');

        const parser = new XMLParser({ ignoreAttributes: false, attributeNamePrefix: "@_" });
        const parsed = parser.parse(pageXml);

        // Check Relationships
        const rels = parsed.PageContents.Relationships.Relationship;
        expect(rels).toBeDefined();

        const relationships = Array.isArray(rels) ? rels : [rels];
        const link = relationships.find((r: any) =>
            r['@_Type'] === 'Container' &&
            r['@_ShapeID'] === container.id &&
            r['@_RelatedShapeID'] === box.id
        );

        expect(link).toBeDefined();
        console.log('Verified Container Relationship:', link);
    });
});
