import { describe, it, expect } from 'vitest';
import { VisioPackage } from '../src/VisioPackage';
import { ShapeModifier } from '../src/ShapeModifier';
import { XMLParser } from 'fast-xml-parser';

describe('Integration: Master Drops', () => {
    it('should successfully load a template, drop a master, and save valid structure', async () => {
        // 1. Setup "Template" Package
        const pkg = await VisioPackage.create();

        // Seed dummy Masters index
        const mastersXml = `
            <Masters>
                <Master ID="5" Name="Router" Type="Shape" />
            </Masters>
        `;
        pkg.updateFile('visio/masters/masters.xml', mastersXml);

        // 2. Drop Master
        const modifier = new ShapeModifier(pkg);
        await modifier.addShape('1', {
            text: 'My Router',
            x: 5, y: 5, width: 1, height: 1,
            masterId: '5'
        });

        // 3. Save (simulated) and Inspect
        // We can inspect the internal map directly to avoid actual disk I/O,
        // effectively treating the "saved" state as the current state of the package.

        const parser = new XMLParser({ ignoreAttributes: false, attributeNamePrefix: '@_' });

        // Verify Page XML
        const pageXml = pkg.getFileText('visio/pages/page1.xml');
        const parsedPage = parser.parse(pageXml);
        const shapes = parsedPage.PageContents.Shapes.Shape;

        // Find our shape
        // Note: VisioPackage.create() creates default page with known ID structure, our new one should be at the end
        const shapeList = Array.isArray(shapes) ? shapes : [shapes];
        const routerShape = shapeList.find((s: any) => s['@_Master'] === '5');

        expect(routerShape).toBeDefined();
        // console.log(JSON.stringify(routerShape, null, 2));
        // fast-xml-parser might merge simple text nodes differently depending on options
        const textVal = routerShape.Text['#text'] || routerShape.Text;
        expect(textVal).toBe('My Router');
        // Ensure NO Geometry
        expect(routerShape.Section).toBeUndefined(); // Or empty/missing geometry section

        // Verify Relationships
        const relsXml = pkg.getFileText('visio/pages/_rels/page1.xml.rels');
        const parsedRels = parser.parse(relsXml);
        const rels = parsedRels.Relationships.Relationship;
        const relList = Array.isArray(rels) ? rels : [rels];

        const masterRel = relList.find((r: any) =>
            r['@_Target'] === '../masters/masters.xml' &&
            r['@_Type'] === 'http://schemas.microsoft.com/visio/2010/relationships/masters'
        );

        expect(masterRel).toBeDefined();
        // The ID of the relationship is less important than its existence,
        // as the Shape uses the Master ID (5), not the rId.
    });
});
