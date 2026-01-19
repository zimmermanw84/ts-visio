import { describe, it, expect } from 'vitest';
import { VisioPackage } from '../src/VisioPackage';
import { ShapeModifier } from '../src/ShapeModifier';
import { VisioDocument } from '../src/VisioDocument';
import { XMLParser } from 'fast-xml-parser';

describe('Integration Tests', () => {
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
        const parser = new XMLParser({ ignoreAttributes: false, attributeNamePrefix: '@_' });

        // Verify Page XML
        const pageXml = pkg.getFileText('visio/pages/page1.xml');
        const parsedPage = parser.parse(pageXml);
        const shapes = parsedPage.PageContents.Shapes.Shape;

        // Find our shape
        const shapeList = Array.isArray(shapes) ? shapes : [shapes];
        const routerShape = shapeList.find((s: any) => s['@_Master'] === '5');

        expect(routerShape).toBeDefined();
        const textVal = routerShape.Text['#text'] || routerShape.Text;
        expect(textVal).toBe('My Router');
        expect(routerShape.Section).toBeUndefined();

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
    });

    it('should support multi-page creation and persistence', async () => {
        const doc = await VisioDocument.create();

        // Default page should exist
        expect(doc.pages).toHaveLength(1);
        expect(doc.pages[0].id).toBe('1'); // Template default start ID is 1

        // Add New Page
        const newId = await doc.addPage('Analysis Page');
        expect(newId).toBe('2');

        // Refresh properties
        expect(doc.pages).toHaveLength(2);
        const p2 = doc.pages.find(p => p.id === newId);
        expect(p2).toBeDefined();
        expect(p2!.name).toBe('Analysis Page');

        // Save and Reload
        const saved = await doc.save();
        const reloaded = await VisioDocument.load(saved);

        expect(reloaded.pages).toHaveLength(2);
        const reloadedP2 = reloaded.pages.find(p => p.name === 'Analysis Page');
        expect(reloadedP2).toBeDefined();
        expect(reloadedP2).toBeDefined();
        expect(reloadedP2!.id).toBe('2');
    });
});
