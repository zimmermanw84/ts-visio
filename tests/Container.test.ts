
import { describe, it, expect } from 'vitest';
import { VisioDocument } from '../src/VisioDocument';
import { XMLParser } from 'fast-xml-parser';
import { ContainerBuilder } from '../src/shapes/ContainerBuilder';

describe('Visio Containers', () => {
    it('should create a container shape with correct metadata', async () => {
        const doc = await VisioDocument.create();
        const page = doc.pages[0];

        const containerShape = await page.addContainer({
            text: 'My Container',
            x: 1, y: 1, width: 4, height: 3
        });
        const containerId = containerShape.id;

        expect(containerId).toBeDefined();

        // Save and inspect XML
        const buffer = await doc.save();
        const pkg = (doc as any).pkg;
        const pageXml = pkg.getFileText('visio/pages/page1.xml');

        const parser = new XMLParser({ ignoreAttributes: false, attributeNamePrefix: "@_" });
        const parsed = parser.parse(pageXml);

        const shapes = parsed.PageContents.Shapes.Shape;
        const container = Array.isArray(shapes) ? shapes.find((s: any) => s['@_ID'] === containerId) : shapes;

        expect(container).toBeDefined();

        // precise container matching
        if (Array.isArray(container)) throw new Error("Found multiple shapes matching ID?");

        // Check User Section for msvStructureType = "Container"
        const userSection = container.Section.find((s: any) => s['@_N'] === 'User');
        expect(userSection).toBeDefined();

        const structType = userSection.Row.find((r: any) => r['@_N'] === 'msvStructureType');
        expect(structType).toBeDefined();
        expect(structType.Cell['@_V']).toBe('"Container"');

        // Check TextXform
        const textXform = container.Section.find((s: any) => s['@_N'] === 'TextXform');
        expect(textXform).toBeDefined();
        // Just verify it exists for now, detailed cell checks might be brittle if we tweak formulas
    });

    it('should robustly handle shapes with single-object sections/rows (XML quirk)', () => {
        // Mock a shape structure that mimics "fast-xml-parser" output when only 1 item exists (no array)
        const quirkyShape = {
            Section: {
                '@_N': 'User',
                // Non-array Row
                Row: { '@_N': 'ExistingProperty', Cell: [] }
            }
        };



        // Should not throw
        ContainerBuilder.makeContainer(quirkyShape);

        // Verification: Should have normalized to Arrays
        expect(Array.isArray(quirkyShape.Section)).toBe(true);
        const userSection = (quirkyShape.Section as unknown as any[]).find((s: any) => s['@_N'] === 'User');
        expect(Array.isArray(userSection.Row)).toBe(true);

        // Should still have successfully added the container metadata
        const typeRow = userSection.Row.find((r: any) => r['@_N'] === 'msvStructureType');
        expect(typeRow).toBeDefined();
    });
});
