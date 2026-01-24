import { describe, it, expect } from 'vitest';
import { VisioDocument } from '../src/index';
import { XMLParser } from 'fast-xml-parser';
import { ShapeModifier } from '../src/ShapeModifier';

describe('Edge Case Regression Tests', () => {
    const parser = new XMLParser({ ignoreAttributes: false, attributeNamePrefix: '@_' });

    describe('ConnectorBuilder - Minimal Cell Shapes', () => {
        it('should handle connectors between shapes', async () => {
            const doc = await VisioDocument.create();
            const page = doc.pages[0];

            // Create shapes
            const shape1 = await page.addShape({ text: 'A', x: 1, y: 1, width: 1, height: 1 });
            const shape2 = await page.addShape({ text: 'B', x: 3, y: 1, width: 1, height: 1 });

            // Connect shapes
            await page.connectShapes(shape1, shape2);

            // Verify connector was created
            const pageXml = (doc as any).pkg.getFileText(`visio/pages/page${page.id}.xml`);
            expect(pageXml).toContain('Dynamic connector');
        });
    });

    describe('setPropertyValue - Nested Shapes', () => {
        it('should set property on shape inside a group', async () => {
            const doc = await VisioDocument.create();
            const page = doc.pages[0];

            // Create a group with a child shape
            const group = await page.addShape({ text: 'Group', x: 2, y: 2, width: 3, height: 3, type: 'Group' });
            const child = await page.addShape({ text: 'Child', x: 1, y: 1, width: 1, height: 1 }, group.id);

            // Add property to child and set value using the low-level API
            const modifier = new ShapeModifier((doc as any).pkg);
            modifier.addPropertyDefinition(page.id, child.id, 'Status', 0); // 0 = string
            modifier.setPropertyValue(page.id, child.id, 'Status', 'Active');

            // Verify XML - property should be set on nested shape
            const pageXml = (doc as any).pkg.getFileText(`visio/pages/page${page.id}.xml`);
            const parsed = parser.parse(pageXml);

            // Find child shape
            const topShapes = Array.isArray(parsed.PageContents.Shapes.Shape)
                ? parsed.PageContents.Shapes.Shape
                : [parsed.PageContents.Shapes.Shape];
            const parentShape = topShapes.find((s: any) => s['@_ID'] === group.id);
            const childShapes = parentShape.Shapes?.Shape;
            const childShape = Array.isArray(childShapes) ? childShapes[0] : childShapes;

            // Verify property section exists
            const sections = Array.isArray(childShape.Section) ? childShape.Section : [childShape.Section];
            const propSection = sections.find((s: any) => s['@_N'] === 'Property');
            expect(propSection).toBeDefined();

            // Verify value was set
            const rows = Array.isArray(propSection.Row) ? propSection.Row : [propSection.Row];
            const statusRow = rows.find((r: any) => r['@_N'] === 'Prop.Status');
            expect(statusRow).toBeDefined();
        });
    });

    describe('Arrow Validation', () => {
        it('should default invalid arrow values to 0', async () => {
            const doc = await VisioDocument.create();
            const page = doc.pages[0];

            const shape1 = await page.addShape({ text: 'A', x: 1, y: 1, width: 1, height: 1 });
            const shape2 = await page.addShape({ text: 'B', x: 3, y: 1, width: 1, height: 1 });

            // Pass invalid arrow values (99 is out of range 0-45)
            await page.connectShapes(shape1, shape2, '99', '-5');

            // Verify XML
            const pageXml = (doc as any).pkg.getFileText(`visio/pages/page${page.id}.xml`);
            const parsed = parser.parse(pageXml);

            const shapes = Array.isArray(parsed.PageContents.Shapes.Shape)
                ? parsed.PageContents.Shapes.Shape
                : [parsed.PageContents.Shapes.Shape];
            // Connector is the last shape added
            const connectorShape = shapes[shapes.length - 1];
            const sections = Array.isArray(connectorShape.Section) ? connectorShape.Section : [connectorShape.Section];
            const lineSection = sections.find((s: any) => s['@_N'] === 'Line');
            const cells = Array.isArray(lineSection.Cell) ? lineSection.Cell : [lineSection.Cell];

            const beginArrow = cells.find((c: any) => c['@_N'] === 'BeginArrow');
            const endArrow = cells.find((c: any) => c['@_N'] === 'EndArrow');

            // Should default to 0 for invalid values
            expect(beginArrow['@_V']).toBe('0');
            expect(endArrow['@_V']).toBe('0');
        });

        it('should accept valid arrow values', async () => {
            const doc = await VisioDocument.create();
            const page = doc.pages[0];

            const shape1 = await page.addShape({ text: 'A', x: 1, y: 1, width: 1, height: 1 });
            const shape2 = await page.addShape({ text: 'B', x: 3, y: 1, width: 1, height: 1 });

            // Pass valid arrow values
            await page.connectShapes(shape1, shape2, '5', '13');

            const pageXml = (doc as any).pkg.getFileText(`visio/pages/page${page.id}.xml`);
            const parsed = parser.parse(pageXml);

            const shapes = Array.isArray(parsed.PageContents.Shapes.Shape)
                ? parsed.PageContents.Shapes.Shape
                : [parsed.PageContents.Shapes.Shape];
            const connectorShape = shapes[shapes.length - 1];
            const sections = Array.isArray(connectorShape.Section) ? connectorShape.Section : [connectorShape.Section];
            const lineSection = sections.find((s: any) => s['@_N'] === 'Line');
            const cells = Array.isArray(lineSection.Cell) ? lineSection.Cell : [lineSection.Cell];

            const beginArrow = cells.find((c: any) => c['@_N'] === 'BeginArrow');
            const endArrow = cells.find((c: any) => c['@_N'] === 'EndArrow');

            expect(beginArrow['@_V']).toBe('5');
            expect(endArrow['@_V']).toBe('13');
        });
    });

    describe('Dimension Validation', () => {
        it('should throw error for zero width', async () => {
            const doc = await VisioDocument.create();
            const page = doc.pages[0];

            await expect(
                page.addShape({ text: 'Invalid', x: 1, y: 1, width: 0, height: 1 })
            ).rejects.toThrow('Shape dimensions must be positive numbers');
        });

        it('should throw error for negative height', async () => {
            const doc = await VisioDocument.create();
            const page = doc.pages[0];

            await expect(
                page.addShape({ text: 'Invalid', x: 1, y: 1, width: 1, height: -1 })
            ).rejects.toThrow('Shape dimensions must be positive numbers');
        });
    });

    describe('Hyperlink XML Escaping', () => {
        it('should escape ampersands in URLs', async () => {
            const doc = await VisioDocument.create();
            const page = doc.pages[0];
            const shape = await page.addShape({ text: 'Link', x: 1, y: 1, width: 1, height: 1 });

            // URL with ampersands that need XML escaping
            await shape.toUrl('https://example.com?foo=1&bar=2');

            const pageXml = (doc as any).pkg.getFileText(`visio/pages/page${page.id}.xml`);

            // Should contain escaped ampersands
            expect(pageXml).toContain('&amp;');
        });

        it('should escape angle brackets and quotes in URLs', async () => {
            const doc = await VisioDocument.create();
            const page = doc.pages[0];
            const shape = await page.addShape({ text: 'Link', x: 1, y: 1, width: 1, height: 1 });

            // Use ShapeModifier directly to test address escaping
            const modifier = new ShapeModifier((doc as any).pkg);
            await modifier.addHyperlink(page.id, shape.id, {
                address: 'https://example.com?q=<test>&b="val"',
                description: 'Test'
            });

            const pageXml = (doc as any).pkg.getFileText(`visio/pages/page${page.id}.xml`);

            // Should contain escaped characters
            expect(pageXml).toContain('&lt;');  // <
            expect(pageXml).toContain('&gt;');  // >
            expect(pageXml).toContain('&quot;'); // "
            expect(pageXml).toContain('&amp;');  // &
        });
    });
});
