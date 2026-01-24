import { describe, it, expect } from 'vitest';
import { VisioDocument } from '../src/VisioDocument';

describe('XML Structure Reproduction', () => {
    it('should place PageSheet before Shapes in Page XML', async () => {
        const doc = await VisioDocument.create();
        const page = doc.pages[0];

        // Add a shape to trigger ShapeModifier logic
        await page.addShape({ text: 'Test Shape', x: 1, y: 1, width: 1, height: 1 });

        // Get the generated XML for page1
        // We can access the package content directly
        const pkg = (doc as any).pkg;
        const pageXml = pkg.getFileText('visio/pages/page1.xml');

        // Check order
        const pageSheetIndex = pageXml.indexOf('<PageSheet');
        const shapesIndex = pageXml.indexOf('<Shapes');

        // Expect PageSheet to exist
        expect(pageSheetIndex).toBeGreaterThan(-1);

        // Expect Shapes to exist
        expect(shapesIndex).toBeGreaterThan(-1);

        // CRITICAL: PageSheet MUST be before Shapes
        expect(pageSheetIndex).toBeLessThan(shapesIndex);
    });

    it('should place PageSheet before Shapes even after adding Hyperlinks', async () => {
        const doc = await VisioDocument.create();
        const page = doc.pages[0];
        const shape = await page.addShape({ text: 'Link Shape', x: 1, y: 1, width: 1, height: 1 });

        await shape.addHyperlink('https://example.com');

        const pkg = (doc as any).pkg;
        const pageXml = pkg.getFileText('visio/pages/page1.xml');

        const pageSheetIndex = pageXml.indexOf('<PageSheet');
        const shapesIndex = pageXml.indexOf('<Shapes');

        expect(pageSheetIndex).toBeGreaterThan(-1);
        expect(shapesIndex).toBeGreaterThan(-1);
        expect(pageSheetIndex).toBeLessThan(shapesIndex);
    });
});
