
import { describe, it, expect } from 'vitest';
import { VisioDocument } from '../src/VisioDocument';
import { XMLParser } from 'fast-xml-parser';

describe('Visio Container Layout', () => {
    it('should auto-resize container to fit members and fix Z-Order', async () => {
        const doc = await VisioDocument.create();
        const page = doc.pages[0];

        // 1. Create Container (Small initial size)
        const container = await page.addContainer({
            text: 'Auto Container',
            x: 0, y: 0, width: 1, height: 1
        });

        // 2. Create Members (Large distribution)
        // Box 1: Left-Bottom at (2, 2) width 2, height 2 -> Center (3, 3)
        // Box 2: Right-Top at (8, 8) width 2, height 2 -> Center (9, 9)

        // Member 1 Geometry:
        // x=3, y=3, w=2, h=2. Bounds: [2, 4] x [2, 4]
        const box1 = await page.addShape({ text: 'B1', x: 3, y: 3, width: 2, height: 2 });

        // Member 2 Geometry:
        // x=9, y=9, w=2, h=2. Bounds: [8, 10] x [8, 10]
        const box2 = await page.addShape({ text: 'B2', x: 9, y: 9, width: 2, height: 2 });

        await container.addMember(box1);
        await container.addMember(box2);

        // 3. Resize To Fit
        await container.resizeToFit(0.5); // 0.5 padding

        // 4. Verify Geometry
        // Expected Bounds: MinX=2, MaxX=10, MinY=2, MaxY=10
        // Padding 0.5 -> Bounds: [1.5, 10.5] x [1.5, 10.5]
        // Width = 9, Height = 9
        // PinX = 1.5 + 4.5 = 6
        // PinY = 1.5 + 4.5 = 6

        expect(container.width).toBeCloseTo(9);
        expect(container.height).toBeCloseTo(9);
        expect(container.x).toBeCloseTo(6);
        expect(container.y).toBeCloseTo(6);

        // 5. Verify Z-Order (Container should be first in Shapes list = Back)
        const pkg = (doc as any).pkg;
        const pageXml = pkg.getFileText('visio/pages/page1.xml');
        const parser = new XMLParser({ ignoreAttributes: false, attributeNamePrefix: "@_" });
        const parsed = parser.parse(pageXml);

        const shapes = parsed.PageContents.Shapes.Shape;
        expect(shapes[0]['@_ID']).toBe(container.id);
    });
});
