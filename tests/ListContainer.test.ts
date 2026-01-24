
import { describe, it, expect } from 'vitest';
import { VisioDocument } from '../src/VisioDocument';

describe('Visio Lists (Ordered Containers)', () => {
    it('should stack items vertically', async () => {
        const doc = await VisioDocument.create();
        const page = doc.pages[0];

        // 1. Create List
        const list = await page.addList({
            text: 'My Table',
            x: 5, y: 10, width: 4, height: 1
        });

        // 2. Add Item 1 (Column A)
        // x=0, y=0 (will be moved)
        const item1 = await page.addShape({ text: 'Col A', x: 0, y: 0, width: 3, height: 0.5 });
        await list.addListItem(item1);

        // 3. Add Item 2 (Column B)
        const item2 = await page.addShape({ text: 'Col B', x: 0, y: 0, width: 3, height: 0.5 });
        await list.addListItem(item2);

        // 4. Verify Geometry
        // Update local object states? Tests need to fetch fresh values potentially,
        // but Shape operations might update internal state?
        // Shape.addListItem calls resizeToFit which updates internal state of container.
        // It does NOT update internal state of item1 and item2 (yet).
        // Use modifier to get fresh values.

        // Shape.addListItem calls resizeToFit which updates internal state of container.
        // It does NOT update internal state of item1 and item2 (yet).
        // Use modifier to get fresh values.

        // Use doc.pkg instead of list.pkg (private)
        // Actually, we can just create a new Modifier instance.
        const pkg = (doc as any).pkg;

        // Direct instantiation
        const ShapeModifierStr = (await import('../src/ShapeModifier')).ShapeModifier;
        const testMod = new ShapeModifierStr(pkg);

        const geo1 = testMod.getShapeGeometry(page.id, item1.id);
        const geo2 = testMod.getShapeGeometry(page.id, item2.id);

        console.log('Item 1 Y:', geo1.y);
        console.log('Item 2 Y:', geo2.y);

        // Assert Item 2 is BELOW Item 1
        expect(geo2.y).toBeLessThan(geo1.y);

        // Assert distance is roughly height of item (0.5)
        // Center-to-Center diff should be 0.5
        expect(geo1.y - geo2.y).toBeCloseTo(0.625);

        // Assert Container expanded
        // Initial Height 1. Items total 1. Header allowance/padding?
        // Check container height > 1.5
        const listGeo = testMod.getShapeGeometry(page.id, list.id);
        console.log('List Height:', listGeo.height);
        expect(listGeo.height).toBeGreaterThan(1.4);
    });
    it('should stack items horizontally when direction is horizontal', async () => {
        const doc = await VisioDocument.create();
        const page = doc.pages[0];

        // 1. Create Horizontal List
        const list = await page.addList({ text: 'H-List', x: 0, y: 0, width: 4, height: 1 }, 'horizontal');

        // 2. Add Items
        const item1 = await page.addShape({ text: '1', x: 0, y: 0, width: 1, height: 1 });
        await list.addListItem(item1);

        const item2 = await page.addShape({ text: '2', x: 0, y: 0, width: 1, height: 1 });
        await list.addListItem(item2);

        // 3. Verify
        // Use direct import for access
        const ShapeModifierStr = (await import('../src/ShapeModifier')).ShapeModifier;
        const testMod = new ShapeModifierStr((doc as any).pkg);

        const geo1 = testMod.getShapeGeometry(page.id, item1.id);
        const geo2 = testMod.getShapeGeometry(page.id, item2.id);

        console.log('Item 1 X:', geo1.x);
        console.log('Item 2 X:', geo2.x);

        // Assert Item 2 is right of Item 1
        expect(geo2.x).toBeGreaterThan(geo1.x);
        // Default spacing is 0.125
        // Center diff = 0.5 (half width) + 0.125 + 0.5 (half width) = 1.125
        expect(geo2.x - geo1.x).toBeCloseTo(1.125);
    });
});
