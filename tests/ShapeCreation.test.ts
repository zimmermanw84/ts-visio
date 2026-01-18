import { describe, it, expect } from 'vitest';
import { VisioPackage } from '../src/VisioPackage';
import { ShapeModifier } from '../src/ShapeModifier';
import { ShapeReader } from '../src/ShapeReader';

describe('Shape Creation', () => {
    it('should add a shape to a blank page', async () => {
        const pkg = await VisioPackage.create();
        const modifier = new ShapeModifier(pkg);
        const reader = new ShapeReader(pkg);

        // Add a shape
        const id = await modifier.addShape('1', {
            text: 'Hello World',
            x: 5,
            y: 5,
            width: 2,
            height: 1
        });

        // Verify it exists in memory
        const shapes = reader.readShapes('visio/pages/page1.xml');
        expect(shapes).toHaveLength(1);

        const shape = shapes[0];
        expect(shape.ID).toBe(id);
        expect(shape.Text).toBe('Hello World');
        expect(Number(shape.Cells['PinX'].V)).toBe(5);
        expect(shape.Cells['Width'].V).toBe('2'); // Note: parser typically returns strings for values

        // Verify Geometry exists
        expect(shape.Sections['Geometry']).toBeDefined();
        // 5 rows: MoveTo + 4 LineTo
        expect(shape.Sections['Geometry'].Rows).toHaveLength(5);
    });

    it('should handle sequential ID generation', async () => {
        const pkg = await VisioPackage.create();
        const modifier = new ShapeModifier(pkg);

        const id1 = await modifier.addShape('1', { text: '1', x: 0, y: 0, width: 1, height: 1 });
        const id2 = await modifier.addShape('1', { text: '2', x: 0, y: 0, width: 1, height: 1 });

        expect(id1).toBe('1');
        expect(id2).toBe('2');
    });
});
