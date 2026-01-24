import { describe, it, expect } from 'vitest';
import { VisioPackage } from '../src/VisioPackage';
import { ShapeModifier } from '../src/ShapeModifier';

describe('Performance Benchmark', () => {
    it('should measure time to add multiple shapes (baseline)', async () => {
        const pkg = await VisioPackage.create();
        const modifier = new ShapeModifier(pkg);
        const count = 100;
        const pageId = '1';

        const startTime = performance.now();

        for (let i = 0; i < count; i++) {
            await modifier.addShape(pageId, {
                text: `Shape ${i}`,
                x: i * 0.1,
                y: i * 0.1,
                width: 1,
                height: 1
            });
        }

        const endTime = performance.now();
        console.log(`Baseline: Time to add ${count} shapes: ${(endTime - startTime).toFixed(2)}ms`);
    }, 20000);

    it('should measure time to add multiple shapes (optimized)', async () => {
        const pkg = await VisioPackage.create();
        const modifier = new ShapeModifier(pkg);
        modifier.autoSave = false;
        const count = 100;
        const pageId = '1';

        const startTime = performance.now();

        for (let i = 0; i < count; i++) {
            await modifier.addShape(pageId, {
                text: `Shape ${i}`,
                x: i * 0.1,
                y: i * 0.1,
                width: 1,
                height: 1
            });
        }
        modifier.flush();

        const endTime = performance.now();
        console.log(`Optimized: Time to add ${count} shapes: ${(endTime - startTime).toFixed(2)}ms`);
    }, 20000);

    it('should measure time to add multiple shapes (batch)', async () => {
        const pkg = await VisioPackage.create();
        const modifier = new ShapeModifier(pkg);
        const count = 100;
        const pageId = '1';

        const startTime = performance.now();

        await modifier.batch(async () => {
            for (let i = 0; i < count; i++) {
                await modifier.addShape(pageId, {
                    text: `Shape ${i}`,
                    x: i * 0.1,
                    y: i * 0.1,
                    width: 1,
                    height: 1
                });
            }
        });

        const endTime = performance.now();
        console.log(`Batch: Time to add ${count} shapes: ${(endTime - startTime).toFixed(2)}ms`);
    }, 20000);

    it('should measure time to add multiple shapes (addShapes)', async () => {
        const pkg = await VisioPackage.create();
        const modifier = new ShapeModifier(pkg);
        const count = 100;
        const pageId = '1';

        const shapes = [];
        for (let i = 0; i < count; i++) {
            shapes.push({
                text: `Shape ${i}`,
                x: i * 0.1,
                y: i * 0.1,
                width: 1,
                height: 1
            });
        }

        const startTime = performance.now();

        await modifier.addShapes(pageId, shapes);

        const endTime = performance.now();
        console.log(`addShapes: Time to add ${count} shapes: ${(endTime - startTime).toFixed(2)}ms`);
    }, 20000);
});
