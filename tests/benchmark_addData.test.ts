import { describe, it } from 'vitest';
import { VisioPackage } from '../src/VisioPackage';
import { VisioDocument } from '../src/VisioDocument';
import { VisioPropType } from '../src/types/VisioTypes';

describe('Shape.addData Performance Benchmark', () => {
    it('should measure time to add multiple properties to a shape', async () => {
        // Create document and page
        const doc = await VisioDocument.create();
        const page = await doc.addPage('BenchmarkPage');
        const shape = await page.addShape({
            text: 'Test Shape',
            x: 1,
            y: 1,
            width: 1,
            height: 1
        });

        const count = 50;
        const startTime = performance.now();

        for (let i = 0; i < count; i++) {
            shape.addData(`prop${i}`, { value: `value${i}`, type: VisioPropType.String });
        }

        const endTime = performance.now();
        console.log(`Time to add ${count} properties: ${(endTime - startTime).toFixed(2)}ms`);
    }, 30000);
});
