import { describe, it, expect, beforeAll } from 'vitest';
import { VisioDocument } from '../src/index';
import { VisioValidator, ValidationResult } from '../src/core/VisioValidator';
import * as fs from 'fs';
import * as path from 'path';

describe('Visio Package Validation', () => {
    const validator = new VisioValidator();
    const examplesDir = path.join(__dirname, '../examples');

    describe('Validator Basics', () => {
        it('should validate a fresh document', async () => {
            const doc = await VisioDocument.create();
            const result = await validator.validate((doc as any).pkg);

            expect(result.valid).toBe(true);
            expect(result.errors).toHaveLength(0);
        });

        it('should validate document with shapes', async () => {
            const doc = await VisioDocument.create();
            const page = doc.pages[0];

            await page.addShape({ text: 'Box', x: 1, y: 1, width: 2, height: 1 });
            await page.addShape({ text: 'Circle', x: 4, y: 1, width: 1, height: 1 });

            const result = await validator.validate((doc as any).pkg);

            expect(result.valid).toBe(true);
            expect(result.errors).toHaveLength(0);
        });

        it('should validate document with connectors', async () => {
            const doc = await VisioDocument.create();
            const page = doc.pages[0];

            const shape1 = await page.addShape({ text: 'A', x: 1, y: 1, width: 1, height: 1 });
            const shape2 = await page.addShape({ text: 'B', x: 4, y: 1, width: 1, height: 1 });
            await page.connectShapes(shape1, shape2);

            const result = await validator.validate((doc as any).pkg);

            expect(result.valid).toBe(true);
            expect(result.errors).toHaveLength(0);
        });

        it('should validate document with multiple pages', async () => {
            const doc = await VisioDocument.create();
            await doc.addPage('Page 2');
            await doc.addPage('Page 3');

            const result = await validator.validate((doc as any).pkg);

            expect(result.valid).toBe(true);
            expect(result.errors).toHaveLength(0);
        });
    });

    describe('Example Files Validation', () => {
        const exampleFiles = [
            'simple-schema.vsdx',
            'network-topology.vsdx',
            'containers_demo.vsdx',
            'lists_demo.vsdx',
            'hyperlinks_demo.vsdx',
            'layers_demo.vsdx',
            'layers_demo_annotated.vsdx'
        ];

        for (const filename of exampleFiles) {
            it(`should validate ${filename}`, async () => {
                const filePath = path.join(examplesDir, filename);

                // Skip if file doesn't exist (examples may not have been generated)
                if (!fs.existsSync(filePath)) {
                    console.warn(`Skipping ${filename} - file not found`);
                    return;
                }

                const buffer = fs.readFileSync(filePath);
                const doc = await VisioDocument.load(buffer);
                const result = await validator.validate((doc as any).pkg);

                if (!result.valid) {
                    console.error(`Validation errors for ${filename}:`, result.errors);
                }
                if (result.warnings.length > 0) {
                    console.warn(`Validation warnings for ${filename}:`, result.warnings);
                }

                expect(result.valid).toBe(true);
                expect(result.errors).toHaveLength(0);
            });
        }
    });
});
