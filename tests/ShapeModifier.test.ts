import { describe, it, expect, beforeEach } from 'vitest';
import JSZip from 'jszip';
import { VisioPackage } from '../src/VisioPackage';
import { ShapeModifier } from '../src/ShapeModifier';
import { ShapeReader } from '../src/ShapeReader';

describe('ShapeModifier', () => {
    let pkg: VisioPackage;
    let modifier: ShapeModifier;
    let reader: ShapeReader;

    beforeEach(async () => {
        pkg = new VisioPackage();
        const zip = new JSZip();
        zip.file('visio/pages/page1.xml', `
            <PageContents>
                <Shapes>
                    <Shape ID="1" Name="Process" Type="Shape">
                        <Text>Original Text</Text>
                    </Shape>
                    <Shape ID="2" Name="Decision" Type="Shape" />
                </Shapes>
            </PageContents>
        `);

        const buffer = await zip.generateAsync({ type: 'nodebuffer' });
        await pkg.load(buffer);
        modifier = new ShapeModifier(pkg);
        reader = new ShapeReader(pkg);
    });

    it('should update text of an existing shape', async () => {
        await modifier.updateShapeText('1', '1', 'Updated Text');

        // Verify in memory
        const shapes = reader.readShapes('visio/pages/page1.xml');
        expect(shapes[0].Text).toBe('Updated Text');
    });

    it('should add text to a shape that had none', async () => {
        await modifier.updateShapeText('1', '2', 'New Text');

        const shapes = reader.readShapes('visio/pages/page1.xml');
        expect(shapes[1].Text).toBe('New Text');
    });

    it('should throw error if page not found', async () => {
        await expect(modifier.updateShapeText('99', '1', 'Fail')).rejects.toThrow('Could not find page file');
    });

    it('should throw error if shape not found', async () => {
        await expect(modifier.updateShapeText('1', '99', 'Fail')).rejects.toThrow('Shape 99 not found');
    });
});
