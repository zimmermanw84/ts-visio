import { describe, it, expect, beforeEach } from 'vitest';
import JSZip from 'jszip';
import { VisioPackage } from '../src/VisioPackage';
import { ShapeModifier } from '../src/ShapeModifier';
import { ShapeReader } from '../src/ShapeReader';

describe('ShapeModifier AutoSave', () => {
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
                </Shapes>
            </PageContents>
        `);

        const buffer = await zip.generateAsync({ type: 'nodebuffer' });
        await pkg.load(buffer);
        modifier = new ShapeModifier(pkg);
        reader = new ShapeReader(pkg);
    });

    it('should NOT update package content immediately when autoSave is false', async () => {
        modifier.autoSave = false;
        await modifier.addShape('1', { text: 'New Shape', x: 0, y: 0, width: 1, height: 1 });

        // Check if package still has original content (only 1 shape)
        const content = pkg.getFileText('visio/pages/page1.xml');
        // Simple check: content should NOT contain "New Shape"
        expect(content).not.toContain('New Shape');

        // Reader reads from pkg, so it should not see it
        const shapes = reader.readShapes('visio/pages/page1.xml');
        expect(shapes.length).toBe(1);
    });

    it('should update package content after flush', async () => {
        modifier.autoSave = false;
        await modifier.addShape('1', { text: 'New Shape', x: 0, y: 0, width: 1, height: 1 });

        modifier.flush();

        const content = pkg.getFileText('visio/pages/page1.xml');
        expect(content).toContain('New Shape');

        const shapes = reader.readShapes('visio/pages/page1.xml');
        expect(shapes.length).toBe(2);
    });

    it('should handle multiple updates in batch', async () => {
        modifier.autoSave = false;
        await modifier.addShape('1', { text: 'Shape 2', x: 0, y: 0, width: 1, height: 1 });
        await modifier.addShape('1', { text: 'Shape 3', x: 10, y: 10, width: 1, height: 1 });
        await modifier.updateShapeText('1', '1', 'Modified Original');

        // Verify nothing in package yet
        let content = pkg.getFileText('visio/pages/page1.xml');
        expect(content).not.toContain('Shape 2');
        expect(content).not.toContain('Shape 3');
        expect(content).toContain('Original Text');

        modifier.flush();

        content = pkg.getFileText('visio/pages/page1.xml');
        expect(content).toContain('Shape 2');
        expect(content).toContain('Shape 3');
        expect(content).toContain('Modified Original');
    });
});
