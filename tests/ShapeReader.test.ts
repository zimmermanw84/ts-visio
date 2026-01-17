import { describe, it, expect, beforeEach } from 'vitest';
import JSZip from 'jszip';
import { VisioPackage } from '../src/VisioPackage';
import { ShapeReader } from '../src/ShapeReader';

describe('ShapeReader', () => {
    let pkg: VisioPackage;
    let reader: ShapeReader;

    beforeEach(async () => {
        pkg = new VisioPackage();
        const zip = new JSZip();
        zip.file('visio/pages/page1.xml', `
            <PageContents>
                <Shapes>
                    <Shape ID="1" Name="Process" Type="Shape">
                        <Text>Step 1</Text>
                        <Cell N="Width" V="2" />
                        <Cell N="Height" V="1" />
                        <Section N="Geometry">
                            <Row T="MoveTo" IX="1">
                                <Cell N="X" V="0"/>
                            </Row>
                        </Section>
                    </Shape>
                </Shapes>
            </PageContents>
        `);

        const buffer = await zip.generateAsync({ type: 'nodebuffer' });
        await pkg.load(buffer);
        reader = new ShapeReader(pkg);
    });

    it('should read shapes from a page file', () => {
        const shapes = reader.readShapes('visio/pages/page1.xml');
        expect(shapes).toHaveLength(1);

        const shape = shapes[0];
        expect(shape.Name).toBe('Process');
        expect(shape.Text).toBe('Step 1');
        expect(shape.Cells['Width'].V).toBe('2');
    });

    it('should read sections and rows', () => {
        const shapes = reader.readShapes('visio/pages/page1.xml');
        const shape = shapes[0];

        expect(shape.Sections['Geometry']).toBeDefined();
        expect(shape.Sections['Geometry'].Rows).toHaveLength(1);
        expect(shape.Sections['Geometry'].Rows[0].T).toBe('MoveTo');
    });

    it('should return empty array if file missing', () => {
        const shapes = reader.readShapes('visio/pages/nonexistent.xml');
        expect(shapes).toEqual([]);
    });
});
