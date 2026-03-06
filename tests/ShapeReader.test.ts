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

    it('should keep numeric-looking attribute values as strings (parseAttributeValue: false)', async () => {
        // Regression test for Bug 25: ShapeReader must use createXmlParser() so that
        // parseAttributeValue: false is in effect — attribute values like IDs and V=
        // must never be coerced to numbers.
        const zip = new JSZip();
        zip.file('visio/pages/page2.xml', `
            <PageContents>
                <Shapes>
                    <Shape ID="42" Name="Box" Type="Shape">
                        <Cell N="Width" V="3" />
                    </Shape>
                </Shapes>
            </PageContents>
        `);
        const buffer = await zip.generateAsync({ type: 'nodebuffer' });
        const pkg2 = new VisioPackage();
        await pkg2.load(buffer);
        const reader2 = new ShapeReader(pkg2);

        const shapes = reader2.readShapes('visio/pages/page2.xml');
        expect(shapes).toHaveLength(1);
        // ID and cell values must be strings, not numbers
        expect(typeof shapes[0].ID).toBe('string');
        expect(shapes[0].ID).toBe('42');
        expect(typeof shapes[0].Cells['Width'].V).toBe('string');
        expect(shapes[0].Cells['Width'].V).toBe('3');
    });
});
