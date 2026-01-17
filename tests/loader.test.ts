import { describe, it, expect, beforeEach } from 'vitest';
import JSZip from 'jszip';
import { VsdxLoader } from '../src/VsdxLoader';

describe('VsdxLoader', () => {
    let loader: VsdxLoader;
    let mockVsdxBuffer: Buffer;

    beforeEach(async () => {
        loader = new VsdxLoader();

        // Create a mock VSDX file (which is just a zip)
        const zip = new JSZip();
        zip.file('[Content_Types].xml', '<Types></Types>');
        zip.file('docProps/app.xml', '<Properties><Application>Microsoft Visio</Application></Properties>');

        mockVsdxBuffer = await zip.generateAsync({ type: 'nodebuffer' });
    });

    it('should load a vsdx file buffer', async () => {
        await loader.load(mockVsdxBuffer);
        const files = loader.getFileNames();
        expect(files).toContain('[Content_Types].xml');
        expect(files).toContain('docProps/app.xml');
    });

    it('should parse docProps/app.xml', async () => {
        await loader.load(mockVsdxBuffer);
        const appXml = await loader.getAppXml();
        expect(appXml).toBeDefined();
        expect(appXml.Properties.Application).toBe('Microsoft Visio');
    });

    it('should list pages from visio/pages/pages.xml', async () => {
        const zip = new JSZip();
        zip.file('visio/pages/pages.xml', `
            <Pages>
                <Page ID="0" Name="Page-1" />
                <Page ID="1" Name="Page-2" />
            </Pages>
        `);
        const buffer = await zip.generateAsync({ type: 'nodebuffer' });

        await loader.load(buffer);
        const pages = await loader.getPages();

        expect(pages).toHaveLength(2);
        expect(pages[0].Name).toBe('Page-1');
        expect(pages[1].Name).toBe('Page-2');
    });

    it('should list shapes from a page file', async () => {
        const zip = new JSZip();
        zip.file('visio/pages/page1.xml', `
            <PageContents>
                <Shapes>
                    <Shape ID="1" Name="Process" Type="Shape">
                        <Text>Step 1</Text>
                        <Cell N="Width" V="2" />
                        <Cell N="Height" V="1" />
                    </Shape>
                    <Shape ID="2" Name="Decision" Type="Shape">
                        <Text>Is it compatible?</Text>
                    </Shape>
                </Shapes>
            </PageContents>
        `);
        const buffer = await zip.generateAsync({ type: 'nodebuffer' });

        await loader.load(buffer);
        const shapes = await loader.getPageShapes('visio/pages/page1.xml');

        expect(shapes).toHaveLength(2);
        expect(shapes[0].Name).toBe('Process');
        expect(shapes[0].Text).toBe('Step 1');

        // Test Cells
        expect(shapes[0].Cells['Width']).toBeDefined();
        expect(shapes[0].Cells['Width'].V).toBe('2');
    });

    it('should update file content and save', async () => {
        const zip = new JSZip();
        zip.file('test.xml', '<Data>Old</Data>');
        const buffer = await zip.generateAsync({ type: 'nodebuffer' });

        await loader.load(buffer);
        await loader.setFileContent('test.xml', '<Data>New</Data>');
        const newBuffer = await loader.save();

        const newLoader = new VsdxLoader();
        await newLoader.load(newBuffer);
        const content = await newLoader.getFileContent('test.xml');
        expect(content).toBe('<Data>New</Data>');
    });

    it('should list connections from a page file', async () => {
        const zip = new JSZip();
        zip.file('visio/pages/page1.xml', `
            <PageContents>
                <Shapes>
                    <Shape ID="1" Name="Process" />
                    <Shape ID="2" Name="Decision" />
                </Shapes>
                <Connects>
                    <Connect FromSheet="3" FromCell="BeginX" ToSheet="1" ToCell="PinX" />
                    <Connect FromSheet="3" FromCell="EndX" ToSheet="2" ToCell="PinX" />
                </Connects>
            </PageContents>
        `);
        const buffer = await zip.generateAsync({ type: 'nodebuffer' });

        await loader.load(buffer);
        const connects = await loader.getPageConnects('visio/pages/page1.xml');

        expect(connects).toHaveLength(2);
        expect(connects[0].FromSheet).toBe('3');
        expect(connects[0].ToSheet).toBe('1');
        expect(connects[1].ToSheet).toBe('2');
    });
});
