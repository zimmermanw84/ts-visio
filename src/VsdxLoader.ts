import JSZip from 'jszip';
import { XMLParser } from 'fast-xml-parser';
import { VisioPage, VisioShape, VisioCell, VisioSection, VisioRow, VisioConnect } from './types/VisioTypes';
import { asArray, parseCells, parseSection } from './utils/VisioParsers';

export class VsdxLoader {
    private zip: JSZip;
    private parser: XMLParser;

    constructor() {
        this.zip = new JSZip();
        this.parser = new XMLParser({
            ignoreAttributes: false,
            attributeNamePrefix: "@_"
        });
    }

    async load(data: Buffer | ArrayBuffer | Uint8Array): Promise<void> {
        this.zip = await JSZip.loadAsync(data);
    }

    getFileNames(): string[] {
        return Object.keys(this.zip.files);
    }

    async getFileContent(path: string): Promise<string | null> {
        const file = this.zip.file(path);
        if (!file) return null;
        return await file.async('string');
    }

    async getAppXml(): Promise<any> {
        const content = await this.getFileContent('docProps/app.xml');
        if (!content) return null;
        return this.parser.parse(content);
    }

    async getPages(): Promise<VisioPage[]> {
        const content = await this.getFileContent('visio/pages/pages.xml');
        if (!content) return [];

        const parsed = this.parser.parse(content);
        const pages = asArray(parsed.Pages?.Page);

        return pages.map((p: any) => ({
            ID: p['@_ID'],
            Name: p['@_Name'],
            NameU: p['@_NameU'],
            Shapes: [],    // Populated later if requested
            Connects: []  // Connects typically live in the Page XML file, not pages.xml
        }));
    }

    // Helper to get connects from a page file
    async getPageConnects(path: string): Promise<VisioConnect[]> {
        const content = await this.getFileContent(path);
        if (!content) return [];

        const parsed = this.parser.parse(content);
        const connects = asArray(parsed.PageContents?.Connects?.Connect);

        return connects.map((c: any) => ({
            FromSheet: c['@_FromSheet'],
            FromCell: c['@_FromCell'],
            FromPart: c['@_FromPart'],
            ToSheet: c['@_ToSheet'],
            ToCell: c['@_ToCell'],
            ToPart: c['@_ToPart']
        }));
    }

    async getPageShapes(path: string): Promise<VisioShape[]> {
        const content = await this.getFileContent(path);
        if (!content) return [];

        const parsed = this.parser.parse(content);
        // Note: PageContents can contain Shapes directly, or deeper structures.
        // For now, we assume simple flat shapes list.
        const shapes = asArray(parsed.PageContents?.Shapes?.Shape);

        return shapes.map((s: any) => this.parseShape(s));
    }

    private parseShape(s: any): VisioShape {
        const shape: VisioShape = {
            ID: s['@_ID'],
            Name: s['@_Name'],
            NameU: s['@_NameU'],
            Type: s['@_Type'],
            Master: s['@_Master'],
            Text: s.Text?.['#text'] || s.Text || undefined, // Simple handling for now
            Cells: parseCells(s),
            Sections: {}
        };

        const sections = asArray(s.Section);
        for (const sec of sections) {
            const section = sec as any;
            if (section['@_N']) {
                shape.Sections[section['@_N']] = parseSection(section);
            }
        }

        return shape;
    }

    async setFileContent(path: string, content: string): Promise<void> {
        this.zip.file(path, content);
    }

    async save(): Promise<Buffer> {
        return await this.zip.generateAsync({ type: 'nodebuffer' });
    }
}
