import { XMLParser } from 'fast-xml-parser';
import { VisioPackage } from './VisioPackage';
import { VisioPage } from './types/VisioTypes';

export class PageManager {
    private parser: XMLParser;

    constructor(private pkg: VisioPackage) {
        this.parser = new XMLParser({
            ignoreAttributes: false,
            attributeNamePrefix: "@_"
        });
    }

    getPages(): VisioPage[] {
        let content: string;
        try {
            content = this.pkg.getFileText('visio/pages/pages.xml');
        } catch {
            return [];
        }

        const parsed = this.parser.parse(content);
        const pagesData = parsed.Pages?.Page;

        if (!pagesData) return [];

        const pagesArray = Array.isArray(pagesData) ? pagesData : [pagesData];

        return pagesArray.map((p: any) => ({
            ID: p['@_ID'],
            Name: p['@_Name'],
            NameU: p['@_NameU'],
            Shapes: [],
            Connects: []
        }));
    }
}
