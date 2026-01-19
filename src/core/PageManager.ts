import { XMLParser } from 'fast-xml-parser';
import { VisioPackage } from '../VisioPackage';

export interface PageEntry {
    id: number;
    name: string;
    relId: string;
    xmlPath: string;
}

export class PageManager {
    private parser: XMLParser;
    private pages: PageEntry[] = [];

    constructor(private pkg: VisioPackage) {
        this.parser = new XMLParser({
            ignoreAttributes: false,
            attributeNamePrefix: "@_"
        });
    }

    load(): PageEntry[] {
        // 1. Load Pages Index
        let pagesContent: string;
        try {
            pagesContent = this.pkg.getFileText('visio/pages/pages.xml');
        } catch {
            return [];
        }

        const parsedPages = this.parser.parse(pagesContent);
        let pageNodes = parsedPages.Pages ? parsedPages.Pages.Page : [];
        if (!Array.isArray(pageNodes)) {
            pageNodes = pageNodes ? [pageNodes] : [];
        }

        if (pageNodes.length === 0) return [];

        // 2. Load Relationships to resolve paths
        let relsContent: string;
        try {
            relsContent = this.pkg.getFileText('visio/pages/_rels/pages.xml.rels');
        } catch {
            relsContent = '';
        }

        const relsMap = new Map<string, string>(); // rId -> Target
        if (relsContent) {
            const parsedRels = this.parser.parse(relsContent);
            let relNodes = parsedRels.Relationships ? parsedRels.Relationships.Relationship : [];
            if (!Array.isArray(relNodes)) relNodes = relNodes ? [relNodes] : [];

            for (const r of relNodes) {
                relsMap.set(r['@_Id'], r['@_Target']);
            }
        }

        // 3. Map Pages
        this.pages = pageNodes.map((node: any) => {
            const rId = node['@_r:id'];
            const target = relsMap.get(rId) || '';
            // Target is usually "page1.xml" or "pages/page1.xml" depending on relative structure.
            // pages.xml is in "visio/pages/", so "page1.xml" means "visio/pages/page1.xml"
            const fullPath = target ? `visio/pages/${target}` : '';

            return {
                id: parseInt(node['@_ID']),
                name: node['@_Name'],
                relId: rId,
                xmlPath: fullPath
            };
        });

        return this.pages;
    }
}
