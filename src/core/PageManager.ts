import { XMLParser, XMLBuilder } from 'fast-xml-parser';
import { XML_NAMESPACES, RELATIONSHIP_TYPES, CONTENT_TYPES } from './VisioConstants';
import { VisioPackage } from '../VisioPackage';
import { RelsManager } from './RelsManager';

export interface PageEntry {
    id: number;
    name: string;
    relId: string;
    xmlPath: string;
}

export class PageManager {
    private parser: XMLParser;
    private builder: XMLBuilder;
    private relsManager: RelsManager;
    private pages: PageEntry[] = [];
    private loaded: boolean = false;

    constructor(private pkg: VisioPackage) {
        this.parser = new XMLParser({
            ignoreAttributes: false,
            attributeNamePrefix: "@_"
        });
        this.builder = new XMLBuilder({
            ignoreAttributes: false,
            attributeNamePrefix: "@_",
            format: true
        });
        this.relsManager = new RelsManager(pkg);
    }

    load(force: boolean = false): PageEntry[] {
        if (!force && this.loaded) {
            return this.pages;
        }

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

        if (pageNodes.length === 0) {
            this.pages = [];
            this.loaded = true;
            return [];
        }

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
            // r:id is in the Rel child element, not as a Page attribute
            const rId = node.Rel?.['@_r:id'] || node['@_r:id']; // Support both for compatibility
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

        this.loaded = true;
        return this.pages;
    }

    async createPage(name: string): Promise<string> {
        this.load(); // Refresh state

        // 1. Calculate ID
        let maxId = 0;
        for (const p of this.pages) {
            if (p.id > maxId) maxId = p.id;
        }
        const newId = maxId + 1;
        const fileName = `page${newId}.xml`;
        const relativePath = `visio/pages/${fileName}`;

        // 2. Create Page File
        const pageContent = `<PageContents xmlns="${XML_NAMESPACES.VISIO_MAIN}" xmlns:r="${XML_NAMESPACES.RELATIONSHIPS_OFFICE}" xml:space="preserve">
    <PageSheet LineStyle="0" FillStyle="0" TextStyle="0">
        <Cell N="PageWidth" V="8.5"/>
        <Cell N="PageHeight" V="11"/>
        <Cell N="PageScale" V="1" Unit="MSG"/>
        <Cell N="DrawingScale" V="1" Unit="MSG"/>
        <Cell N="DrawingSizeType" V="0"/>
        <Cell N="DrawingScaleType" V="0"/>
        <Cell N="Inhibited" V="0"/>
        <Cell N="UIVisibility" V="0"/>
        <Cell N="PageDrawSizeType" V="0"/>
    </PageSheet>
    <Shapes/>
    <Connects/>
</PageContents>`;
        this.pkg.updateFile(relativePath, pageContent);

        // 3. Update Content Types
        const ctPath = '[Content_Types].xml';
        const ctContent = this.pkg.getFileText(ctPath);
        const parsedCt = this.parser.parse(ctContent);

        // Format: <Override PartName="/visio/pages/page2.xml" ContentType="application/vnd.ms-visio.page+xml"/>
        // Ensure Types.Override array exists
        if (!parsedCt.Types.Override) parsedCt.Types.Override = [];
        if (!Array.isArray(parsedCt.Types.Override)) parsedCt.Types.Override = [parsedCt.Types.Override];

        parsedCt.Types.Override.push({
            '@_PartName': `/${relativePath}`,
            '@_ContentType': CONTENT_TYPES.VISIO_PAGE
        });
        this.pkg.updateFile(ctPath, this.builder.build(parsedCt));

        // 4. Update Relationships (pages.xml -> new page file)
        // Source is "visio/pages/pages.xml", Target is "page{ID}.xml" (relative to source dir)
        const rId = await this.relsManager.ensureRelationship(
            'visio/pages/pages.xml',
            fileName,
            RELATIONSHIP_TYPES.PAGE
        );

        // 5. Update Pages Index (visio/pages/pages.xml)
        const pagesPath = 'visio/pages/pages.xml';
        const pagesContent = this.pkg.getFileText(pagesPath);
        const parsedPages = this.parser.parse(pagesContent);

        if (!parsedPages.Pages.Page) parsedPages.Pages.Page = [];
        if (!Array.isArray(parsedPages.Pages.Page)) parsedPages.Pages.Page = [parsedPages.Pages.Page];

        parsedPages.Pages.Page.push({
            '@_ID': newId.toString(),
            '@_Name': name,
            'Rel': { '@_r:id': rId }
        });

        this.pkg.updateFile(pagesPath, this.builder.build(parsedPages));

        // Reload to include new page
        this.load(true);

        return newId.toString();
    }
}
