import { VisioPackage } from './VisioPackage';
import { PageManager } from './core/PageManager';
import { Page } from './Page';
import { MediaManager } from './core/MediaManager';
import { MetadataManager } from './core/MetadataManager';
import { StyleSheetManager } from './core/StyleSheetManager';
import { ColorManager } from './core/ColorManager';
import { VisioPage, DocumentMetadata, StyleProps, StyleRecord, ColorEntry } from './types/VisioTypes';

export class VisioDocument {
    private pageManager: PageManager;
    private _pageCache: Page[] | null = null;
    private mediaManager: MediaManager;
    private metadataManager: MetadataManager;
    private styleSheetManager: StyleSheetManager;
    private colorManager: ColorManager;

    private constructor(private pkg: VisioPackage) {
        this.pageManager       = new PageManager(pkg);
        this.mediaManager      = new MediaManager(pkg);
        this.metadataManager   = new MetadataManager(pkg);
        this.styleSheetManager = new StyleSheetManager(pkg);
        this.colorManager      = new ColorManager(pkg);
    }

    static async create(): Promise<VisioDocument> {
        const pkg = await VisioPackage.create();
        return new VisioDocument(pkg);
    }

    static async load(pathOrBuffer: string | Buffer | ArrayBuffer | Uint8Array): Promise<VisioDocument> {
        const pkg = new VisioPackage();

        if (typeof pathOrBuffer === 'string') {
            const fs = await import('fs/promises');
            const buffer = await fs.readFile(pathOrBuffer);
            await pkg.load(buffer);
        } else {
            await pkg.load(pathOrBuffer);
        }

        return new VisioDocument(pkg);
    }

    async addPage(name: string): Promise<Page> {
        const newId = await this.pageManager.createPage(name);
        this._pageCache = null;

        const pageStub: VisioPage = {
            ID: newId,
            Name: name,
            xmlPath: `visio/pages/page${newId}.xml`,
            Shapes: [],
            Connects: []
        };
        return new Page(pageStub, this.pkg, this.mediaManager);
    }

    get pages(): Page[] {
        if (!this._pageCache) {
            const internalPages = this.pageManager.load();
            this._pageCache = internalPages.map(p => {
                const pageStub: VisioPage = {
                    ID: p.id.toString(),
                    Name: p.name,
                    // Thread the relationship-resolved path so that loaded files
                    // with non-sequential page filenames are handled correctly.
                    xmlPath: p.xmlPath,
                    Shapes: [],
                    Connects: [],
                    isBackground: p.isBackground,
                    backPageId: p.backPageId
                };
                return new Page(pageStub, this.pkg, this.mediaManager);
            });
        }

        return this._pageCache;
    }

    /**
     * Add a background page to the document
     */
    async addBackgroundPage(name: string): Promise<Page> {
        const newId = await this.pageManager.createBackgroundPage(name);
        this._pageCache = null;

        const pageStub: VisioPage = {
            ID: newId,
            Name: name,
            xmlPath: `visio/pages/page${newId}.xml`,
            Shapes: [],
            Connects: [],
            isBackground: true
        };
        return new Page(pageStub, this.pkg, this.mediaManager);
    }

    /**
     * Set a background page for a foreground page
     */
    async setBackgroundPage(foregroundPage: Page, backgroundPage: Page): Promise<void> {
        await this.pageManager.setBackgroundPage(foregroundPage.id, backgroundPage.id);
        this._pageCache = null;
    }

    /**
     * Find a page by name. Returns undefined if no page with that name exists.
     */
    getPage(name: string): Page | undefined {
        return this.pages.find(p => p.name === name);
    }

    /**
     * Delete a page from the document.
     * Removes the page XML, its relationships, the Content Types entry,
     * and any BackPage references from other pages.
     */
    async deletePage(page: Page): Promise<void> {
        await this.pageManager.deletePage(page.id);
        this._pageCache = null;
    }

    /**
     * Read document metadata from `docProps/core.xml` and `docProps/app.xml`.
     * Fields not present in the file are returned as `undefined`.
     */
    getMetadata(): DocumentMetadata {
        return this.metadataManager.getMetadata();
    }

    /**
     * Write document metadata. Only the supplied fields are changed;
     * all other fields keep their existing values.
     *
     * @example
     * doc.setMetadata({ title: 'My Diagram', author: 'Alice', company: 'ACME' });
     */
    setMetadata(props: Partial<DocumentMetadata>): void {
        this.metadataManager.setMetadata(props);
    }

    /**
     * Create a named document-level stylesheet and return its record.
     * The returned `id` can be passed to `addShape({ styleId })` or `shape.applyStyle()`.
     *
     * @example
     * const s = doc.createStyle('Header', { fillColor: '#4472C4', fontColor: '#ffffff', bold: true });
     * const shape = await page.addShape({ text: 'Title', x: 1, y: 1, width: 3, height: 1, styleId: s.id });
     */
    createStyle(name: string, props: StyleProps = {}): StyleRecord {
        return this.styleSheetManager.createStyle(name, props);
    }

    /**
     * Return all stylesheets defined in the document (including built-in styles).
     */
    getStyles(): StyleRecord[] {
        return this.styleSheetManager.getStyles();
    }

    /**
     * Add a color to the document-level color palette and return its integer index (IX).
     *
     * If the color is already registered the existing index is returned without
     * creating a duplicate. The two built-in colors are always present:
     *   - index 0  →  `#000000`  (black)
     *   - index 1  →  `#FFFFFF`  (white)
     *
     * User colors receive indices starting at 2.
     *
     * @param hex  CSS hex string — `'#4472C4'`, `'#ABC'`, or `'4472c4'` are all accepted.
     * @returns    Integer IX that uniquely identifies this color in the palette.
     *
     * @example
     * const blueIx = doc.addColor('#4472C4');  // → 2
     * const redIx  = doc.addColor('#FF0000');  // → 3
     * doc.addColor('#4472C4');                 // → 2 (de-duplicated)
     */
    addColor(hex: string): number {
        return this.colorManager.addColor(hex);
    }

    /**
     * Return all color entries in the document palette, ordered by index.
     *
     * @example
     * doc.addColor('#4472C4');
     * doc.getColors();
     * // → [{ index: 0, rgb: '#000000' }, { index: 1, rgb: '#FFFFFF' }, { index: 2, rgb: '#4472C4' }]
     */
    getColors(): ColorEntry[] {
        return this.colorManager.getColors();
    }

    /**
     * Look up the palette index of a color by its hex value.
     * Returns `undefined` if the color has not been added to the palette.
     *
     * @example
     * doc.addColor('#4472C4');
     * doc.getColorIndex('#4472C4');  // → 2
     * doc.getColorIndex('#FF0000');  // → undefined
     */
    getColorIndex(hex: string): number | undefined {
        return this.colorManager.getColorIndex(hex);
    }

    async save(filename?: string): Promise<Buffer> {
        return this.pkg.save(filename);
    }
}
