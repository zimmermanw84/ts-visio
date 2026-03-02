import { VisioPackage } from './VisioPackage';
import { PageManager } from './core/PageManager';
import { MasterManager } from './core/MasterManager';
import { Page } from './Page';
import { MediaManager } from './core/MediaManager';
import { MetadataManager } from './core/MetadataManager';
import { StyleSheetManager } from './core/StyleSheetManager';
import { ColorManager } from './core/ColorManager';
import { VisioPage, DocumentMetadata, StyleProps, StyleRecord, ColorEntry, MasterRecord, ShapeGeometry } from './types/VisioTypes';

export class VisioDocument {
    private pageManager: PageManager;
    private masterManager: MasterManager;
    private _pageCache: Page[] | null = null;
    private mediaManager: MediaManager;
    private metadataManager: MetadataManager;
    private styleSheetManager: StyleSheetManager;
    private colorManager: ColorManager;

    private constructor(private pkg: VisioPackage) {
        this.pageManager       = new PageManager(pkg);
        this.masterManager     = new MasterManager(pkg);
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
     * Rename a page. Updates `page.name` in-memory as well as the pages.xml record.
     *
     * @example
     * doc.renamePage(page, 'Architecture Overview');
     */
    renamePage(page: Page, newName: string): void {
        this.pageManager.renamePage(page.id, newName);
        page._updateName(newName);
    }

    /**
     * Move a page to a new 0-based position in the tab order.
     * Clamps `toIndex` to the valid range automatically.
     *
     * @example
     * doc.movePage(page, 0);  // move to first position
     */
    movePage(page: Page, toIndex: number): void {
        this.pageManager.reorderPage(page.id, toIndex);
        this._pageCache = null;
    }

    /**
     * Duplicate a page and return the new Page object.
     * The duplicate is inserted directly after the source page in the tab order.
     * If `newName` is omitted, `"<original name> (Copy)"` is used.
     *
     * @example
     * const copy = await doc.duplicatePage(page, 'Page 2');
     */
    async duplicatePage(page: Page, newName?: string): Promise<Page> {
        const resolvedName = newName ?? `${page.name} (Copy)`;
        const newId = await this.pageManager.duplicatePage(page.id, resolvedName);
        this._pageCache = null;

        const pageStub: VisioPage = {
            ID: newId,
            Name: resolvedName,
            xmlPath: `visio/pages/page${newId}.xml`,
            Shapes: [],
            Connects: []
        };
        return new Page(pageStub, this.pkg, this.mediaManager);
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

    // -------------------------------------------------------------------------
    // Masters
    // -------------------------------------------------------------------------

    /**
     * Return all master shapes currently defined in the document.
     * Each record's `id` can be passed as `masterId` to `page.addShape()`.
     *
     * @example
     * const masters = doc.getMasters();
     * const box = masters.find(m => m.name === 'Box')!;
     * await page.addShape({ text: '', x: 2, y: 3, width: 1, height: 1, masterId: box.id });
     */
    getMasters(): MasterRecord[] {
        return this.masterManager.load();
    }

    /**
     * Define a new master shape in the document and return its record.
     * Creates all necessary OPC infrastructure (`masters.xml`, content-type
     * overrides, and document-level relationships) on the first call.
     *
     * @param name     Display name shown in the stencil panel.
     * @param geometry Visual outline; defaults to `'rectangle'`.
     *
     * @example
     * const boxMaster = doc.createMaster('Box');
     * const ellMaster = doc.createMaster('Process', 'ellipse');
     * await page.addShape({ text: 'Start', x: 1, y: 1, width: 2, height: 1, masterId: ellMaster.id });
     */
    createMaster(name: string, geometry: ShapeGeometry = 'rectangle'): MasterRecord {
        return this.masterManager.createMaster(name, geometry);
    }

    /**
     * Import all masters from a `.vssx` stencil file into this document.
     * Each master is assigned a fresh ID that does not conflict with any
     * master already present. Returns the array of imported master records.
     *
     * @param pathOrBuffer  Filesystem path or raw buffer of the `.vssx` file.
     *
     * @example
     * const masters = await doc.importMastersFromStencil('./Basic_Shapes.vssx');
     * const arrowMaster = masters.find(m => m.name === 'Arrow');
     * await page.addShape({ text: '', x: 3, y: 4, width: 1, height: 0.5, masterId: arrowMaster!.id });
     */
    importMastersFromStencil(
        pathOrBuffer: string | Buffer | ArrayBuffer | Uint8Array
    ): Promise<MasterRecord[]> {
        return this.masterManager.importFromStencil(pathOrBuffer);
    }

    async save(filename?: string): Promise<Buffer> {
        return this.pkg.save(filename);
    }
}
