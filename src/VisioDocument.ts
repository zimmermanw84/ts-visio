import { VisioPackage } from './VisioPackage';
import { PageManager } from './core/PageManager';
import { Page } from './Page';

export class VisioDocument {
    private pageManager: PageManager;
    private _pageCache: Page[] | null = null;

    private constructor(private pkg: VisioPackage) {
        this.pageManager = new PageManager(pkg);
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

        const pageStub = {
            ID: newId,
            Name: name,
            Shapes: [],
            Connects: []
        };
        return new Page(pageStub as any, this.pkg);
    }

    get pages(): Page[] {
        if (!this._pageCache) {
            const internalPages = this.pageManager.load();
            this._pageCache = internalPages.map(p => {
                // Adapter for VisioPage interface
                const pageStub = {
                    ID: p.id.toString(),
                    Name: p.name,
                    Shapes: [],
                    Connects: []
                };
                return new Page(pageStub as any, this.pkg);
            });
        }
        return this._pageCache;
    }

    async save(filename?: string): Promise<Buffer> {
        return this.pkg.save(filename);
    }
}
