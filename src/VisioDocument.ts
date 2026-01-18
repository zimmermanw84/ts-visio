import { VisioPackage } from './VisioPackage';
import { PageManager } from './PageManager';
import { Page } from './Page';

export class VisioDocument {
    private constructor(private pkg: VisioPackage) { }

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

    get pages(): Page[] {
        const pm = new PageManager(this.pkg);
        const internalPages = pm.getPages();
        return internalPages.map(p => new Page(p, this.pkg));
    }

    async save(filename?: string): Promise<Buffer> {
        return this.pkg.save(filename);
    }
}
