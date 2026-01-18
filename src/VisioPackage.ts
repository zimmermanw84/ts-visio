import JSZip from 'jszip';
import * as T from './templates/MinimalVsdx';

export class VisioPackage {
    private zip: JSZip | null = null;
    private files: Map<string, string> = new Map();

    static async create(): Promise<VisioPackage> {
        const pkg = new VisioPackage();
        pkg.zip = new JSZip();

        // Populate minimal structure
        pkg.updateFile('[Content_Types].xml', T.CONTENT_TYPES_XML);
        pkg.updateFile('_rels/.rels', T.RELS_XML);
        pkg.updateFile('visio/document.xml', T.DOCUMENT_XML);
        pkg.updateFile('visio/_rels/document.xml.rels', T.DOCUMENT_RELS_XML);
        pkg.updateFile('visio/pages/pages.xml', T.PAGES_XML);
        pkg.updateFile('visio/pages/_rels/pages.xml.rels', T.PAGES_RELS_XML);
        pkg.updateFile('visio/pages/page1.xml', T.PAGE1_XML);
        pkg.updateFile('visio/windows.xml', T.WINDOWS_XML);
        pkg.updateFile('docProps/core.xml', T.CORE_XML);
        pkg.updateFile('docProps/app.xml', T.APP_XML);

        return pkg;
    }

    async load(buffer: Buffer | ArrayBuffer | Uint8Array): Promise<void> {
        this.files.clear();
        this.zip = await JSZip.loadAsync(buffer);

        const promises: Promise<void>[] = [];
        this.zip.forEach((relativePath, file) => {
            if (!file.dir) {
                promises.push(
                    file.async('string').then(content => {
                        this.files.set(relativePath, content);
                    })
                );
            }
        });

        await Promise.all(promises);
    }

    updateFile(path: string, content: string): void {
        if (!this.zip) {
            throw new Error("Package not loaded");
        }
        this.files.set(path, content);
        this.zip.file(path, content);
    }

    async save(filename?: string): Promise<Buffer> {
        if (!this.zip) {
            throw new Error("Package not loaded");
        }
        const buffer = await this.zip.generateAsync({ type: 'nodebuffer' });

        if (filename) {
            const fs = await import('fs/promises');
            await fs.writeFile(filename, buffer);
        }

        return buffer;
    }

    getFileText(path: string): string {
        const content = this.files.get(path);
        if (content === undefined) {
            throw new Error(`File not found: ${path}`);
        }
        return content;
    }
}
