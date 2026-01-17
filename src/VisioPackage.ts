import JSZip from 'jszip';

export class VisioPackage {
    private zip: JSZip | null = null;
    private files: Map<string, string> = new Map();

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

    async save(): Promise<Buffer> {
        if (!this.zip) {
            throw new Error("Package not loaded");
        }
        return await this.zip.generateAsync({ type: 'nodebuffer' });
    }

    getFileText(path: string): string {
        const content = this.files.get(path);
        if (content === undefined) {
            throw new Error(`File not found: ${path}`);
        }
        return content;
    }
}
