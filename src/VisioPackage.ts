import JSZip from 'jszip';

export class VisioPackage {
    private files: Map<string, string> = new Map();

    async load(buffer: Buffer | ArrayBuffer | Uint8Array): Promise<void> {
        this.files.clear();
        const zip = await JSZip.loadAsync(buffer);

        const promises: Promise<void>[] = [];
        zip.forEach((relativePath, file) => {
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

    getFileText(path: string): string {
        const content = this.files.get(path);
        if (content === undefined) {
            throw new Error(`File not found: ${path}`);
        }
        return content;
    }
}
