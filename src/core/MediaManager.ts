import { VisioPackage } from '../VisioPackage';
import { XMLParser, XMLBuilder } from 'fast-xml-parser';
import { MIME_TYPES } from './MediaConstants';

export class MediaManager {
    private parser: XMLParser;
    private builder: XMLBuilder;

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
    }



    private getContentType(extension: string): string {
        return MIME_TYPES[extension.toLowerCase()] || 'application/octet-stream';
    }

    addMedia(filename: string, data: Buffer): string {
        const path = `visio/media/${filename}`;
        this.pkg.updateFile(path, data);

        // Ensure content type
        if (filename.includes('.')) {
            const ext = filename.split('.').pop()?.toLowerCase();
            if (ext) {
                this.ensureContentType(ext);
            }
        }

        // Return path relative to page relationships
        // Typically pages are in visio/pages/ and link to ../media/...
        return `../media/${filename}`;
    }

    private ensureContentType(extension: string) {
        const ctPath = '[Content_Types].xml';
        const content = this.pkg.getFileText(ctPath);
        const parsed = this.parser.parse(content);

        // Ensure structure
        if (!parsed.Types) parsed.Types = {};
        if (!parsed.Types.Default) parsed.Types.Default = [];

        // Ensure array
        if (!Array.isArray(parsed.Types.Default)) {
            parsed.Types.Default = [parsed.Types.Default];
        }

        const defaults = parsed.Types.Default;
        const exists = defaults.some((d: any) => d['@_Extension']?.toLowerCase() === extension);

        if (!exists) {
            defaults.push({
                '@_Extension': extension,
                '@_ContentType': this.getContentType(extension)
            });
            // Array reference is already mutating object

            const newXml = this.builder.build(parsed);
            this.pkg.updateFile(ctPath, newXml);
        }
    }
}
