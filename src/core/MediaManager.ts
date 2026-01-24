import { VisioPackage } from '../VisioPackage';
import { XMLParser, XMLBuilder } from 'fast-xml-parser';
import { MIME_TYPES } from './MediaConstants';

export class MediaManager {
    private parser: XMLParser;
    private builder: XMLBuilder;

    private deduplicationMap = new Map<string, string>(); // hash -> paths
    private indexed = false;

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

    private ensureIndex() {
        if (this.indexed) return;

        const crypto = require('crypto');
        for (const [path, content] of this.pkg.filesMap.entries()) {
            if (path.startsWith('visio/media/') && Buffer.isBuffer(content)) {
                const hash = crypto.createHash('sha1').update(content).digest('hex');
                // Store relative path suitable for relationships
                const filename = path.split('/').pop();
                if (filename) {
                    this.deduplicationMap.set(hash, `../media/${filename}`);
                }
            }
        }
        this.indexed = true;
    }

    private getContentType(extension: string): string {
        return MIME_TYPES[extension.toLowerCase()] || 'application/octet-stream';
    }



    addMedia(name: string, data: Buffer): string {
        this.ensureIndex();
        const crypto = require('crypto');
        const hash = crypto.createHash('sha1').update(data).digest('hex');

        if (this.deduplicationMap.has(hash)) {
            return this.deduplicationMap.get(hash)!;
        }

        const extension = name.split('.').pop() || '';
        // If name already exists, we might need a unique name.
        // Note: Visio doesn't strictly require unique filenames in the media folder if we manage rels,
        // but it's safer to avoid collisions.

        // Use hash in filename to ensure uniqueness and automatic deduplication naming
        const uniqueFileName = `${name.replace(`.${extension}`, '')}_${hash.substring(0, 8)}.${extension}`;
        const mediaPath = `visio/media/${uniqueFileName}`;
        const relPath = `../media/${uniqueFileName}`;

        this.pkg.updateFile(mediaPath, data);

        if (extension) {
            this.ensureContentType(extension);
        }

        this.deduplicationMap.set(hash, relPath);
        return relPath;
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
