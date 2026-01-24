import { XMLParser, XMLBuilder } from 'fast-xml-parser';
import { VisioPackage } from '../VisioPackage';
import { XML_NAMESPACES, RELATIONSHIP_TYPES } from './VisioConstants';

export class RelsManager {
    private parser: XMLParser;
    private builder: XMLBuilder;
    private relsCache: Map<string, any> = new Map();

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

    private getRelsPath(partPath: string): string {
        // file.xml -> _rels/file.xml.rels
        // dir/file.xml -> dir/_rels/file.xml.rels
        const parts = partPath.split('/');
        const fileName = parts.pop();
        const dir = parts.join('/');
        return `${dir}/_rels/${fileName}.rels`;
    }

    async ensureRelationship(sourcePath: string, target: string, type: string): Promise<string> {
        const relsPath = this.getRelsPath(sourcePath);
        let parsed;

        if (this.relsCache.has(relsPath)) {
            parsed = this.relsCache.get(relsPath);
        } else {
            let content = '';
            try {
                content = this.pkg.getFileText(relsPath);
            } catch {
                // If .rels doesn't exist, start fresh
                content = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="${XML_NAMESPACES.RELATIONSHIPS}">
</Relationships>`;
            }
            parsed = this.parser.parse(content);
            this.relsCache.set(relsPath, parsed);
        }

        if (!parsed.Relationships) {
            parsed.Relationships = { Relationship: [] };
        }

        let rels = parsed.Relationships.Relationship;
        if (!Array.isArray(rels)) {
            rels = rels ? [rels] : [];
            parsed.Relationships.Relationship = rels;
        }

        // Check for existing
        const existing = rels.find((r: any) => r['@_Target'] === target && r['@_Type'] === type);
        if (existing) {
            return existing['@_Id'];
        }

        // Generate new ID (rId1, rId2...)
        let maxId = 0;
        for (const r of rels) {
            const idStr = r['@_Id']; // "rId5"
            const idNum = parseInt(idStr.replace('rId', ''));
            if (!isNaN(idNum) && idNum > maxId) maxId = idNum;
        }
        const newId = `rId${maxId + 1}`;

        // Add new relationship
        rels.push({
            '@_Id': newId,
            '@_Type': type,
            '@_Target': target
        });

        // Save back
        const newXml = this.builder.build(parsed);
        this.pkg.updateFile(relsPath, newXml);

        return newId;
    }

    async addPageImageRel(pageId: string, mediaPath: string): Promise<string> {
        const pagePath = `visio/pages/page${pageId}.xml`;
        return this.addImageRelationship(pagePath, mediaPath);
    }

    async addImageRelationship(sourcePath: string, target: string): Promise<string> {
        return this.ensureRelationship(sourcePath, target, RELATIONSHIP_TYPES.IMAGE);
    }
}
