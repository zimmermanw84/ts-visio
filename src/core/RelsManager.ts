import { XMLParser, XMLBuilder } from 'fast-xml-parser';
import { VisioPackage } from '../VisioPackage';
import { XML_NAMESPACES, RELATIONSHIP_TYPES } from './VisioConstants';
import { createXmlParser, createXmlBuilder, buildXml } from '../utils/XmlHelper';

export class RelsManager {
    private parser: XMLParser;
    private builder: XMLBuilder;

    constructor(private pkg: VisioPackage) {
        this.parser = createXmlParser();
        this.builder = createXmlBuilder();
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
        let content = '';

        try {
            content = this.pkg.getFileText(relsPath);
        } catch {
            // If .rels doesn't exist, start fresh
            content = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="${XML_NAMESPACES.RELATIONSHIPS}">
</Relationships>`;
        }

        const parsed = this.parser.parse(content);

        if (!parsed.Relationships) {
            parsed.Relationships = { Relationship: [] };
        }

        let rels = parsed.Relationships.Relationship;
        if (!Array.isArray(rels)) {
            rels = rels ? [rels] : [];
            parsed.Relationships.Relationship = rels;
        }

        const existing = rels.find((r: any) => r['@_Target'] === target && r['@_Type'] === type);
        if (existing) {
            return existing['@_Id'];
        }

        let maxId = 0;
        for (const r of rels) {
            const idNum = parseInt(r['@_Id'].replace('rId', ''));
            if (!isNaN(idNum) && idNum > maxId) maxId = idNum;
        }
        const newId = `rId${maxId + 1}`;

        rels.push({
            '@_Id': newId,
            '@_Type': type,
            '@_Target': target
        });

        const newXml = buildXml(this.builder, parsed);
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
