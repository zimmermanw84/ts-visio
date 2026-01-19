import { XMLParser } from 'fast-xml-parser';
import { VisioPackage } from '../VisioPackage';

export interface Master {
    id: string;
    name: string;
    nameU: string;
    type: string;
    xmlPath: string; // Typically "masters/masterX.xml"
}

export class MasterManager {
    private parser: XMLParser;
    private masters: Master[] = [];

    constructor(private pkg: VisioPackage) {
        this.parser = new XMLParser({
            ignoreAttributes: false,
            attributeNamePrefix: "@_"
        });
    }

    load(): Master[] {
        const path = 'visio/masters/masters.xml';
        let content: string;
        try {
            content = this.pkg.getFileText(path);
        } catch (e) {
            // It's possible the file doesn't exist if no masters are used
            return [];
        }

        const parsed = this.parser.parse(content);

        let masterNodes = parsed.Masters ? parsed.Masters.Master : [];
        if (!Array.isArray(masterNodes)) {
            masterNodes = [masterNodes];
        }

        this.masters = masterNodes.map((node: any) => ({
            id: node['@_ID'],
            name: node['@_Name'],
            nameU: node['@_NameU'],
            type: node['@_Type'],
            // Implicit path convention in Visio typically follows patterns,
            // but precise mapping usually requires checking relationships (.rels).
            // For Phase 1, we'll assume standard naming or just store what we have.
            // Relationships are handled in Phase 3.
            xmlPath: ''
        }));

        return this.masters;
    }
}
