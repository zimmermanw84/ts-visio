import JSZip from 'jszip';
import { XMLParser, XMLBuilder } from 'fast-xml-parser';
import { VisioPackage } from '../VisioPackage';
import { MasterRecord, ShapeGeometry } from '../types/VisioTypes';
import { XML_NAMESPACES, RELATIONSHIP_TYPES, CONTENT_TYPES } from './VisioConstants';
import { GeometryBuilder } from '../shapes/GeometryBuilder';
import { createXmlParser, createXmlBuilder, buildXml } from '../utils/XmlHelper';

export class MasterManager {
    private parser: XMLParser;
    private builder: XMLBuilder;

    constructor(private pkg: VisioPackage) {
        this.parser = createXmlParser();
        this.builder = createXmlBuilder();
    }

    // -------------------------------------------------------------------------
    // Read
    // -------------------------------------------------------------------------

    /**
     * Return all master shapes defined in the document.
     * Returns an empty array when no masters are present.
     */
    load(): MasterRecord[] {
        const mastersPath = 'visio/masters/masters.xml';
        let content: string;
        try {
            content = this.pkg.getFileText(mastersPath);
        } catch {
            return [];
        }

        const parsed = this.parser.parse(content);
        let masterNodes = parsed.Masters?.Master ?? [];
        if (!Array.isArray(masterNodes)) {
            masterNodes = masterNodes ? [masterNodes] : [];
        }

        // Build rId → target map from masters.xml.rels so we can populate xmlPath
        const relsMap = this.loadMastersRels();

        return masterNodes.map((node: any) => {
            const rId: string = node.Rel?.['@_r:id'] ?? '';
            const target: string = relsMap.get(rId) ?? '';
            const xmlPath = target ? `visio/masters/${target}` : '';
            return {
                id:      String(node['@_ID']),
                name:    node['@_Name'],
                nameU:   node['@_NameU'] ?? node['@_Name'],
                xmlPath,
            };
        });
    }

    // -------------------------------------------------------------------------
    // Create
    // -------------------------------------------------------------------------

    /**
     * Define a new master shape in the document and return its record.
     * The returned `id` can be passed as `masterId` when calling `page.addShape()`.
     *
     * @param name     Display name (also used as NameU).
     * @param geometry Visual outline; defaults to `'rectangle'`.
     */
    createMaster(name: string, geometry: ShapeGeometry = 'rectangle'): MasterRecord {
        this.ensureInfrastructure();

        const existingIds = this.load().map(m => parseInt(m.id)).filter(n => !isNaN(n));
        const newId = existingIds.length > 0 ? Math.max(...existingIds) + 1 : 1;
        const fileName = `master${newId}.xml`;
        const filePath = `visio/masters/${fileName}`;

        // Build the master's canonical shape object (parametric, 1×1 unit)
        const masterShape = this.buildMasterShapeObject(geometry);

        // Write masterN.xml (MasterContents)
        const masterContentsObj = {
            MasterContents: {
                '@_xmlns':        XML_NAMESPACES.VISIO_MAIN,
                '@_xmlns:r':      XML_NAMESPACES.RELATIONSHIPS_OFFICE,
                '@_xml:space':    'preserve',
                Shapes: { Shape: masterShape },
            }
        };
        this.pkg.updateFile(filePath, buildXml(this.builder, masterContentsObj));

        // Register content-type and OPC relationships
        this.addContentTypeOverride(`/${filePath}`, CONTENT_TYPES.VISIO_MASTER);
        const rId = this.addMastersRel(fileName);

        // Append entry to masters.xml
        this.addMasterEntry(newId, name, name, masterShape, rId);

        return { id: String(newId), name, nameU: name, xmlPath: filePath };
    }

    // -------------------------------------------------------------------------
    // Import from stencil (.vssx)
    // -------------------------------------------------------------------------

    /**
     * Import all masters from a `.vssx` stencil file into this document.
     * Each imported master is assigned a new ID that does not conflict with
     * masters already in the document. Returns the array of imported masters.
     *
     * @param pathOrBuffer  Filesystem path or raw buffer of the `.vssx` file.
     */
    async importFromStencil(
        pathOrBuffer: string | Buffer | ArrayBuffer | Uint8Array
    ): Promise<MasterRecord[]> {
        // Load stencil ZIP
        let buf: Buffer | ArrayBuffer | Uint8Array;
        if (typeof pathOrBuffer === 'string') {
            const fs = await import('fs/promises');
            buf = await fs.readFile(pathOrBuffer);
        } else {
            buf = pathOrBuffer;
        }

        const stencilZip = await JSZip.loadAsync(buf);

        // Read master index from stencil
        const mastersXmlFile = stencilZip.file('visio/masters/masters.xml');
        if (!mastersXmlFile) return [];

        const mastersXmlContent = await mastersXmlFile.async('string');
        const parsedMasters = this.parser.parse(mastersXmlContent);
        let masterNodes: any[] = parsedMasters.Masters?.Master ?? [];
        if (!Array.isArray(masterNodes)) masterNodes = masterNodes ? [masterNodes] : [];
        if (masterNodes.length === 0) return [];

        // Build rId → path map from stencil's masters.xml.rels
        const stencilRels = new Map<string, string>();
        const stencilRelsFile = stencilZip.file('visio/masters/_rels/masters.xml.rels');
        if (stencilRelsFile) {
            const relsContent = await stencilRelsFile.async('string');
            const parsedRels = this.parser.parse(relsContent);
            let rels: any[] = parsedRels.Relationships?.Relationship ?? [];
            if (!Array.isArray(rels)) rels = rels ? [rels] : [];
            for (const r of rels) {
                stencilRels.set(r['@_Id'], r['@_Target']);
            }
        }

        this.ensureInfrastructure();

        const imported: MasterRecord[] = [];

        for (const node of masterNodes) {
            // Assign a fresh ID that doesn't collide with any existing master
            const taken = this.load().map(m => parseInt(m.id)).filter(n => !isNaN(n));
            const newId = taken.length > 0 ? Math.max(...taken) + 1 : 1;
            const newFileName = `master${newId}.xml`;
            const newFilePath = `visio/masters/${newFileName}`;

            // Retrieve the individual master content from the stencil
            const stencilRId: string = node.Rel?.['@_r:id'] ?? '';
            const stencilTarget = stencilRels.get(stencilRId) ?? '';
            const stencilMasterPath = stencilTarget ? `visio/masters/${stencilTarget}` : '';

            let masterContentsXml: string;
            if (stencilMasterPath) {
                const masterFile = stencilZip.file(stencilMasterPath);
                masterContentsXml = masterFile
                    ? await masterFile.async('string')
                    : this.buildMasterContentsXml(node.Shapes);
            } else {
                masterContentsXml = this.buildMasterContentsXml(node.Shapes);
            }

            this.pkg.updateFile(newFilePath, masterContentsXml);
            this.addContentTypeOverride(`/${newFilePath}`, CONTENT_TYPES.VISIO_MASTER);
            const rId = this.addMastersRel(newFileName);

            const masterName: string  = node['@_Name']  ?? `Master${newId}`;
            const masterNameU: string = node['@_NameU'] ?? masterName;
            this.addMasterEntryFromNode(newId, masterName, masterNameU, node.Shapes, rId);

            imported.push({ id: String(newId), name: masterName, nameU: masterNameU, xmlPath: newFilePath });
        }

        return imported;
    }

    // -------------------------------------------------------------------------
    // Private helpers
    // -------------------------------------------------------------------------

    /**
     * Ensure `visio/masters/masters.xml` and its OPC bookkeeping exist.
     * Creates them on the first call; subsequent calls are no-ops.
     */
    private ensureInfrastructure(): void {
        try {
            this.pkg.getFileText('visio/masters/masters.xml');
            return; // already exists
        } catch { /* fall through to create */ }

        // Create empty masters.xml
        const mastersObj = {
            Masters: {
                '@_xmlns':     XML_NAMESPACES.VISIO_MAIN,
                '@_xmlns:r':   XML_NAMESPACES.RELATIONSHIPS_OFFICE,
                '@_xml:space': 'preserve',
                Master: [],
            }
        };
        this.pkg.updateFile('visio/masters/masters.xml', buildXml(this.builder, mastersObj));

        // Content-type for masters.xml
        this.addContentTypeOverride('/visio/masters/masters.xml', CONTENT_TYPES.VISIO_MASTERS);

        // Relationship from document.xml → masters/masters.xml
        this.ensureDocumentMastersRel();
    }

    /**
     * Add (or verify) a relationship from `visio/document.xml` to `masters/masters.xml`
     * in `visio/_rels/document.xml.rels`.
     */
    private ensureDocumentMastersRel(): void {
        const docRelsPath = 'visio/_rels/document.xml.rels';
        let content: string;
        try {
            content = this.pkg.getFileText(docRelsPath);
        } catch {
            content = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="${XML_NAMESPACES.RELATIONSHIPS}"></Relationships>`;
        }

        const parsed = this.parser.parse(content);
        if (!parsed.Relationships) parsed.Relationships = {};
        let rels: any[] = parsed.Relationships.Relationship ?? [];
        if (!Array.isArray(rels)) rels = rels ? [rels] : [];

        const already = rels.find(
            (r: any) => r['@_Type'] === RELATIONSHIP_TYPES.MASTERS &&
                        r['@_Target'] === 'masters/masters.xml'
        );
        if (already) return;

        const maxId = rels.reduce((max: number, r: any) => {
            const n = parseInt(r['@_Id']?.replace('rId', '') ?? '0');
            return isNaN(n) ? max : Math.max(max, n);
        }, 0);

        rels.push({
            '@_Id':     `rId${maxId + 1}`,
            '@_Type':   RELATIONSHIP_TYPES.MASTERS,
            '@_Target': 'masters/masters.xml',
        });

        parsed.Relationships.Relationship = rels;
        this.pkg.updateFile(docRelsPath, buildXml(this.builder, parsed));
    }

    /**
     * Add an Override entry to `[Content_Types].xml` unless it already exists.
     */
    private addContentTypeOverride(partName: string, contentType: string): void {
        const ctPath = '[Content_Types].xml';
        const parsed = this.parser.parse(this.pkg.getFileText(ctPath));
        if (!parsed.Types.Override) parsed.Types.Override = [];
        if (!Array.isArray(parsed.Types.Override)) {
            parsed.Types.Override = [parsed.Types.Override];
        }

        const already = (parsed.Types.Override as any[]).find(
            (o: any) => o['@_PartName'] === partName
        );
        if (already) return;

        parsed.Types.Override.push({ '@_PartName': partName, '@_ContentType': contentType });
        this.pkg.updateFile(ctPath, buildXml(this.builder, parsed));
    }

    /**
     * Add a relationship from `masters.xml` to `fileName` in
     * `visio/masters/_rels/masters.xml.rels`. Returns the rId.
     */
    private addMastersRel(fileName: string): string {
        const relsPath = 'visio/masters/_rels/masters.xml.rels';
        let content: string;
        try {
            content = this.pkg.getFileText(relsPath);
        } catch {
            content = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="${XML_NAMESPACES.RELATIONSHIPS}"></Relationships>`;
        }

        const parsed = this.parser.parse(content);
        if (!parsed.Relationships) parsed.Relationships = {};
        let rels: any[] = parsed.Relationships.Relationship ?? [];
        if (!Array.isArray(rels)) rels = rels ? [rels] : [];

        const existing = rels.find((r: any) => r['@_Target'] === fileName);
        if (existing) return existing['@_Id'];

        const maxId = rels.reduce((max: number, r: any) => {
            const n = parseInt(r['@_Id']?.replace('rId', '') ?? '0');
            return isNaN(n) ? max : Math.max(max, n);
        }, 0);
        const rId = `rId${maxId + 1}`;

        rels.push({ '@_Id': rId, '@_Type': RELATIONSHIP_TYPES.MASTER, '@_Target': fileName });
        parsed.Relationships.Relationship = rels;
        this.pkg.updateFile(relsPath, buildXml(this.builder, parsed));
        return rId;
    }

    /**
     * Append a `<Master>` entry to `visio/masters/masters.xml` for a shape
     * created via `createMaster()` (includes an inline Shapes element).
     */
    private addMasterEntry(
        id: number, name: string, nameU: string, masterShape: any, rId: string
    ): void {
        const mastersPath = 'visio/masters/masters.xml';
        const parsed = this.parser.parse(this.pkg.getFileText(mastersPath));
        if (!parsed.Masters) parsed.Masters = {};
        if (!parsed.Masters.Master) parsed.Masters.Master = [];
        if (!Array.isArray(parsed.Masters.Master)) {
            parsed.Masters.Master = [parsed.Masters.Master];
        }

        parsed.Masters.Master.push({
            '@_ID':    String(id),
            '@_Name':  name,
            '@_NameU': nameU,
            PageSheet: {
                '@_LineStyle': '0',
                '@_FillStyle': '0',
                '@_TextStyle': '0',
                Cell: [
                    { '@_N': 'PageWidth',  '@_V': '0.5' },
                    { '@_N': 'PageHeight', '@_V': '0.5' },
                ],
            },
            Shapes: { Shape: masterShape },
            Rel: { '@_r:id': rId },
        });

        this.pkg.updateFile(mastersPath, buildXml(this.builder, parsed));
    }

    /**
     * Append a `<Master>` entry to `visio/masters/masters.xml` from an imported
     * stencil node. The `shapesNode` may be undefined when the master had no
     * inline Shapes element in the stencil.
     */
    private addMasterEntryFromNode(
        id: number, name: string, nameU: string, shapesNode: any, rId: string
    ): void {
        const mastersPath = 'visio/masters/masters.xml';
        const parsed = this.parser.parse(this.pkg.getFileText(mastersPath));
        if (!parsed.Masters) parsed.Masters = {};
        if (!parsed.Masters.Master) parsed.Masters.Master = [];
        if (!Array.isArray(parsed.Masters.Master)) {
            parsed.Masters.Master = [parsed.Masters.Master];
        }

        const entry: any = {
            '@_ID':    String(id),
            '@_Name':  name,
            '@_NameU': nameU,
            Rel: { '@_r:id': rId },
        };
        if (shapesNode) entry.Shapes = shapesNode;

        parsed.Masters.Master.push(entry);
        this.pkg.updateFile(mastersPath, buildXml(this.builder, parsed));
    }

    /**
     * Build a parametric master shape object (1×1 unit, scales via formulas).
     */
    private buildMasterShapeObject(geometry: ShapeGeometry): any {
        // fillColor='#FFFFFF' so NoFill=0 — the shape is visible by default
        const geomSection = GeometryBuilder.build({
            width: 1, height: 1, geometry, fillColor: '#FFFFFF'
        });
        return {
            '@_ID':        '1',
            '@_Type':      'Shape',
            '@_LineStyle': '0',
            '@_FillStyle': '0',
            '@_TextStyle': '0',
            Cell: [
                { '@_N': 'Width',    '@_V': '1',   '@_F': 'Inh' },
                { '@_N': 'Height',   '@_V': '1',   '@_F': 'Inh' },
                { '@_N': 'PinX',     '@_V': '0.5', '@_F': 'Width*0.5' },
                { '@_N': 'PinY',     '@_V': '0.5', '@_F': 'Height*0.5' },
                { '@_N': 'LocPinX',  '@_V': '0.5', '@_F': 'Width*0.5' },
                { '@_N': 'LocPinY',  '@_V': '0.5', '@_F': 'Height*0.5' },
            ],
            Section: [geomSection],
        };
    }

    /**
     * Build a MasterContents XML string from an inline `<Shapes>` node extracted
     * from a stencil's masters.xml. Used as a fallback when the stencil has no
     * individual masterN.xml file.
     */
    private buildMasterContentsXml(shapesNode: any): string {
        const obj = {
            MasterContents: {
                '@_xmlns':     XML_NAMESPACES.VISIO_MAIN,
                '@_xmlns:r':   XML_NAMESPACES.RELATIONSHIPS_OFFICE,
                '@_xml:space': 'preserve',
                Shapes: shapesNode ?? {},
            }
        };
        return buildXml(this.builder, obj);
    }

    // -------------------------------------------------------------------------
    // Internal utility
    // -------------------------------------------------------------------------

    private loadMastersRels(): Map<string, string> {
        const map = new Map<string, string>();
        try {
            const content = this.pkg.getFileText('visio/masters/_rels/masters.xml.rels');
            const parsed = this.parser.parse(content);
            let rels: any[] = parsed.Relationships?.Relationship ?? [];
            if (!Array.isArray(rels)) rels = rels ? [rels] : [];
            for (const r of rels) map.set(r['@_Id'], r['@_Target']);
        } catch { /* no rels file — no masters */ }
        return map;
    }
}
