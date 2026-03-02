import { describe, it, expect, afterEach } from 'vitest';
import { VisioDocument } from '../src/VisioDocument';
import { VisioPackage } from '../src/VisioPackage';
import { XMLParser } from 'fast-xml-parser';
import fs from 'fs';
import path from 'path';

const parser = new XMLParser({ ignoreAttributes: false, attributeNamePrefix: '@_' });

function parseText(pkg: VisioPackage, filePath: string): any {
    return parser.parse(pkg.filesMap.get(filePath) as string ?? '');
}

describe('doc.createMaster()', () => {
    const testFile = path.resolve(__dirname, 'master_creation_test.vsdx');
    afterEach(() => { if (fs.existsSync(testFile)) fs.unlinkSync(testFile); });

    it('returns a MasterRecord with correct fields', async () => {
        const doc = await VisioDocument.create();
        const rec = doc.createMaster('Box');
        expect(rec.id).toBe('1');
        expect(rec.name).toBe('Box');
        expect(rec.nameU).toBe('Box');
        expect(rec.xmlPath).toBe('visio/masters/master1.xml');
    });

    it('assigns sequential IDs for multiple masters', async () => {
        const doc = await VisioDocument.create();
        const r1 = doc.createMaster('Box');
        const r2 = doc.createMaster('Circle', 'ellipse');
        const r3 = doc.createMaster('Diamond', 'diamond');
        expect(r1.id).toBe('1');
        expect(r2.id).toBe('2');
        expect(r3.id).toBe('3');
    });

    it('getMasters() returns the created masters', async () => {
        const doc = await VisioDocument.create();
        doc.createMaster('Box');
        doc.createMaster('Circle', 'ellipse');
        const masters = doc.getMasters();
        expect(masters).toHaveLength(2);
        expect(masters.map(m => m.name)).toEqual(['Box', 'Circle']);
    });

    it('creates visio/masters/masters.xml with master entry', async () => {
        const doc = await VisioDocument.create();
        doc.createMaster('Box');

        // Access the internal pkg to inspect written files
        const pkg = (doc as any).pkg as VisioPackage;
        const mastersXml = pkg.filesMap.get('visio/masters/masters.xml') as string;
        expect(mastersXml).toBeDefined();
        const parsed = parser.parse(mastersXml);
        const masterNode = parsed.Masters.Master;
        const node = Array.isArray(masterNode) ? masterNode[0] : masterNode;
        expect(node['@_ID']).toBe('1');
        expect(node['@_Name']).toBe('Box');
    });

    it('creates individual masterN.xml (MasterContents)', async () => {
        const doc = await VisioDocument.create();
        doc.createMaster('Box');

        const pkg = (doc as any).pkg as VisioPackage;
        const masterXml = pkg.filesMap.get('visio/masters/master1.xml') as string;
        expect(masterXml).toBeDefined();
        const parsed = parser.parse(masterXml);
        expect(parsed.MasterContents).toBeDefined();
        expect(parsed.MasterContents.Shapes).toBeDefined();
    });

    it('creates visio/masters/_rels/masters.xml.rels', async () => {
        const doc = await VisioDocument.create();
        doc.createMaster('Box');

        const pkg = (doc as any).pkg as VisioPackage;
        const relsXml = pkg.filesMap.get('visio/masters/_rels/masters.xml.rels') as string;
        expect(relsXml).toBeDefined();
        const parsed = parser.parse(relsXml);
        const rel = parsed.Relationships.Relationship;
        const relNode = Array.isArray(rel) ? rel[0] : rel;
        expect(relNode['@_Target']).toBe('master1.xml');
    });

    it('adds master content-type override to [Content_Types].xml', async () => {
        const doc = await VisioDocument.create();
        doc.createMaster('Box');

        const pkg = (doc as any).pkg as VisioPackage;
        const ctXml = pkg.filesMap.get('[Content_Types].xml') as string;
        const parsed = parser.parse(ctXml);
        const overrides: any[] = parsed.Types.Override;
        const masterCT = overrides.find((o: any) => o['@_PartName'] === '/visio/masters/master1.xml');
        expect(masterCT).toBeDefined();
        expect(masterCT['@_ContentType']).toBe('application/vnd.ms-visio.master+xml');
        const mastersCT = overrides.find((o: any) => o['@_PartName'] === '/visio/masters/masters.xml');
        expect(mastersCT).toBeDefined();
        expect(mastersCT['@_ContentType']).toBe('application/vnd.ms-visio.masters+xml');
    });

    it('adds masters relationship to visio/_rels/document.xml.rels', async () => {
        const doc = await VisioDocument.create();
        doc.createMaster('Box');

        const pkg = (doc as any).pkg as VisioPackage;
        const relsXml = pkg.filesMap.get('visio/_rels/document.xml.rels') as string;
        expect(relsXml).toContain('masters/masters.xml');
    });

    it('master shape object has a Geometry Section', async () => {
        const doc = await VisioDocument.create();
        doc.createMaster('Box');

        const pkg = (doc as any).pkg as VisioPackage;
        const masterXml = pkg.filesMap.get('visio/masters/master1.xml') as string;
        const parsed = parser.parse(masterXml);
        const shape = parsed.MasterContents.Shapes.Shape;
        const sections: any[] = Array.isArray(shape.Section) ? shape.Section : [shape.Section];
        const geom = sections.find((s: any) => s['@_N'] === 'Geometry');
        expect(geom).toBeDefined();
    });

    it('can stamp a master instance on a page and shape has @_Master attribute', async () => {
        const doc = await VisioDocument.create();
        const rec = doc.createMaster('Box');
        const page = doc.pages[0];
        const shape = await page.addShape({ text: 'Instance', x: 1, y: 1, width: 2, height: 1, masterId: rec.id });
        expect(shape).toBeDefined();

        const pkg = (doc as any).pkg as VisioPackage;
        const pageXml = pkg.filesMap.get('visio/pages/page1.xml') as string;
        expect(pageXml).toContain(`Master="${rec.id}"`);
    });

    it('master instances do not have a Geometry section (inherited from master)', async () => {
        const doc = await VisioDocument.create();
        const rec = doc.createMaster('Box');
        const page = doc.pages[0];
        await page.addShape({ text: '', x: 1, y: 1, width: 1, height: 1, masterId: rec.id });

        const pkg = (doc as any).pkg as VisioPackage;
        const pageXml = pkg.filesMap.get('visio/pages/page1.xml') as string;
        const parsedPage = parser.parse(pageXml);
        const shape = parsedPage.PageContents.Shapes.Shape;

        let sections = shape.Section ?? [];
        if (!Array.isArray(sections)) sections = [sections];
        const geom = sections.find((s: any) => s['@_N'] === 'Geometry');
        expect(geom).toBeUndefined();
    });

    it('supports ellipse geometry', async () => {
        const doc = await VisioDocument.create();
        doc.createMaster('Process', 'ellipse');

        const pkg = (doc as any).pkg as VisioPackage;
        const masterXml = pkg.filesMap.get('visio/masters/master1.xml') as string;
        const parsed = parser.parse(masterXml);
        const shape = parsed.MasterContents.Shapes.Shape;
        const sections: any[] = Array.isArray(shape.Section) ? shape.Section : [shape.Section];
        const geom = sections.find((s: any) => s['@_N'] === 'Geometry');
        expect(geom).toBeDefined();
        // Ellipse uses a single Ellipse row rather than LineTo rows
        const rows: any[] = Array.isArray(geom.Row) ? geom.Row : [geom.Row];
        expect(rows[0]['@_T']).toBe('Ellipse');
    });

    it('produces a valid .vsdx that round-trips through VisioDocument.load()', async () => {
        const doc = await VisioDocument.create();
        doc.createMaster('Box');
        doc.createMaster('Circle', 'ellipse');
        const page = doc.pages[0];
        const masters = doc.getMasters();
        await page.addShape({ text: '', x: 1, y: 1, width: 1, height: 1, masterId: masters[0].id });

        const buf = await doc.save(testFile);
        expect(buf.length).toBeGreaterThan(0);

        const doc2 = await VisioDocument.load(testFile);
        const m2 = doc2.getMasters();
        expect(m2).toHaveLength(2);
        expect(m2[0].name).toBe('Box');
        expect(m2[1].name).toBe('Circle');
    });
});
