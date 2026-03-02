import { describe, it, expect, afterEach } from 'vitest';
import { VisioDocument } from '../src/VisioDocument';
import { VisioPackage } from '../src/VisioPackage';
import { XMLParser } from 'fast-xml-parser';
import JSZip from 'jszip';
import fs from 'fs';
import path from 'path';

const parser = new XMLParser({ ignoreAttributes: false, attributeNamePrefix: '@_' });

// ---------------------------------------------------------------------------
// Helper: build a minimal in-memory .vssx buffer containing the given masters
// ---------------------------------------------------------------------------
interface StencilMasterDef {
    id: string;
    name: string;
    nameU?: string;
}

async function buildMinimalVssx(masterDefs: StencilMasterDef[]): Promise<Buffer> {
    const zip = new JSZip();

    // Build masters.xml and individual masterN.xml files
    const masterEntries = masterDefs.map((def, i) => {
        const rId = `rId${i + 1}`;
        const fileName = `master${i + 1}.xml`;
        return { def, rId, fileName };
    });

    // masters.xml
    const mastersEntries = masterEntries.map(({ def, rId }) => `
        <Master ID="${def.id}" Name="${def.name}" NameU="${def.nameU ?? def.name}">
            <Shapes>
                <Shape ID="1" Type="Shape" LineStyle="0" FillStyle="0" TextStyle="0">
                    <Cell N="Width" V="1" F="Inh"/>
                    <Cell N="Height" V="1" F="Inh"/>
                </Shape>
            </Shapes>
            <Rel r:id="${rId}"/>
        </Master>`).join('');

    zip.file('visio/masters/masters.xml',
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Masters xmlns="http://schemas.microsoft.com/office/visio/2012/main"
         xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
         xml:space="preserve">${mastersEntries}</Masters>`);

    // _rels/masters.xml.rels
    const relEntries = masterEntries.map(({ rId, fileName }) =>
        `<Relationship Id="${rId}" Type="http://schemas.microsoft.com/visio/2010/relationships/master" Target="${fileName}"/>`
    ).join('');
    zip.file('visio/masters/_rels/masters.xml.rels',
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
${relEntries}
</Relationships>`);

    // Individual masterN.xml files
    for (const { fileName } of masterEntries) {
        zip.file(`visio/masters/${fileName}`,
            `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<MasterContents xmlns="http://schemas.microsoft.com/office/visio/2012/main"
                xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                xml:space="preserve">
    <Shapes>
        <Shape ID="1" Type="Shape" LineStyle="0" FillStyle="0" TextStyle="0">
            <Cell N="Width" V="1" F="Inh"/>
            <Cell N="Height" V="1" F="Inh"/>
        </Shape>
    </Shapes>
</MasterContents>`);
    }

    return zip.generateAsync({ type: 'nodebuffer' });
}

// ---------------------------------------------------------------------------

describe('doc.importMastersFromStencil()', () => {
    const testFile = path.resolve(__dirname, 'stencil_import_test.vsdx');
    afterEach(() => { if (fs.existsSync(testFile)) fs.unlinkSync(testFile); });

    it('imports a single master from a .vssx buffer', async () => {
        const stencilBuf = await buildMinimalVssx([{ id: '1', name: 'Arrow' }]);
        const doc = await VisioDocument.create();
        const imported = await doc.importMastersFromStencil(stencilBuf);

        expect(imported).toHaveLength(1);
        expect(imported[0].name).toBe('Arrow');
        expect(imported[0].id).toBe('1');
        expect(imported[0].xmlPath).toBe('visio/masters/master1.xml');
    });

    it('imports multiple masters and assigns non-conflicting IDs', async () => {
        const stencilBuf = await buildMinimalVssx([
            { id: '1', name: 'Box' },
            { id: '2', name: 'Circle' },
            { id: '3', name: 'Diamond' },
        ]);
        const doc = await VisioDocument.create();
        const imported = await doc.importMastersFromStencil(stencilBuf);

        expect(imported).toHaveLength(3);
        // IDs should be sequential starting from 1
        const ids = imported.map(m => parseInt(m.id)).sort((a, b) => a - b);
        expect(ids).toEqual([1, 2, 3]);
        expect(imported.map(m => m.name)).toEqual(['Box', 'Circle', 'Diamond']);
    });

    it('avoids ID collisions when document already has masters', async () => {
        const doc = await VisioDocument.create();
        doc.createMaster('Existing');  // Gets ID 1

        const stencilBuf = await buildMinimalVssx([{ id: '1', name: 'Imported' }]);
        const imported = await doc.importMastersFromStencil(stencilBuf);

        expect(imported[0].id).toBe('2'); // no collision with existing ID 1
        const allMasters = doc.getMasters();
        expect(allMasters).toHaveLength(2);
    });

    it('getMasters() reflects imported masters', async () => {
        const stencilBuf = await buildMinimalVssx([
            { id: '1', name: 'Rect' },
            { id: '2', name: 'Oval' },
        ]);
        const doc = await VisioDocument.create();
        await doc.importMastersFromStencil(stencilBuf);

        const masters = doc.getMasters();
        expect(masters).toHaveLength(2);
        expect(masters.map(m => m.name).sort()).toEqual(['Oval', 'Rect']);
    });

    it('writes individual masterN.xml files', async () => {
        const stencilBuf = await buildMinimalVssx([{ id: '1', name: 'Box' }]);
        const doc = await VisioDocument.create();
        await doc.importMastersFromStencil(stencilBuf);

        const pkg = (doc as any).pkg as VisioPackage;
        const masterXml = pkg.filesMap.get('visio/masters/master1.xml') as string;
        expect(masterXml).toBeDefined();
        const parsed = parser.parse(masterXml);
        expect(parsed.MasterContents).toBeDefined();
    });

    it('can use an imported master to stamp a shape instance', async () => {
        const stencilBuf = await buildMinimalVssx([{ id: '1', name: 'NetworkDevice' }]);
        const doc = await VisioDocument.create();
        const imported = await doc.importMastersFromStencil(stencilBuf);

        const page = doc.pages[0];
        const shape = await page.addShape({
            text: 'Router',
            x: 2, y: 3, width: 1, height: 1,
            masterId: imported[0].id,
        });
        expect(shape).toBeDefined();

        const pkg = (doc as any).pkg as VisioPackage;
        const pageXml = pkg.filesMap.get('visio/pages/page1.xml') as string;
        expect(pageXml).toContain(`Master="${imported[0].id}"`);
    });

    it('returns empty array for a .vssx with no masters', async () => {
        const zip = new JSZip();
        zip.file('visio/masters/masters.xml',
            `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Masters xmlns="http://schemas.microsoft.com/office/visio/2012/main" xml:space="preserve"/>
`);
        const buf = await zip.generateAsync({ type: 'nodebuffer' });
        const doc = await VisioDocument.create();
        const imported = await doc.importMastersFromStencil(buf);
        expect(imported).toHaveLength(0);
    });

    it('returns empty array for a .vssx with no masters.xml', async () => {
        const zip = new JSZip();
        zip.file('dummy.txt', 'not a stencil');
        const buf = await zip.generateAsync({ type: 'nodebuffer' });
        const doc = await VisioDocument.create();
        const imported = await doc.importMastersFromStencil(buf);
        expect(imported).toHaveLength(0);
    });

    it('round-trips through save/load', async () => {
        const stencilBuf = await buildMinimalVssx([
            { id: '1', name: 'Box' },
            { id: '2', name: 'Circle' },
        ]);
        const doc = await VisioDocument.create();
        await doc.importMastersFromStencil(stencilBuf);
        await doc.save(testFile);

        const doc2 = await VisioDocument.load(testFile);
        const m2 = doc2.getMasters();
        expect(m2).toHaveLength(2);
        expect(m2.map(m => m.name).sort()).toEqual(['Box', 'Circle']);
    });
});
