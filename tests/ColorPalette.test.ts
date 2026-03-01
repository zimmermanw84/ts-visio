import { describe, it, expect } from 'vitest';
import { VisioDocument } from '../src/VisioDocument';
import { VisioPackage } from '../src/VisioPackage';
import { XMLParser } from 'fast-xml-parser';

// ── helpers ────────────────────────────────────────────────────────────────────

async function freshDoc() {
    return VisioDocument.create();
}

async function getDocXml(doc: VisioDocument): Promise<any> {
    const buf    = await doc.save();
    const pkg    = new VisioPackage();
    await pkg.load(buf);
    const xml    = pkg.getFileText('visio/document.xml');
    const parser = new XMLParser({ ignoreAttributes: false, attributeNamePrefix: '@_' });
    const parsed = parser.parse(xml);
    return parsed['VisioDocument'];
}

function normArray(val: any): any[] {
    if (!val) return [];
    return Array.isArray(val) ? val : [val];
}

// ── getColors — built-ins ──────────────────────────────────────────────────────

describe('doc.getColors() — built-ins', () => {
    it('new document always has black (index 0) and white (index 1)', async () => {
        const doc    = await freshDoc();
        const colors = doc.getColors();
        const indices = colors.map(c => c.index);
        expect(indices).toContain(0);
        expect(indices).toContain(1);
    });

    it('built-in index 0 is #000000', async () => {
        const doc = await freshDoc();
        const black = doc.getColors().find(c => c.index === 0);
        expect(black?.rgb).toBe('#000000');
    });

    it('built-in index 1 is #FFFFFF', async () => {
        const doc = await freshDoc();
        const white = doc.getColors().find(c => c.index === 1);
        expect(white?.rgb).toBe('#FFFFFF');
    });

    it('colors are ordered by index', async () => {
        const doc = await freshDoc();
        doc.addColor('#AA0000');
        doc.addColor('#00AA00');
        const indices = doc.getColors().map(c => c.index);
        for (let i = 1; i < indices.length; i++) {
            expect(indices[i]).toBeGreaterThan(indices[i - 1]);
        }
    });
});

// ── addColor ───────────────────────────────────────────────────────────────────

describe('doc.addColor()', () => {
    it('returns an index >= 2 for a new user color', async () => {
        const doc = await freshDoc();
        const ix  = doc.addColor('#4472C4');
        expect(ix).toBeGreaterThanOrEqual(2);
    });

    it('sequential colors get sequential indices', async () => {
        const doc = await freshDoc();
        const a   = doc.addColor('#AA0000');
        const b   = doc.addColor('#00AA00');
        const c   = doc.addColor('#0000AA');
        expect(b).toBe(a + 1);
        expect(c).toBe(b + 1);
    });

    it('same color added twice returns the same index', async () => {
        const doc = await freshDoc();
        const ix1 = doc.addColor('#4472C4');
        const ix2 = doc.addColor('#4472C4');
        expect(ix1).toBe(ix2);
    });

    it('adding black (#000000) returns 0 (existing built-in)', async () => {
        const doc = await freshDoc();
        expect(doc.addColor('#000000')).toBe(0);
    });

    it('adding white (#FFFFFF) returns 1 (existing built-in)', async () => {
        const doc = await freshDoc();
        expect(doc.addColor('#FFFFFF')).toBe(1);
    });

    it('newly added color appears in getColors()', async () => {
        const doc = await freshDoc();
        const ix  = doc.addColor('#FF8000');
        const entry = doc.getColors().find(c => c.index === ix);
        expect(entry?.rgb).toBe('#FF8000');
    });

    it('adding multiple distinct colors all appear in getColors()', async () => {
        const doc = await freshDoc();
        doc.addColor('#111111');
        doc.addColor('#222222');
        doc.addColor('#333333');
        const rgbs = doc.getColors().map(c => c.rgb);
        expect(rgbs).toContain('#111111');
        expect(rgbs).toContain('#222222');
        expect(rgbs).toContain('#333333');
    });
});

// ── hex normalisation ──────────────────────────────────────────────────────────

describe('hex normalisation', () => {
    it('lowercase hex is normalised to uppercase', async () => {
        const doc = await freshDoc();
        const ix  = doc.addColor('#4472c4');
        const entry = doc.getColors().find(c => c.index === ix);
        expect(entry?.rgb).toBe('#4472C4');
    });

    it('shorthand #RGB expands to #RRGGBB', async () => {
        const doc = await freshDoc();
        const ix  = doc.addColor('#F0F');
        const entry = doc.getColors().find(c => c.index === ix);
        expect(entry?.rgb).toBe('#FF00FF');
    });

    it('shorthand and longhand of the same color are de-duplicated', async () => {
        const doc = await freshDoc();
        const ix1 = doc.addColor('#FFF');
        const ix2 = doc.addColor('#FFFFFF');
        expect(ix1).toBe(ix2);
    });

    it('hex without leading # is accepted', async () => {
        const doc = await freshDoc();
        const ix  = doc.addColor('AA1122');
        expect(ix).toBeGreaterThanOrEqual(2);
        const entry = doc.getColors().find(c => c.index === ix);
        expect(entry?.rgb).toBe('#AA1122');
    });

    it('case-insensitive de-duplication: #AABBCC and #aabbcc are the same', async () => {
        const doc = await freshDoc();
        const ix1 = doc.addColor('#AABBCC');
        const ix2 = doc.addColor('#aabbcc');
        expect(ix1).toBe(ix2);
        const entries = doc.getColors().filter(c => c.rgb === '#AABBCC');
        expect(entries).toHaveLength(1);
    });
});

// ── getColorIndex ─────────────────────────────────────────────────────────────

describe('doc.getColorIndex()', () => {
    it('returns the correct index for a registered color', async () => {
        const doc = await freshDoc();
        const ix  = doc.addColor('#4472C4');
        expect(doc.getColorIndex('#4472C4')).toBe(ix);
    });

    it('returns undefined for an unregistered color', async () => {
        const doc = await freshDoc();
        expect(doc.getColorIndex('#123456')).toBeUndefined();
    });

    it('returns 0 for #000000 (built-in black)', async () => {
        const doc = await freshDoc();
        expect(doc.getColorIndex('#000000')).toBe(0);
    });

    it('returns 1 for #FFFFFF (built-in white)', async () => {
        const doc = await freshDoc();
        expect(doc.getColorIndex('#FFFFFF')).toBe(1);
    });

    it('is case-insensitive', async () => {
        const doc = await freshDoc();
        doc.addColor('#4472C4');
        expect(doc.getColorIndex('#4472c4')).toBeDefined();
    });
});

// ── document.xml structure ────────────────────────────────────────────────────

describe('color palette in document.xml', () => {
    it('new document xml has <Colors> with IX=0 and IX=1', async () => {
        const doc  = await freshDoc();
        const vDoc = await getDocXml(doc);
        const entries = normArray(vDoc?.Colors?.ColorEntry);
        const ixValues = entries.map((e: any) => e['@_IX']);
        expect(ixValues).toContain('0');
        expect(ixValues).toContain('1');
    });

    it('addColor persists the entry to document.xml', async () => {
        const doc = await freshDoc();
        doc.addColor('#4472C4');
        const vDoc    = await getDocXml(doc);
        const entries = normArray(vDoc?.Colors?.ColorEntry);
        const match   = entries.find((e: any) => e['@_RGB'] === '#4472C4');
        expect(match).toBeDefined();
    });

    it('IX attribute in XML matches the returned index', async () => {
        const doc = await freshDoc();
        const ix  = doc.addColor('#CC0000');
        const vDoc    = await getDocXml(doc);
        const entries = normArray(vDoc?.Colors?.ColorEntry);
        const match   = entries.find((e: any) => e['@_RGB'] === '#CC0000');
        expect(match?.['@_IX']).toBe(ix.toString());
    });

    it('duplicate addColor does not create two XML entries', async () => {
        const doc = await freshDoc();
        doc.addColor('#4472C4');
        doc.addColor('#4472C4');
        const vDoc    = await getDocXml(doc);
        const entries = normArray(vDoc?.Colors?.ColorEntry);
        const matches = entries.filter((e: any) => e['@_RGB'] === '#4472C4');
        expect(matches).toHaveLength(1);
    });
});

// ── round-trip ────────────────────────────────────────────────────────────────

describe('color palette round-trip save/reload', () => {
    it('added colors survive save/reload', async () => {
        const doc1 = await freshDoc();
        const ix   = doc1.addColor('#4472C4');
        const buf  = await doc1.save();

        const doc2   = await VisioDocument.load(buf);
        const colors = doc2.getColors();
        const entry  = colors.find(c => c.index === ix);
        expect(entry?.rgb).toBe('#4472C4');
    });

    it('multiple colors survive save/reload', async () => {
        const doc1 = await freshDoc();
        const ixA  = doc1.addColor('#111111');
        const ixB  = doc1.addColor('#222222');
        const ixC  = doc1.addColor('#333333');
        const buf  = await doc1.save();

        const doc2   = await VisioDocument.load(buf);
        const colors = doc2.getColors();
        expect(colors.find(c => c.index === ixA)?.rgb).toBe('#111111');
        expect(colors.find(c => c.index === ixB)?.rgb).toBe('#222222');
        expect(colors.find(c => c.index === ixC)?.rgb).toBe('#333333');
    });

    it('getColorIndex works correctly after reload', async () => {
        const doc1 = await freshDoc();
        const ix   = doc1.addColor('#ABCDEF');
        const buf  = await doc1.save();

        const doc2 = await VisioDocument.load(buf);
        expect(doc2.getColorIndex('#ABCDEF')).toBe(ix);
    });

    it('addColor after reload continues from the correct next index', async () => {
        const doc1 = await freshDoc();
        const ix1  = doc1.addColor('#111111');
        const buf  = await doc1.save();

        const doc2 = await VisioDocument.load(buf);
        const ix2  = doc2.addColor('#222222');
        expect(ix2).toBe(ix1 + 1);
    });
});
