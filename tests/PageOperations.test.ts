import { describe, it, expect } from 'vitest';
import { VisioDocument } from '../src/VisioDocument';

async function freshDoc() {
    return VisioDocument.create();
}

// ── renamePage ─────────────────────────────────────────────────────────────────

describe('doc.renamePage()', () => {
    it('updates page.name in memory immediately', async () => {
        const doc = await freshDoc();
        const page = doc.pages[0];
        doc.renamePage(page, 'Renamed');
        expect(page.name).toBe('Renamed');
    });

    it('is reflected in doc.pages after rename', async () => {
        const doc = await freshDoc();
        const page = doc.pages[0];
        doc.renamePage(page, 'Updated Name');
        expect(doc.pages[0].name).toBe('Updated Name');
    });

    it('survives save/reload', async () => {
        const doc = await freshDoc();
        doc.renamePage(doc.pages[0], 'Persisted');
        const buf = await doc.save();
        const doc2 = await VisioDocument.load(buf);
        expect(doc2.pages[0].name).toBe('Persisted');
    });

    it('getPage() finds the page by new name after rename', async () => {
        const doc = await freshDoc();
        doc.renamePage(doc.pages[0], 'FindMe');
        expect(doc.getPage('FindMe')).toBeDefined();
        expect(doc.getPage('Page-1')).toBeUndefined();
    });

    it('throws if page id is unknown', async () => {
        const doc = await freshDoc();
        // Construct a fake page reference by adding a page, deleting it,
        // then try to rename the deleted page via the manager directly.
        const page = await doc.addPage('Temp');
        await doc.deletePage(page);
        expect(() => doc.renamePage(page, 'Ghost')).toThrow();
    });
});

// ── movePage ───────────────────────────────────────────────────────────────────

describe('doc.movePage()', () => {
    it('moves a page to the front', async () => {
        const doc = await freshDoc();
        await doc.addPage('B');
        await doc.addPage('C');
        const [, b] = doc.pages;
        doc.movePage(b, 0);
        expect(doc.pages[0].name).toBe('B');
    });

    it('moves a page to the end', async () => {
        const doc = await freshDoc();
        await doc.addPage('B');
        await doc.addPage('C');
        const [a] = doc.pages;
        doc.movePage(a, 2);
        expect(doc.pages[doc.pages.length - 1].name).toBe(a.name);
    });

    it('clamps toIndex below 0 to 0', async () => {
        const doc = await freshDoc();
        await doc.addPage('B');
        const [, b] = doc.pages;
        doc.movePage(b, -99);
        expect(doc.pages[0].name).toBe('B');
    });

    it('clamps toIndex above last to last', async () => {
        const doc = await freshDoc();
        await doc.addPage('B');
        const [a] = doc.pages;
        doc.movePage(a, 999);
        expect(doc.pages[doc.pages.length - 1].name).toBe(a.name);
    });

    it('page order survives save/reload', async () => {
        const doc = await freshDoc();
        await doc.addPage('B');
        await doc.addPage('C');
        const [, , c] = doc.pages;
        doc.movePage(c, 0);

        const buf = await doc.save();
        const doc2 = await VisioDocument.load(buf);
        expect(doc2.pages[0].name).toBe('C');
    });
});

// ── duplicatePage ──────────────────────────────────────────────────────────────

describe('doc.duplicatePage()', () => {
    it('returns a new Page with the given name', async () => {
        const doc = await freshDoc();
        const copy = await doc.duplicatePage(doc.pages[0], 'Clone');
        expect(copy.name).toBe('Clone');
    });

    it('uses "<name> (Copy)" when no name is supplied', async () => {
        const doc = await freshDoc();
        const [original] = doc.pages;
        const copy = await doc.duplicatePage(original);
        expect(copy.name).toBe(`${original.name} (Copy)`);
    });

    it('increments doc.pages length by one', async () => {
        const doc = await freshDoc();
        const before = doc.pages.length;
        await doc.duplicatePage(doc.pages[0], 'Dup');
        expect(doc.pages.length).toBe(before + 1);
    });

    it('duplicate appears directly after the source in tab order', async () => {
        const doc = await freshDoc();
        await doc.addPage('B');
        const [a] = doc.pages;
        await doc.duplicatePage(a, 'A-Copy');
        expect(doc.pages[0].name).toBe(a.name);
        expect(doc.pages[1].name).toBe('A-Copy');
    });

    it('copy has a distinct ID from the source', async () => {
        const doc = await freshDoc();
        const [src] = doc.pages;
        const copy = await doc.duplicatePage(src, 'Dup');
        expect(copy.id).not.toBe(src.id);
    });

    it('shapes added to source appear on the copy after save/reload', async () => {
        const doc = await freshDoc();
        const [src] = doc.pages;
        await src.addShape({ text: 'Hello', x: 1, y: 1, width: 2, height: 1 });
        const copy = await doc.duplicatePage(src, 'Copy');

        const buf = await doc.save();
        const doc2 = await VisioDocument.load(buf);

        // Original source page still has the shape
        const srcShapes = doc2.pages.find(p => p.id === src.id)!.getShapes();
        expect(srcShapes.length).toBeGreaterThan(0);

        // Copy also has the shape (XML was copied)
        const copyShapes = doc2.pages.find(p => p.id === copy.id)!.getShapes();
        expect(copyShapes.length).toBe(srcShapes.length);
    });

    it('duplicate survives save/reload with the correct name', async () => {
        const doc = await freshDoc();
        await doc.duplicatePage(doc.pages[0], 'Duplicate');

        const buf = await doc.save();
        const doc2 = await VisioDocument.load(buf);
        expect(doc2.pages.some(p => p.name === 'Duplicate')).toBe(true);
    });

    it('can duplicate multiple times independently', async () => {
        const doc = await freshDoc();
        const [src] = doc.pages;
        await doc.duplicatePage(src, 'Copy1');
        await doc.duplicatePage(src, 'Copy2');
        expect(doc.pages.length).toBe(3);
        const names = doc.pages.map(p => p.name);
        expect(names).toContain('Copy1');
        expect(names).toContain('Copy2');
    });
});
