import { describe, it, expect } from 'vitest';
import { VisioDocument } from '../src/VisioDocument';

// ---------------------------------------------------------------------------
// doc.getPage(name)
// ---------------------------------------------------------------------------

describe('VisioDocument.getPage(name)', () => {
    it('finds an existing page by name', async () => {
        const doc = await VisioDocument.create();
        await doc.addPage('Dashboard');

        const page = doc.getPage('Dashboard');
        expect(page).toBeDefined();
        expect(page!.name).toBe('Dashboard');
    });

    it('returns undefined for a name that does not exist', async () => {
        const doc = await VisioDocument.create();
        expect(doc.getPage('Nope')).toBeUndefined();
    });

    it('returns the first matching page when names are unique', async () => {
        const doc = await VisioDocument.create();
        await doc.addPage('Alpha');
        await doc.addPage('Beta');

        expect(doc.getPage('Alpha')!.name).toBe('Alpha');
        expect(doc.getPage('Beta')!.name).toBe('Beta');
    });

    it('finds the default page created at document initialisation', async () => {
        const doc = await VisioDocument.create();
        const page = doc.getPage('Page-1');
        expect(page).toBeDefined();
        expect(page!.id).toBe('1');
    });

    it('returns undefined after the page has been deleted', async () => {
        const doc = await VisioDocument.create();
        const p2 = await doc.addPage('Temp');
        await doc.deletePage(p2);
        expect(doc.getPage('Temp')).toBeUndefined();
    });

    it('is case-sensitive', async () => {
        const doc = await VisioDocument.create();
        await doc.addPage('CaseSensitive');
        expect(doc.getPage('casesensitive')).toBeUndefined();
        expect(doc.getPage('CaseSensitive')).toBeDefined();
    });
});

// ---------------------------------------------------------------------------
// page.getShapeById(id)
// ---------------------------------------------------------------------------

describe('Page.getShapeById(id)', () => {
    it('finds a top-level shape by ID', async () => {
        const doc = await VisioDocument.create();
        const page = doc.pages[0];
        const s = await page.addShape({ text: 'Hello', x: 1, y: 1, width: 2, height: 1 });

        const found = page.getShapeById(s.id);
        expect(found).toBeDefined();
        expect(found!.id).toBe(s.id);
        expect(found!.text).toBe('Hello');
    });

    it('returns undefined for a non-existent ID', async () => {
        const doc = await VisioDocument.create();
        const page = doc.pages[0];
        expect(page.getShapeById('9999')).toBeUndefined();
    });

    it('finds a shape nested inside a group', async () => {
        const doc = await VisioDocument.create();
        const page = doc.pages[0];
        const group = await page.addShape({ text: '', x: 3, y: 3, width: 4, height: 4, type: 'Group' });
        const child = await page.addShape({ text: 'Nested', x: 3, y: 3, width: 1, height: 1 }, group.id);

        // getShapes() (top-level only) should NOT find the child
        const topLevel = page.getShapes().map(s => s.id);
        expect(topLevel).not.toContain(child.id);

        // getShapeById() SHOULD find it
        const found = page.getShapeById(child.id);
        expect(found).toBeDefined();
        expect(found!.id).toBe(child.id);
        expect(found!.text).toBe('Nested');
    });

    it('returns correct geometry for a found shape', async () => {
        const doc = await VisioDocument.create();
        const page = doc.pages[0];
        const s = await page.addShape({ text: 'Geo', x: 5, y: 3, width: 2, height: 1 });

        const found = page.getShapeById(s.id)!;
        expect(found.x).toBe(5);
        expect(found.y).toBe(3);
        expect(found.width).toBe(2);
        expect(found.height).toBe(1);
    });

    it('returns undefined after the shape has been deleted', async () => {
        const doc = await VisioDocument.create();
        const page = doc.pages[0];
        const s = await page.addShape({ text: 'Gone', x: 1, y: 1, width: 1, height: 1 });
        await s.delete();

        expect(page.getShapeById(s.id)).toBeUndefined();
    });
});

// ---------------------------------------------------------------------------
// page.findShapes(predicate)
// ---------------------------------------------------------------------------

describe('Page.findShapes(predicate)', () => {
    it('filters top-level shapes by text', async () => {
        const doc = await VisioDocument.create();
        const page = doc.pages[0];
        await page.addShape({ text: 'Alpha', x: 1, y: 1, width: 1, height: 1 });
        await page.addShape({ text: 'Beta',  x: 3, y: 1, width: 1, height: 1 });
        await page.addShape({ text: 'Alpha', x: 5, y: 1, width: 1, height: 1 });

        const alphas = page.findShapes(s => s.text === 'Alpha');
        expect(alphas).toHaveLength(2);
        alphas.forEach(s => expect(s.text).toBe('Alpha'));
    });

    it('returns an empty array when no shapes match', async () => {
        const doc = await VisioDocument.create();
        const page = doc.pages[0];
        await page.addShape({ text: 'Only', x: 1, y: 1, width: 1, height: 1 });

        expect(page.findShapes(s => s.text === 'Missing')).toHaveLength(0);
    });

    it('returns an empty array on an empty page', async () => {
        const doc = await VisioDocument.create();
        const page = doc.pages[0];
        expect(page.findShapes(() => true)).toHaveLength(0);
    });

    it('includes nested shapes in group children', async () => {
        const doc = await VisioDocument.create();
        const page = doc.pages[0];
        const group = await page.addShape({ text: '', x: 3, y: 3, width: 4, height: 4, type: 'Group' });
        await page.addShape({ text: 'Inner', x: 3, y: 3, width: 1, height: 1 }, group.id);
        await page.addShape({ text: 'Inner', x: 5, y: 1, width: 1, height: 1 }); // top-level

        const inners = page.findShapes(s => s.text === 'Inner');
        expect(inners).toHaveLength(2);
    });

    it('can filter by geometry (width)', async () => {
        const doc = await VisioDocument.create();
        const page = doc.pages[0];
        await page.addShape({ text: 'Wide', x: 1, y: 1, width: 5, height: 1 });
        await page.addShape({ text: 'Narrow', x: 1, y: 3, width: 1, height: 1 });

        const wide = page.findShapes(s => s.width > 3);
        expect(wide).toHaveLength(1);
        expect(wide[0].text).toBe('Wide');
    });

    it('returns shapes as proper Shape instances with full API', async () => {
        const doc = await VisioDocument.create();
        const page = doc.pages[0];
        await page.addShape({ text: 'Test', x: 2, y: 2, width: 2, height: 1 });

        const [shape] = page.findShapes(() => true);
        // Verify it has the full Shape API
        expect(typeof shape.setText).toBe('function');
        expect(typeof shape.setStyle).toBe('function');
        expect(typeof shape.delete).toBe('function');
    });
});
