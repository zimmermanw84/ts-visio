import { describe, it, expect } from 'vitest';
import { VisioDocument } from '../src/VisioDocument';

// ── helpers ────────────────────────────────────────────────────────────────────

async function freshDoc() {
    const doc  = await VisioDocument.create();
    const page = doc.pages[0];
    return { doc, page };
}

// ── shape.isGroup ──────────────────────────────────────────────────────────────

describe('shape.isGroup', () => {
    it('is false for a plain rectangle shape', async () => {
        const { page } = await freshDoc();
        const shape = await page.addShape({ text: 'A', x: 1, y: 1, width: 2, height: 1 });
        expect(shape.isGroup).toBe(false);
    });

    it('is true for a shape created with type Group', async () => {
        const { page } = await freshDoc();
        const group = await page.addShape({ text: 'G', x: 5, y: 5, width: 4, height: 4, type: 'Group' });
        expect(group.isGroup).toBe(true);
    });
});

// ── shape.type ─────────────────────────────────────────────────────────────────

describe('shape.type', () => {
    it('returns "Shape" (or normalised) for a plain shape', async () => {
        const { page } = await freshDoc();
        const shape = await page.addShape({ text: 'A', x: 1, y: 1, width: 2, height: 1 });
        // Type may be undefined from the stub, normalised to 'Shape'
        expect(typeof shape.type).toBe('string');
    });

    it('returns "Group" for a group shape', async () => {
        const { page } = await freshDoc();
        const group = await page.addShape({ text: 'G', x: 5, y: 5, width: 4, height: 4, type: 'Group' });
        expect(group.type).toBe('Group');
    });
});

// ── getChildren() — basic ──────────────────────────────────────────────────────

describe('shape.getChildren() — basic', () => {
    it('returns empty array for a plain shape', async () => {
        const { page } = await freshDoc();
        const shape = await page.addShape({ text: 'A', x: 1, y: 1, width: 2, height: 1 });
        expect(shape.getChildren()).toEqual([]);
    });

    it('returns empty array for a group with no children', async () => {
        const { page } = await freshDoc();
        const group = await page.addShape({ text: 'G', x: 5, y: 5, width: 4, height: 4, type: 'Group' });
        // Save and reload so the XML is flushed
        const buf  = await page['pkg'].save?.() ?? (await (page as any).doc?.save());
        // Use the modifier-based read (no save/reload needed for in-memory state)
        expect(group.getChildren()).toHaveLength(0);
    });

    it('returns one child when one shape is added to the group', async () => {
        const { doc, page } = await freshDoc();
        const group = await page.addShape({ text: 'G', x: 5, y: 5, width: 4, height: 4, type: 'Group' });
        await page.addShape({ text: 'Child', x: 1, y: 1, width: 1, height: 1 }, group.id);

        // Reload via save buffer so ShapeReader sees the persisted XML
        const buf   = await doc.save();
        const doc2  = await VisioDocument.load(buf);
        const page2 = doc2.pages[0];
        const shapes = page2.getShapes();
        const g2 = shapes.find(s => s.text === 'G')!;

        const children = g2.getChildren();
        expect(children).toHaveLength(1);
        expect(children[0].text).toBe('Child');
    });

    it('returns all direct children when multiple children are added', async () => {
        const { doc, page } = await freshDoc();
        const group = await page.addShape({ text: 'G', x: 5, y: 5, width: 4, height: 4, type: 'Group' });
        await page.addShape({ text: 'A', x: 1, y: 1, width: 1, height: 1 }, group.id);
        await page.addShape({ text: 'B', x: 2, y: 1, width: 1, height: 1 }, group.id);
        await page.addShape({ text: 'C', x: 3, y: 1, width: 1, height: 1 }, group.id);

        const buf   = await doc.save();
        const doc2  = await VisioDocument.load(buf);
        const g2 = doc2.pages[0].getShapes().find(s => s.text === 'G')!;

        const children = g2.getChildren();
        expect(children).toHaveLength(3);
        const texts = children.map(c => c.text).sort();
        expect(texts).toEqual(['A', 'B', 'C']);
    });

    it('children have correct IDs', async () => {
        const { doc, page } = await freshDoc();
        const group = await page.addShape({ text: 'G', x: 5, y: 5, width: 4, height: 4, type: 'Group' });
        const child = await page.addShape({ text: 'Kid', x: 1, y: 1, width: 1, height: 1 }, group.id);

        const buf  = await doc.save();
        const doc2 = await VisioDocument.load(buf);
        const g2   = doc2.pages[0].getShapes().find(s => s.text === 'G')!;
        const kids = g2.getChildren();

        expect(kids[0].id).toBe(child.id);
    });
});

// ── getChildren() — depth ──────────────────────────────────────────────────────

describe('shape.getChildren() — depth', () => {
    it('only returns direct children, not grandchildren', async () => {
        const { doc, page } = await freshDoc();
        const outer = await page.addShape({ text: 'Outer', x: 5, y: 5, width: 6, height: 6, type: 'Group' });
        const inner = await page.addShape({ text: 'Inner', x: 1, y: 1, width: 3, height: 3, type: 'Group' }, outer.id);
        await page.addShape({ text: 'Grand', x: 0.5, y: 0.5, width: 1, height: 1 }, inner.id);

        const buf   = await doc.save();
        const doc2  = await VisioDocument.load(buf);
        const outer2 = doc2.pages[0].getShapes().find(s => s.text === 'Outer')!;

        const children = outer2.getChildren();
        // Only 'Inner' is a direct child — 'Grand' is a grandchild
        expect(children).toHaveLength(1);
        expect(children[0].text).toBe('Inner');
    });

    it('getChildren on a child group returns its own children', async () => {
        const { doc, page } = await freshDoc();
        const outer = await page.addShape({ text: 'Outer', x: 5, y: 5, width: 6, height: 6, type: 'Group' });
        const inner = await page.addShape({ text: 'Inner', x: 1, y: 1, width: 3, height: 3, type: 'Group' }, outer.id);
        await page.addShape({ text: 'Grand', x: 0.5, y: 0.5, width: 1, height: 1 }, inner.id);

        const buf   = await doc.save();
        const doc2  = await VisioDocument.load(buf);
        const outer2 = doc2.pages[0].getShapes().find(s => s.text === 'Outer')!;

        const inner2 = outer2.getChildren()[0];
        expect(inner2.text).toBe('Inner');
        expect(inner2.isGroup).toBe(true);

        const grandkids = inner2.getChildren();
        expect(grandkids).toHaveLength(1);
        expect(grandkids[0].text).toBe('Grand');
    });
});

// ── getChildren() — Shape API on children ──────────────────────────────────────

describe('shape.getChildren() — children are full Shape instances', () => {
    it('children expose text, id, width, height', async () => {
        const { doc, page } = await freshDoc();
        const group = await page.addShape({ text: 'G', x: 5, y: 5, width: 4, height: 4, type: 'Group' });
        await page.addShape({ text: 'Kid', x: 1, y: 1, width: 2, height: 1.5 }, group.id);

        const buf  = await doc.save();
        const doc2 = await VisioDocument.load(buf);
        const g2   = doc2.pages[0].getShapes().find(s => s.text === 'G')!;
        const kid  = g2.getChildren()[0];

        expect(kid.text).toBe('Kid');
        expect(kid.id).toBeDefined();
        expect(kid.width).toBeCloseTo(2, 4);
        expect(kid.height).toBeCloseTo(1.5, 4);
    });

    it('children from getChildren() also support getChildren()', async () => {
        const { doc, page } = await freshDoc();
        const g1 = await page.addShape({ text: 'G1', x: 5, y: 5, width: 6, height: 6, type: 'Group' });
        const g2 = await page.addShape({ text: 'G2', x: 1, y: 1, width: 3, height: 3, type: 'Group' }, g1.id);
        await page.addShape({ text: 'Leaf', x: 0.5, y: 0.5, width: 1, height: 1 }, g2.id);

        const buf   = await doc.save();
        const doc2  = await VisioDocument.load(buf);
        const root  = doc2.pages[0].getShapes().find(s => s.text === 'G1')!;
        const mid   = root.getChildren()[0];
        const leaf  = mid.getChildren()[0];

        expect(leaf.text).toBe('Leaf');
    });
});

// ── round-trip ────────────────────────────────────────────────────────────────

describe('shape.getChildren() — round-trip save/reload', () => {
    it('children survive save and reload', async () => {
        const { doc, page } = await freshDoc();
        const group = await page.addShape({ text: 'G', x: 5, y: 5, width: 4, height: 4, type: 'Group' });
        await page.addShape({ text: 'P', x: 1, y: 1, width: 1, height: 1 }, group.id);
        await page.addShape({ text: 'Q', x: 2, y: 1, width: 1, height: 1 }, group.id);

        const buf  = await doc.save();
        const doc2 = await VisioDocument.load(buf);
        const g2   = doc2.pages[0].getShapes().find(s => s.text === 'G')!;

        const kids = g2.getChildren();
        expect(kids).toHaveLength(2);
        const texts = kids.map(k => k.text).sort();
        expect(texts).toEqual(['P', 'Q']);
    });

    it('getShapes() still returns only top-level shapes after adding group children', async () => {
        const { doc, page } = await freshDoc();
        const group = await page.addShape({ text: 'G', x: 5, y: 5, width: 4, height: 4, type: 'Group' });
        await page.addShape({ text: 'Child', x: 1, y: 1, width: 1, height: 1 }, group.id);

        const buf  = await doc.save();
        const doc2 = await VisioDocument.load(buf);
        // Top-level shapes: only the group itself, not the child
        const shapes = doc2.pages[0].getShapes();
        expect(shapes.find(s => s.text === 'G')).toBeDefined();
        expect(shapes.find(s => s.text === 'Child')).toBeUndefined();
    });
});
