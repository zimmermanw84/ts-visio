import { describe, it, expect } from 'vitest';
import { VisioDocument } from '../src/VisioDocument';
import type { LengthUnit } from '../src/types/VisioTypes';

async function freshPage() {
    const doc  = await VisioDocument.create();
    const page = doc.pages[0];
    return { doc, page };
}

// ── getDrawingScale() — default state ─────────────────────────────────────────

describe('page.getDrawingScale() — default state', () => {
    it('returns null on a freshly created page (no custom scale)', async () => {
        const { page } = await freshPage();
        expect(page.getDrawingScale()).toBeNull();
    });
});

// ── setDrawingScale() — basic ──────────────────────────────────────────────────

describe('page.setDrawingScale() — basic', () => {
    it('returns the page for chaining', async () => {
        const { page } = await freshPage();
        expect(page.setDrawingScale(1, 'in', 10, 'ft')).toBe(page);
    });

    it('getDrawingScale() reflects the values just set', async () => {
        const { page } = await freshPage();
        page.setDrawingScale(1, 'in', 10, 'ft');
        const s = page.getDrawingScale()!;
        expect(s.pageScale).toBe(1);
        expect(s.pageUnit).toBe('in');
        expect(s.drawingScale).toBe(10);
        expect(s.drawingUnit).toBe('ft');
    });

    it('supports metric units (1:100)', async () => {
        const { page } = await freshPage();
        page.setDrawingScale(1, 'cm', 100, 'cm');
        const s = page.getDrawingScale()!;
        expect(s.pageScale).toBe(1);
        expect(s.pageUnit).toBe('cm');
        expect(s.drawingScale).toBe(100);
        expect(s.drawingUnit).toBe('cm');
    });

    it('supports mixed units (1 in = 1 m)', async () => {
        const { page } = await freshPage();
        page.setDrawingScale(1, 'in', 1, 'm');
        const s = page.getDrawingScale()!;
        expect(s.pageUnit).toBe('in');
        expect(s.drawingUnit).toBe('m');
    });

    it('supports all imperial LengthUnit values', async () => {
        const { page } = await freshPage();
        const units: LengthUnit[] = ['in', 'ft', 'yd', 'mi'];
        for (const u of units) {
            page.setDrawingScale(1, u, 1, u);
            expect(page.getDrawingScale()!.pageUnit).toBe(u);
        }
    });

    it('supports all metric LengthUnit values', async () => {
        const { page } = await freshPage();
        const units: LengthUnit[] = ['mm', 'cm', 'm', 'km'];
        for (const u of units) {
            page.setDrawingScale(1, u, 1, u);
            expect(page.getDrawingScale()!.pageUnit).toBe(u);
        }
    });

    it('overwriting an existing scale updates it correctly', async () => {
        const { page } = await freshPage();
        page.setDrawingScale(1, 'in', 10, 'ft');
        page.setDrawingScale(2, 'cm', 50, 'm');
        const s = page.getDrawingScale()!;
        expect(s.pageScale).toBe(2);
        expect(s.pageUnit).toBe('cm');
        expect(s.drawingScale).toBe(50);
        expect(s.drawingUnit).toBe('m');
    });

    it('throws when pageScale is zero or negative', async () => {
        const { page } = await freshPage();
        expect(() => page.setDrawingScale(0, 'in', 10, 'ft')).toThrow();
        expect(() => page.setDrawingScale(-1, 'in', 10, 'ft')).toThrow();
    });

    it('throws when drawingScale is zero or negative', async () => {
        const { page } = await freshPage();
        expect(() => page.setDrawingScale(1, 'in', 0, 'ft')).toThrow();
        expect(() => page.setDrawingScale(1, 'in', -5, 'ft')).toThrow();
    });
});

// ── clearDrawingScale() ────────────────────────────────────────────────────────

describe('page.clearDrawingScale()', () => {
    it('returns the page for chaining', async () => {
        const { page } = await freshPage();
        expect(page.clearDrawingScale()).toBe(page);
    });

    it('getDrawingScale() returns null after clear', async () => {
        const { page } = await freshPage();
        page.setDrawingScale(1, 'in', 10, 'ft');
        page.clearDrawingScale();
        expect(page.getDrawingScale()).toBeNull();
    });

    it('clearing on a page with no scale is a no-op (still null)', async () => {
        const { page } = await freshPage();
        page.clearDrawingScale();
        expect(page.getDrawingScale()).toBeNull();
    });
});

// ── page size is unaffected by drawing scale ───────────────────────────────────

describe('drawing scale does not affect page canvas size', () => {
    it('pageWidth and pageHeight are unchanged after setDrawingScale', async () => {
        const { page } = await freshPage();
        const w = page.pageWidth;
        const h = page.pageHeight;
        page.setDrawingScale(1, 'in', 100, 'ft');
        expect(page.pageWidth).toBe(w);
        expect(page.pageHeight).toBe(h);
    });
});

// ── save / reload round-trips ─────────────────────────────────────────────────

describe('drawing scale — save/reload round-trip', () => {
    it('scale survives save and reload', async () => {
        const { doc, page } = await freshPage();
        page.setDrawingScale(1, 'in', 10, 'ft');

        const buf  = await doc.save();
        const doc2 = await VisioDocument.load(buf);
        const s = doc2.pages[0].getDrawingScale()!;
        expect(s.pageScale).toBe(1);
        expect(s.pageUnit).toBe('in');
        expect(s.drawingScale).toBe(10);
        expect(s.drawingUnit).toBe('ft');
    });

    it('cleared scale (null) survives save and reload', async () => {
        const { doc, page } = await freshPage();
        page.setDrawingScale(1, 'cm', 100, 'cm');
        page.clearDrawingScale();

        const buf  = await doc.save();
        const doc2 = await VisioDocument.load(buf);
        expect(doc2.pages[0].getDrawingScale()).toBeNull();
    });

    it('scale on a second page survives save and reload independently', async () => {
        const doc  = await VisioDocument.create();
        const pg1  = doc.pages[0];
        const pg2  = await doc.addPage('P2');
        pg1.setDrawingScale(1, 'in', 10,  'ft');
        pg2.setDrawingScale(1, 'cm', 100, 'cm');

        const buf  = await doc.save();
        const doc2 = await VisioDocument.load(buf);
        const s1   = doc2.pages[0].getDrawingScale()!;
        const s2   = doc2.pages[1].getDrawingScale()!;
        expect(s1.drawingUnit).toBe('ft');
        expect(s2.drawingUnit).toBe('cm');
    });
});
