import { describe, it, expect } from 'vitest';
import { VisioDocument } from '../src/VisioDocument';
import { PageSizes, PageSizeName } from '../src/types/VisioTypes';

// ---------------------------------------------------------------------------
// PageSizes constant
// ---------------------------------------------------------------------------

describe('PageSizes constant', () => {
    it('exports Letter dimensions', () => {
        expect(PageSizes.Letter.width).toBe(8.5);
        expect(PageSizes.Letter.height).toBe(11);
    });

    it('exports Legal dimensions', () => {
        expect(PageSizes.Legal.width).toBe(8.5);
        expect(PageSizes.Legal.height).toBe(14);
    });

    it('exports Tabloid dimensions', () => {
        expect(PageSizes.Tabloid.width).toBe(11);
        expect(PageSizes.Tabloid.height).toBe(17);
    });

    it('exports A3 dimensions', () => {
        expect(PageSizes.A3.width).toBeCloseTo(11.693, 3);
        expect(PageSizes.A3.height).toBeCloseTo(16.535, 3);
    });

    it('exports A4 dimensions', () => {
        expect(PageSizes.A4.width).toBeCloseTo(8.268, 3);
        expect(PageSizes.A4.height).toBeCloseTo(11.693, 3);
    });

    it('exports A5 dimensions', () => {
        expect(PageSizes.A5.width).toBeCloseTo(5.827, 3);
        expect(PageSizes.A5.height).toBeCloseTo(8.268, 3);
    });
});

// ---------------------------------------------------------------------------
// page.pageWidth / page.pageHeight defaults
// ---------------------------------------------------------------------------

describe('Page dimension defaults', () => {
    it('returns 8.5 × 11 when no page size has been set (US Letter default)', async () => {
        const doc = await VisioDocument.create();
        const page = doc.pages[0];

        expect(page.pageWidth).toBe(8.5);
        expect(page.pageHeight).toBe(11);
    });

    it('orientation is portrait for default dimensions', async () => {
        const doc = await VisioDocument.create();
        const page = doc.pages[0];

        expect(page.orientation).toBe('portrait');
    });
});

// ---------------------------------------------------------------------------
// page.setSize()
// ---------------------------------------------------------------------------

describe('Page.setSize()', () => {
    it('updates pageWidth and pageHeight', async () => {
        const doc = await VisioDocument.create();
        const page = doc.pages[0];

        page.setSize(11, 8.5);

        expect(page.pageWidth).toBe(11);
        expect(page.pageHeight).toBe(8.5);
    });

    it('returns this for chaining', async () => {
        const doc = await VisioDocument.create();
        const page = doc.pages[0];

        const result = page.setSize(11, 8.5);

        expect(result).toBe(page);
    });

    it('sets orientation to landscape when width > height', async () => {
        const doc = await VisioDocument.create();
        const page = doc.pages[0];

        page.setSize(14, 8.5);

        expect(page.orientation).toBe('landscape');
    });

    it('sets orientation to portrait when height > width', async () => {
        const doc = await VisioDocument.create();
        const page = doc.pages[0];

        page.setSize(8.5, 14);

        expect(page.orientation).toBe('portrait');
    });

    it('can set size multiple times, always reflecting the latest values', async () => {
        const doc = await VisioDocument.create();
        const page = doc.pages[0];

        page.setSize(11, 8.5);
        page.setSize(8.268, 11.693);

        expect(page.pageWidth).toBeCloseTo(8.268, 3);
        expect(page.pageHeight).toBeCloseTo(11.693, 3);
    });

    it('throws for non-positive width', async () => {
        const doc = await VisioDocument.create();
        const page = doc.pages[0];

        expect(() => page.setSize(0, 11)).toThrow();
        expect(() => page.setSize(-1, 11)).toThrow();
    });

    it('throws for non-positive height', async () => {
        const doc = await VisioDocument.create();
        const page = doc.pages[0];

        expect(() => page.setSize(8.5, 0)).toThrow();
        expect(() => page.setSize(8.5, -5)).toThrow();
    });
});

// ---------------------------------------------------------------------------
// page.setNamedSize()
// ---------------------------------------------------------------------------

describe('Page.setNamedSize()', () => {
    const sizes: PageSizeName[] = ['Letter', 'Legal', 'Tabloid', 'A3', 'A4', 'A5'];

    for (const name of sizes) {
        it(`sets ${name} portrait correctly`, async () => {
            const doc = await VisioDocument.create();
            const page = doc.pages[0];

            page.setNamedSize(name, 'portrait');

            expect(page.pageWidth).toBeCloseTo(PageSizes[name].width, 3);
            expect(page.pageHeight).toBeCloseTo(PageSizes[name].height, 3);
            expect(page.orientation).toBe('portrait');
        });

        it(`sets ${name} landscape correctly`, async () => {
            const doc = await VisioDocument.create();
            const page = doc.pages[0];

            page.setNamedSize(name, 'landscape');

            expect(page.pageWidth).toBeCloseTo(PageSizes[name].height, 3);
            expect(page.pageHeight).toBeCloseTo(PageSizes[name].width, 3);
            expect(page.orientation).toBe('landscape');
        });
    }

    it('defaults to portrait when orientation is omitted', async () => {
        const doc = await VisioDocument.create();
        const page = doc.pages[0];

        page.setNamedSize('Letter');

        expect(page.pageWidth).toBe(8.5);
        expect(page.pageHeight).toBe(11);
        expect(page.orientation).toBe('portrait');
    });

    it('returns this for chaining', async () => {
        const doc = await VisioDocument.create();
        const page = doc.pages[0];

        const result = page.setNamedSize('A4');

        expect(result).toBe(page);
    });
});

// ---------------------------------------------------------------------------
// page.setOrientation()
// ---------------------------------------------------------------------------

describe('Page.setOrientation()', () => {
    it('swaps dimensions when switching portrait → landscape', async () => {
        const doc = await VisioDocument.create();
        const page = doc.pages[0];

        // Start at Letter portrait (8.5 × 11)
        page.setNamedSize('Letter', 'portrait');
        page.setOrientation('landscape');

        expect(page.pageWidth).toBeCloseTo(11, 5);
        expect(page.pageHeight).toBeCloseTo(8.5, 5);
        expect(page.orientation).toBe('landscape');
    });

    it('swaps dimensions when switching landscape → portrait', async () => {
        const doc = await VisioDocument.create();
        const page = doc.pages[0];

        page.setNamedSize('Letter', 'landscape');
        page.setOrientation('portrait');

        expect(page.pageWidth).toBeCloseTo(8.5, 5);
        expect(page.pageHeight).toBeCloseTo(11, 5);
        expect(page.orientation).toBe('portrait');
    });

    it('is a no-op when orientation already matches', async () => {
        const doc = await VisioDocument.create();
        const page = doc.pages[0];

        page.setNamedSize('Letter', 'landscape');
        page.setOrientation('landscape');

        expect(page.pageWidth).toBeCloseTo(11, 5);
        expect(page.pageHeight).toBeCloseTo(8.5, 5);
    });

    it('returns this for chaining', async () => {
        const doc = await VisioDocument.create();
        const page = doc.pages[0];

        const result = page.setOrientation('landscape');

        expect(result).toBe(page);
    });
});

// ---------------------------------------------------------------------------
// Persistence: changes survive serialisation round-trip
// ---------------------------------------------------------------------------

describe('Page size persistence', () => {
    it('dimensions survive save/reload cycle', async () => {
        const doc = await VisioDocument.create();
        const page = doc.pages[0];
        page.setNamedSize('A4', 'landscape');

        const buffer = await doc.save();

        const reloaded = await VisioDocument.load(buffer);
        const reloadedPage = reloaded.pages[0];

        expect(reloadedPage.pageWidth).toBeCloseTo(PageSizes.A4.height, 3);
        expect(reloadedPage.pageHeight).toBeCloseTo(PageSizes.A4.width, 3);
        expect(reloadedPage.orientation).toBe('landscape');
    });

    it('newly added page defaults to 8.5 × 11 portrait', async () => {
        const doc = await VisioDocument.create();
        await doc.addPage('Page-2');

        const page2 = doc.pages[1];

        expect(page2.pageWidth).toBe(8.5);
        expect(page2.pageHeight).toBe(11);
        expect(page2.orientation).toBe('portrait');
    });

    it('second page size is independent of first page', async () => {
        const doc = await VisioDocument.create();
        const page1 = doc.pages[0];
        page1.setSize(11, 8.5);

        await doc.addPage('Page-2');
        const page2 = doc.pages[1];

        expect(page2.pageWidth).toBe(8.5);
        expect(page2.pageHeight).toBe(11);
    });
});

// ---------------------------------------------------------------------------
// Export regression
// ---------------------------------------------------------------------------

describe('Public exports', () => {
    it('PageSizes is exported from the package root', async () => {
        const root = await import('../src/index');
        expect(root.PageSizes).toBeDefined();
        expect(root.PageSizes.Letter).toBeDefined();
    });
});
