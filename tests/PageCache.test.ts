import { describe, it, expect } from 'vitest';
import { VisioDocument } from '../src/VisioDocument';

describe('PageManager Caching', () => {
    it('should update pages list immediately after adding a page', async () => {
        const doc = await VisioDocument.create();

        // Initial state (usually 1 default page)
        const initialCount = doc.pages.length;
        expect(initialCount).toBeGreaterThan(0);

        // First access (populates cache)
        const pages1 = doc.pages;
        expect(pages1).toHaveLength(initialCount);

        // Add a page
        await doc.addPage('Cache Test Page');

        // Second access (should be invalidated and refreshed)
        const pages2 = doc.pages;
        expect(pages2).toHaveLength(initialCount + 1);
        expect(pages2.find(p => p.name === 'Cache Test Page')).toBeDefined();

        // Third access (should use new cache)
        const pages3 = doc.pages;
        expect(pages3).toHaveLength(initialCount + 1);

        // Ensure object identity might be different (new objects created from cache data)
        // actually, VisioDocument.pages maps internal entries to new Page objects every time.
        // So they won't be strictly equal objects, but content should match.
    });
});
