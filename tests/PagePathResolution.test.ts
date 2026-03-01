/**
 * Regression tests for BUG 3: Page ID-to-path mapping was hardcoded as
 * `page${id}.xml`, which breaks for loaded files where page IDs and filenames
 * are not in sync (e.g. after pages have been deleted and re-added in Visio).
 *
 * The correct path must always be resolved through the .rels relationship
 * chain, not inferred from the page ID.
 */
import { describe, it, expect, vi } from 'vitest';
import { Page } from '../src/Page';
import { ShapeModifier } from '../src/ShapeModifier';
import { VisioPackage } from '../src/VisioPackage';
import { VisioDocument } from '../src/VisioDocument';

const BLANK_PAGE_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<PageContents xmlns="http://schemas.microsoft.com/office/visio/2012/main"
              xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
              xml:space="preserve">
  <Shapes/>
  <Connects/>
</PageContents>`;

describe('Page path resolution (BUG 3)', () => {
    it('ShapeModifier.registerPage overrides the ID-derived fallback path', () => {
        const mockPkg = {
            getFileText: vi.fn().mockReturnValue(BLANK_PAGE_XML),
            updateFile: vi.fn(),
        } as unknown as VisioPackage;

        const modifier = new ShapeModifier(mockPkg);

        // Page ID 5 exists in a file named page2.xml (common after deletions)
        modifier.registerPage('5', 'visio/pages/page2.xml');

        // Trigger any read operation — addShape internally calls getPagePath('5')
        modifier.addShape('5', { text: 'X', x: 1, y: 1, width: 1, height: 1 });

        const writtenPath = (mockPkg.updateFile as ReturnType<typeof vi.fn>).mock.calls[0][0];
        expect(writtenPath).toBe('visio/pages/page2.xml');
        expect(writtenPath).not.toBe('visio/pages/page5.xml');
    });

    it('Page uses xmlPath for getShapes() rather than computing from ID', () => {
        const mockPkg = {
            getFileText: vi.fn().mockReturnValue(BLANK_PAGE_XML),
            updateFile: vi.fn(),
        } as unknown as VisioPackage;

        // Page ID 7, but the actual file is page3.xml
        const page = new Page(
            { ID: '7', Name: 'Test', xmlPath: 'visio/pages/page3.xml', Shapes: [], Connects: [] },
            mockPkg
        );

        page.getShapes();

        const readPath = (mockPkg.getFileText as ReturnType<typeof vi.fn>).mock.calls[0][0];
        expect(readPath).toBe('visio/pages/page3.xml');
        expect(readPath).not.toBe('visio/pages/page7.xml');
    });

    it('Page falls back to ID-derived path when xmlPath is absent', () => {
        const mockPkg = {
            getFileText: vi.fn().mockReturnValue(BLANK_PAGE_XML),
            updateFile: vi.fn(),
        } as unknown as VisioPackage;

        // No xmlPath provided — should derive from ID (correct for newly-created pages)
        const page = new Page(
            { ID: '2', Name: 'Test', Shapes: [], Connects: [] },
            mockPkg
        );

        page.getShapes();

        const readPath = (mockPkg.getFileText as ReturnType<typeof vi.fn>).mock.calls[0][0];
        expect(readPath).toBe('visio/pages/page2.xml');
    });

    it('VisioDocument.pages passes resolved xmlPath from PageEntry to Page', async () => {
        // Create a minimal document, save it, then reload. The reloaded pages
        // must use the rels-resolved path, not a guess from the page ID.
        const doc = await VisioDocument.create();
        const page = doc.pages[0];
        await page.addShape({ text: 'Hello', x: 1, y: 1, width: 1, height: 1 });

        const saved = await doc.save();
        const reloaded = await VisioDocument.load(saved);

        // Verify shapes can be read back on the reloaded page using resolved path
        const reloadedPage = reloaded.pages[0];
        const shapes = reloadedPage.getShapes();
        expect(shapes).toHaveLength(1);
        expect(shapes[0].text).toBe('Hello');
    });
});
