import { describe, it, expect } from 'vitest';
import { VisioDocument } from '../src/VisioDocument';

describe('Hyperlinks', () => {
    it('should add a hyperlink with description', async () => {
        const doc = await VisioDocument.create();
        const page = doc.pages[0];
        const shape = await page.addShape({ text: 'Link', x: 1, y: 1, width: 1, height: 1 });

        await shape.addHyperlink('https://google.com', 'Google Search');

        // Verify XML via Modifier (internal access or save and parse)
        const ShapeModifierStr = (await import('../src/ShapeModifier')).ShapeModifier;
        const testMod = new ShapeModifierStr((doc as any).pkg);

        // Re-parse to verify internal structure
        const parsed = (testMod as any).getParsed(page.id);
        const shapes = (testMod as any).getAllShapes(parsed);
        const s = shapes.find((x: any) => x['@_ID'] == shape.id);

        const linkSec = s.Section.find((x: any) => x['@_N'] === 'Hyperlink');
        expect(linkSec).toBeDefined();

        const rows = Array.isArray(linkSec.Row) ? linkSec.Row : [linkSec.Row];
        const row = rows[0];
        expect(row['@_N']).toBe('Hyperlink.Row_1');

        const cells = Array.isArray(row.Cell) ? row.Cell : [row.Cell];
        const addr = cells.find((c: any) => c['@_N'] === 'Address');
        expect(addr['@_V']).toBe('https://google.com');

        const desc = cells.find((c: any) => c['@_N'] === 'Description');
        expect(desc['@_V']).toBe('Google Search');
    });

    it('should escape special characters in URLs', async () => {
        const doc = await VisioDocument.create();
        const page = doc.pages[0];
        const shape = await page.addShape({ text: 'Complex Link', x: 1, y: 1, width: 1, height: 1 });

        const complexUrl = 'https://example.com?q=1&b=2';
        await shape.addHyperlink(complexUrl);

        const ShapeModifierStr = (await import('../src/ShapeModifier')).ShapeModifier;
        const testMod = new ShapeModifierStr((doc as any).pkg);
        const parsed = (testMod as any).getParsed(page.id);
        const shapes = (testMod as any).getAllShapes(parsed);
        const s = shapes.find((x: any) => x['@_ID'] == shape.id);

        const linkSec = s.Section.find((x: any) => x['@_N'] === 'Hyperlink');
        const rows = Array.isArray(linkSec.Row) ? linkSec.Row : [linkSec.Row];
        const row = rows[0];
        const cells = Array.isArray(row.Cell) ? row.Cell : [row.Cell];

        const addr = cells.find((c: any) => c['@_N'] === 'Address');

        expect(addr['@_V']).toContain('amp;');
    });

    it('should support multiple links', async () => {
        const doc = await VisioDocument.create();
        const page = doc.pages[0];
        const shape = await page.addShape({ text: 'Multi Link', x: 1, y: 1, width: 1, height: 1 });

        await shape.addHyperlink('https://link1.com');
        await shape.addHyperlink('https://link2.com');

        const ShapeModifierStr = (await import('../src/ShapeModifier')).ShapeModifier;
        const testMod = new ShapeModifierStr((doc as any).pkg);
        const parsed = (testMod as any).getParsed(page.id);
        const shapes = (testMod as any).getAllShapes(parsed);
        const s = shapes.find((x: any) => x['@_ID'] == shape.id);

        const linkSec = s.Section.find((x: any) => x['@_N'] === 'Hyperlink');
        const rows = Array.isArray(linkSec.Row) ? linkSec.Row : [linkSec.Row];
        expect(rows).toHaveLength(2);
        expect(rows[0]['@_N']).toBe('Hyperlink.Row_1');
        expect(rows[1]['@_N']).toBe('Hyperlink.Row_2');
        expect(rows).toHaveLength(2);
        expect(rows[0]['@_N']).toBe('Hyperlink.Row_1');
        expect(rows[1]['@_N']).toBe('Hyperlink.Row_2');
    });

    it('should support internal page linking', async () => {
        const doc = await VisioDocument.create();
        const page1 = doc.pages[0];
        const page2 = await doc.addPage('Details Page');

        const shape = await page1.addShape({ text: 'Go to Details', x: 1, y: 1, width: 1, height: 1 });
        await shape.linkToPage(page2, 'Drill Down');

        const ShapeModifierStr = (await import('../src/ShapeModifier')).ShapeModifier;
        const testMod = new ShapeModifierStr((doc as any).pkg);
        const parsed = (testMod as any).getParsed(page1.id);
        const shapes = (testMod as any).getAllShapes(parsed);
        const s = shapes.find((x: any) => x['@_ID'] == shape.id);

        const linkSec = s.Section.find((x: any) => x['@_N'] === 'Hyperlink');
        const rows = Array.isArray(linkSec.Row) ? linkSec.Row : [linkSec.Row];
        const row = rows[0];
        const cells = Array.isArray(row.Cell) ? row.Cell : [row.Cell];

        // Address should be empty
        const addr = cells.find((c: any) => c['@_N'] === 'Address');
        // Address might be missing or empty? My code pushes it only if details.address is provided.
        // In linkToPage I pass address: ''. So it pushes Address=''.
        if (addr) expect(addr['@_V']).toBe('');

        // SubAddress should be Page Name
        const subAddr = cells.find((c: any) => c['@_N'] === 'SubAddress');
        expect(subAddr['@_V']).toBe('Details Page');
    });

    it('should support fluent API chaining (toUrl/toPage)', async () => {
        const doc = await VisioDocument.create();
        const page = doc.pages[0];
        const page2 = await doc.addPage('Other Page');

        const shape = await page.addShape({ text: 'Chain', x: 1, y: 1, width: 1, height: 1 });

        // Chain calls
        await (await shape.toUrl('https://example.com'))
            .toPage(page2);

        const ShapeModifierStr = (await import('../src/ShapeModifier')).ShapeModifier;
        const testMod = new ShapeModifierStr((doc as any).pkg);
        const parsed = (testMod as any).getParsed(page.id);
        const shapes = (testMod as any).getAllShapes(parsed);
        const s = shapes.find((x: any) => x['@_ID'] == shape.id);

        const linkSec = s.Section.find((x: any) => x['@_N'] === 'Hyperlink');
        const rows = Array.isArray(linkSec.Row) ? linkSec.Row : [linkSec.Row];

        expect(rows).toHaveLength(2);
        // Row 1: External
        const cells1 = Array.isArray(rows[0].Cell) ? rows[0].Cell : [rows[0].Cell];
        expect(cells1.find((c: any) => c['@_N'] === 'Address')['@_V']).toBe('https://example.com');

        // Row 2: Internal
        const cells2 = Array.isArray(rows[1].Cell) ? rows[1].Cell : [rows[1].Cell];
        expect(cells2.find((c: any) => c['@_N'] === 'SubAddress')['@_V']).toBe('Other Page');
    });
});
