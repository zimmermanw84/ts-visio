import { describe, it, expect } from 'vitest';
import { VisioDocument } from '../src/VisioDocument';

describe('Layers', () => {
    it('should add layers with sequential indices', async () => {
        const doc = await VisioDocument.create();
        const page = doc.pages[0];

        const l1 = await page.addLayer('Background');
        const l2 = await page.addLayer('Wireframe');
        const l3 = await page.addLayer('Comments');

        expect(l1.index).toBe(0);
        expect(l2.index).toBe(1);
        expect(l3.index).toBe(2);

        expect(l1.name).toBe('Background');
        expect(l2.name).toBe('Wireframe');
        expect(l3.name).toBe('Comments');

        // Verify XML structure via Modifier
        const ShapeModifierStr = (await import('../src/ShapeModifier')).ShapeModifier;
        const testMod = new ShapeModifierStr((doc as any).pkg);
        const parsed = (testMod as any).getParsed(page.id);

        const sections = parsed.PageContents.PageSheet.Section ? (Array.isArray(parsed.PageContents.PageSheet.Section) ? parsed.PageContents.PageSheet.Section : [parsed.PageContents.PageSheet.Section]) : [];
        const layerSec = sections.find((s: any) => s['@_N'] === 'Layer');

        expect(layerSec).toBeDefined();

        const rows = Array.isArray(layerSec.Row) ? layerSec.Row : [layerSec.Row];
        expect(rows).toHaveLength(3);

        const row0 = rows[0];
        expect(row0['@_IX']).toBe('0');

        const getCell = (r: any, n: string) => {
            const cells = Array.isArray(r.Cell) ? r.Cell : [r.Cell];
            return cells.find((c: any) => c['@_N'] === n)['@_V'];
        }

        expect(getCell(row0, 'Name')).toBe('Background');
        expect(getCell(row0, 'Visible')).toBe('1');
    });

    it('should respect custom layer options', async () => {
        const doc = await VisioDocument.create();
        const page = doc.pages[0];

        const hiddenLayer = await page.addLayer('Hidden', { visible: false });
        const lockedLayer = await page.addLayer('Locked', { lock: true });

        // Verify XML
        const ShapeModifierStr = (await import('../src/ShapeModifier')).ShapeModifier;
        const testMod = new ShapeModifierStr((doc as any).pkg);
        const parsed = (testMod as any).getParsed(page.id);
        const sections = parsed.PageContents.PageSheet.Section ? (Array.isArray(parsed.PageContents.PageSheet.Section) ? parsed.PageContents.PageSheet.Section : [parsed.PageContents.PageSheet.Section]) : [];
        const layerSec = sections.find((s: any) => s['@_N'] === 'Layer');

        const rows = Array.isArray(layerSec.Row) ? layerSec.Row : [layerSec.Row];

        // Hidden Layer (IX=0)
        const getCell = (r: any, n: string) => {
            const cells = Array.isArray(r.Cell) ? r.Cell : [r.Cell];
            return cells.find((c: any) => c['@_N'] === n)['@_V'];
        }

        expect(getCell(rows[0], 'Visible')).toBe('0');

        // Locked Layer (IX=1)
        expect(getCell(rows[1], 'Lock')).toBe('1');
    });
});
