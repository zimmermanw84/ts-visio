import { PageXmlCache } from './PageXmlCache';
import { SECTION_NAMES } from './VisioConstants';

/**
 * Manages page layers: creation, assignment, property updates, and deletion.
 */
export class LayerEditor {
    constructor(private cache: PageXmlCache) {}

    async addLayer(
        pageId: string,
        name: string,
        options: { visible?: boolean; lock?: boolean; print?: boolean } = {},
    ): Promise<{ name: string; index: number }> {
        const parsed = this.cache.getParsed(pageId);
        this.cache.ensurePageSheet(parsed);
        const pageSheet = parsed.PageContents.PageSheet;

        if (!pageSheet.Section) pageSheet.Section = [];
        if (!Array.isArray(pageSheet.Section)) pageSheet.Section = [pageSheet.Section];

        let layerSection = pageSheet.Section.find((s: any) => s['@_N'] === SECTION_NAMES.Layer);
        if (!layerSection) {
            layerSection = { '@_N': SECTION_NAMES.Layer, Row: [] };
            pageSheet.Section.push(layerSection);
        }

        if (!layerSection.Row) layerSection.Row = [];
        if (!Array.isArray(layerSection.Row)) layerSection.Row = [layerSection.Row];

        let maxIx = -1;
        for (const row of layerSection.Row) {
            const ix = parseInt(row['@_IX']);
            if (!isNaN(ix) && ix > maxIx) maxIx = ix;
        }
        const newIndex = maxIx + 1;

        layerSection.Row.push({
            '@_IX': newIndex.toString(),
            Cell: [
                { '@_N': 'Name',    '@_V': name },
                { '@_N': 'Visible', '@_V': (options.visible ?? true)  ? '1' : '0' },
                { '@_N': 'Lock',    '@_V': (options.lock    ?? false) ? '1' : '0' },
                { '@_N': 'Print',   '@_V': (options.print   ?? true)  ? '1' : '0' },
            ],
        });

        this.cache.saveParsed(pageId, parsed);
        return { name, index: newIndex };
    }

    async assignLayer(pageId: string, shapeId: string, layerIndex: number): Promise<void> {
        const parsed = this.cache.getParsed(pageId);
        const shape = this.cache.getShapeMap(parsed).get(shapeId);
        if (!shape) throw new Error(`Shape ${shapeId} not found`);

        if (!shape.Section) shape.Section = [];
        if (!Array.isArray(shape.Section)) shape.Section = [shape.Section];

        let memSection = shape.Section.find((s: any) => s['@_N'] === SECTION_NAMES.LayerMem);
        if (!memSection) {
            memSection = { '@_N': SECTION_NAMES.LayerMem, Row: [] };
            shape.Section.push(memSection);
        }

        if (!memSection.Row) memSection.Row = [];
        if (!Array.isArray(memSection.Row)) memSection.Row = [memSection.Row];

        // LayerMem contains a single row with a semicolon-separated LayerMember cell.
        if (memSection.Row.length === 0) {
            memSection.Row.push({ Cell: [] });
        }
        const row = memSection.Row[0];

        if (!row.Cell) row.Cell = [];
        if (!Array.isArray(row.Cell)) row.Cell = [row.Cell];

        let cell = row.Cell.find((c: any) => c['@_N'] === 'LayerMember');
        if (!cell) {
            cell = { '@_N': 'LayerMember', '@_V': '' };
            row.Cell.push(cell);
        }

        const currentVal = cell['@_V'] || '';
        const indices = currentVal.split(';').filter((s: string) => s.length > 0);
        const idxStr = layerIndex.toString();

        if (!indices.includes(idxStr)) {
            indices.push(idxStr);
            cell['@_V'] = indices.join(';');
            this.cache.saveParsed(pageId, parsed);
        }
    }

    async updateLayerProperty(pageId: string, layerIndex: number, propName: string, value: string): Promise<void> {
        const parsed = this.cache.getParsed(pageId);
        this.cache.ensurePageSheet(parsed);
        const pageSheet = parsed.PageContents.PageSheet;

        if (!pageSheet.Section) return;
        const sections = Array.isArray(pageSheet.Section) ? pageSheet.Section : [pageSheet.Section];
        const layerSection = sections.find((s: any) => s['@_N'] === SECTION_NAMES.Layer);
        if (!layerSection?.Row) return;

        const rows = Array.isArray(layerSection.Row) ? layerSection.Row : [layerSection.Row];
        const row = rows.find((r: any) => r['@_IX'] === layerIndex.toString());
        if (!row) return;

        if (!row.Cell) row.Cell = [];
        if (!Array.isArray(row.Cell)) row.Cell = [row.Cell];

        let cell = row.Cell.find((c: any) => c['@_N'] === propName);
        if (!cell) {
            cell = { '@_N': propName, '@_V': value };
            row.Cell.push(cell);
        } else {
            cell['@_V'] = value;
        }

        this.cache.saveParsed(pageId, parsed);
    }

    /**
     * Return all layers defined in the page's PageSheet as plain objects.
     */
    getPageLayers(pageId: string): Array<{ name: string; index: number; visible: boolean; locked: boolean; print: boolean }> {
        const parsed = this.cache.getParsed(pageId);
        const pageSheet = parsed.PageContents?.PageSheet;
        if (!pageSheet?.Section) return [];

        const sections = Array.isArray(pageSheet.Section) ? pageSheet.Section : [pageSheet.Section];
        const layerSection = sections.find((s: any) => s['@_N'] === SECTION_NAMES.Layer);
        if (!layerSection?.Row) return [];

        const rows = Array.isArray(layerSection.Row) ? layerSection.Row : [layerSection.Row];
        return rows.map((row: any) => {
            const cells: any[] = Array.isArray(row.Cell) ? row.Cell : (row.Cell ? [row.Cell] : []);
            const getVal = (name: string) => cells.find((c: any) => c['@_N'] === name)?.['@_V'];
            return {
                name:    getVal('Name')    ?? '',
                index:   parseInt(row['@_IX'], 10),
                visible: getVal('Visible') !== '0',
                locked:  getVal('Lock')    === '1',
                print:   getVal('Print')   !== '0',
            };
        });
    }

    /**
     * Delete a layer by index, re-index remaining layers to close gaps, and
     * update all shape LayerMember cells to use the new indices.
     */
    deleteLayer(pageId: string, layerIndex: number): void {
        const parsed = this.cache.getParsed(pageId);

        // Build old-index → new-index mapping while removing the deleted layer.
        const indexRemap = new Map<string, string>();

        const pageSheet = parsed.PageContents?.PageSheet;
        if (pageSheet?.Section) {
            const sections = Array.isArray(pageSheet.Section) ? pageSheet.Section : [pageSheet.Section];
            const layerSection = sections.find((s: any) => s['@_N'] === SECTION_NAMES.Layer);
            if (layerSection?.Row) {
                const rows: any[] = Array.isArray(layerSection.Row) ? layerSection.Row : [layerSection.Row];

                // Sort by numeric IX so re-indexing is deterministic.
                const remaining = rows
                    .filter((r: any) => r['@_IX'] !== layerIndex.toString())
                    .sort((a: any, b: any) => parseInt(a['@_IX'], 10) - parseInt(b['@_IX'], 10));

                remaining.forEach((row: any, newIx: number) => {
                    const oldIx = row['@_IX'];
                    if (oldIx !== newIx.toString()) {
                        indexRemap.set(oldIx, newIx.toString());
                        row['@_IX'] = newIx.toString();
                    }
                });

                layerSection.Row = remaining;
            }
        }

        // Update every shape's LayerMember cell: remove deleted index, remap survivors.
        const deletedStr = layerIndex.toString();
        for (const [, shape] of this.cache.getShapeMap(parsed)) {
            if (!shape.Section) continue;
            const sections: any[] = Array.isArray(shape.Section) ? shape.Section : [shape.Section];
            const memSec = sections.find((s: any) => s['@_N'] === SECTION_NAMES.LayerMem);
            if (!memSec?.Row) continue;

            const rows = Array.isArray(memSec.Row) ? memSec.Row : [memSec.Row];
            const row = rows[0];
            if (!row?.Cell) continue;

            const cells: any[] = Array.isArray(row.Cell) ? row.Cell : [row.Cell];
            const memberCell = cells.find((c: any) => c['@_N'] === 'LayerMember');
            if (!memberCell?.['@_V']) continue;

            const updated = memberCell['@_V']
                .split(';')
                .filter((s: string) => s.length > 0 && s !== deletedStr)
                .map((s: string) => indexRemap.get(s) ?? s);
            memberCell['@_V'] = updated.join(';');
        }

        this.cache.saveParsed(pageId, parsed);
    }

    /**
     * Read back the layer indices a shape is assigned to.
     * Returns an empty array if the shape has no layer assignment.
     */
    getShapeLayerIndices(pageId: string, shapeId: string): number[] {
        const parsed = this.cache.getParsed(pageId);
        const shape = this.cache.getShapeMap(parsed).get(shapeId);
        if (!shape) throw new Error(`Shape ${shapeId} not found on page ${pageId}`);

        if (!shape.Section) return [];
        const sections = Array.isArray(shape.Section) ? shape.Section : [shape.Section];
        const memSection = sections.find((s: any) => s['@_N'] === SECTION_NAMES.LayerMem);
        if (!memSection?.Row) return [];

        const rows = Array.isArray(memSection.Row) ? memSection.Row : [memSection.Row];
        const row = rows[0];
        if (!row) return [];

        const cells: any[] = Array.isArray(row.Cell) ? row.Cell : (row.Cell ? [row.Cell] : []);
        const memberCell = cells.find((c: any) => c['@_N'] === 'LayerMember');
        if (!memberCell?.['@_V']) return [];

        return (memberCell['@_V'] as string)
            .split(';')
            .filter((s: string) => s.length > 0)
            .map((s: string) => parseInt(s))
            .filter((n: number) => !isNaN(n));
    }
}
