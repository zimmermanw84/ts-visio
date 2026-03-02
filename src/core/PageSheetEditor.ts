import { PageXmlCache } from './PageXmlCache';
import { LENGTH_UNIT_TO_VISIO, VISIO_TO_LENGTH_UNIT } from './VisioConstants';
import type { DrawingScaleInfo, LengthUnit } from '../types/VisioTypes';

/**
 * Reads and writes page-level sheet properties: canvas size and drawing scale.
 */
export class PageSheetEditor {
    constructor(private cache: PageXmlCache) {}

    /**
     * Set the page canvas size. Writes PageWidth / PageHeight into the PageSheet
     * and sets DrawingSizeType=0 (Custom) so Visio does not override the values.
     */
    setPageSize(pageId: string, width: number, height: number): void {
        if (width <= 0 || height <= 0) throw new Error('Page dimensions must be positive');
        const parsed = this.cache.getParsed(pageId);
        this.cache.ensurePageSheet(parsed);
        const ps = parsed.PageContents.PageSheet;

        const upsert = (name: string, value: string) => {
            const existing = ps.Cell.find((c: any) => c['@_N'] === name);
            if (existing) existing['@_V'] = value;
            else ps.Cell.push({ '@_N': name, '@_V': value });
        };

        upsert('PageWidth',       width.toString());
        upsert('PageHeight',      height.toString());
        upsert('DrawingSizeType', '0');
        upsert('PageDrawSizeType', '0');

        this.cache.saveParsed(pageId, parsed);
    }

    /**
     * Read the current page canvas dimensions.
     * Returns 8.5 × 11 (US Letter) if no PageSheet cells are present.
     */
    getPageDimensions(pageId: string): { width: number; height: number } {
        const parsed = this.cache.getParsed(pageId);
        const ps = parsed.PageContents?.PageSheet;
        if (!ps?.Cell) return { width: 8.5, height: 11 };
        const cells: any[] = Array.isArray(ps.Cell) ? ps.Cell : [ps.Cell];
        const getVal = (name: string, def: number): number => {
            const c = cells.find((cell: any) => cell['@_N'] === name);
            return c ? parseFloat(c['@_V']) : def;
        };
        return { width: getVal('PageWidth', 8.5), height: getVal('PageHeight', 11) };
    }

    /**
     * Read the drawing scale from the PageSheet.
     * Returns `null` when no custom scale is set (i.e. 1:1 / no-scale default).
     */
    getDrawingScale(pageId: string): DrawingScaleInfo | null {
        const parsed = this.cache.getParsed(pageId);
        const ps = parsed.PageContents?.PageSheet;
        if (!ps?.Cell) return null;
        const cells: any[] = Array.isArray(ps.Cell) ? ps.Cell : [ps.Cell];
        const getCell = (name: string) => cells.find((c: any) => c['@_N'] === name);

        const psCell = getCell('PageScale');
        const dsCell = getCell('DrawingScale');
        if (!psCell && !dsCell) return null;

        const psUnit = psCell?.['@_Unit'] ?? 'MSG';
        const dsUnit = dsCell?.['@_Unit'] ?? 'MSG';
        // "MSG" is Visio's sentinel for "no real unit" — treat as 1:1
        if (psUnit === 'MSG' && dsUnit === 'MSG') return null;

        const toUserUnit = (v: string): LengthUnit =>
            (VISIO_TO_LENGTH_UNIT[v] as LengthUnit | undefined) ?? 'in';

        return {
            pageScale:    parseFloat(psCell?.['@_V'] ?? '1'),
            pageUnit:     toUserUnit(psUnit),
            drawingScale: parseFloat(dsCell?.['@_V'] ?? '1'),
            drawingUnit:  toUserUnit(dsUnit),
        };
    }

    /**
     * Set a custom drawing scale on the PageSheet.
     * @param pageScale    Measurement on paper (e.g. 1).
     * @param pageUnit     Unit for the paper measurement (e.g. `'in'`).
     * @param drawingScale Real-world measurement (e.g. 10).
     * @param drawingUnit  Unit for the real-world measurement (e.g. `'ft'`).
     */
    setDrawingScale(
        pageId: string,
        pageScale: number, pageUnit: LengthUnit,
        drawingScale: number, drawingUnit: LengthUnit,
    ): void {
        if (pageScale <= 0 || drawingScale <= 0) {
            throw new Error('Drawing scale values must be positive');
        }
        const parsed = this.cache.getParsed(pageId);
        this.cache.ensurePageSheet(parsed);
        const ps = parsed.PageContents.PageSheet;

        const upsertWithUnit = (name: string, value: string, unit: string) => {
            const existing = ps.Cell.find((c: any) => c['@_N'] === name);
            if (existing) {
                existing['@_V'] = value;
                existing['@_Unit'] = unit;
            } else {
                ps.Cell.push({ '@_N': name, '@_V': value, '@_Unit': unit });
            }
        };

        upsertWithUnit('PageScale',    pageScale.toString(),    LENGTH_UNIT_TO_VISIO[pageUnit]);
        upsertWithUnit('DrawingScale', drawingScale.toString(), LENGTH_UNIT_TO_VISIO[drawingUnit]);

        this.cache.saveParsed(pageId, parsed);
    }

    /**
     * Reset the drawing scale to 1:1 (no custom scale).
     * Restores `PageScale` and `DrawingScale` to `V="1" Unit="MSG"`.
     */
    clearDrawingScale(pageId: string): void {
        const parsed = this.cache.getParsed(pageId);
        this.cache.ensurePageSheet(parsed);
        const ps = parsed.PageContents.PageSheet;

        const reset = (name: string) => {
            const existing = ps.Cell.find((c: any) => c['@_N'] === name);
            if (existing) {
                existing['@_V'] = '1';
                existing['@_Unit'] = 'MSG';
            } else {
                ps.Cell.push({ '@_N': name, '@_V': '1', '@_Unit': 'MSG' });
            }
        };

        reset('PageScale');
        reset('DrawingScale');
        this.cache.saveParsed(pageId, parsed);
    }
}
