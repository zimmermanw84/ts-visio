
import { VisioPackage } from '../VisioPackage';
import { StyleProps, StyleRecord } from '../types/VisioTypes';
import { hexToRgb } from '../utils/StyleHelpers';
import { createXmlParser, createXmlBuilder, buildXml } from '../utils/XmlHelper';

const HORZ_ALIGN: Record<string, string> = {
    left: '0', center: '1', right: '2', justify: '3',
};
const VERT_ALIGN: Record<string, string> = {
    top: '0', middle: '1', bottom: '2',
};

/**
 * Manages document-level stylesheets stored in `visio/document.xml`.
 *
 * Stylesheets define reusable sets of line, fill, and text properties that
 * shapes inherit via the `LineStyle`, `FillStyle`, and `TextStyle` attributes.
 *
 * The minimal document template ships with:
 *   - ID 0  "No Style"  — the inheritance root (all shapes fall back here)
 *   - ID 1  "Normal"    — empty style inheriting everything from ID 0
 *
 * User-created styles receive IDs starting at 2.
 */
export class StyleSheetManager {
    private parser = createXmlParser();
    private builder = createXmlBuilder();

    constructor(private pkg: VisioPackage) {}

    // ── Public API ────────────────────────────────────────────────────────────

    /**
     * Create a new named document-level stylesheet and return its record.
     * The returned `id` can be passed to `addShape({ styleId })` or `shape.applyStyle()`.
     */
    createStyle(name: string, props: StyleProps = {}): StyleRecord {
        const parsed = this.getParsedDoc();
        const doc    = parsed['VisioDocument'] as Record<string, any>;

        this.ensureStyleSheets(doc);

        const styleSheets = doc['StyleSheets'] as Record<string, any>;
        const existing    = this.normalizeArray(styleSheets['StyleSheet']);
        const newId       = this.nextId(existing);

        existing.push(this.buildStyleSheetXml(newId, name, props));
        styleSheets['StyleSheet'] = existing;

        this.saveParsedDoc(parsed);
        return { id: newId, name };
    }

    /**
     * Return all stylesheets defined in the document (including the built-in ones).
     */
    getStyles(): StyleRecord[] {
        const parsed     = this.getParsedDoc();
        const doc        = parsed['VisioDocument'] as Record<string, any>;
        const styleSheets = doc?.['StyleSheets'];
        if (!styleSheets) return [];

        return this.normalizeArray(styleSheets['StyleSheet']).map((s: any) => ({
            id:   parseInt(s['@_ID'], 10),
            name: s['@_Name'] ?? s['@_NameU'] ?? `Style ${s['@_ID']}`,
        }));
    }

    // ── Private helpers ───────────────────────────────────────────────────────

    private getParsedDoc(): Record<string, any> {
        const xml = this.pkg.getFileText('visio/document.xml');
        return this.parser.parse(xml) as Record<string, any>;
    }

    private saveParsedDoc(parsed: Record<string, any>): void {
        this.pkg.updateFile('visio/document.xml', buildXml(this.builder, parsed));
    }

    /** Ensure `<StyleSheets>` contains at least Style 0 and Style 1. */
    private ensureStyleSheets(doc: Record<string, any>): void {
        if (!doc['StyleSheets']) {
            doc['StyleSheets'] = {};
        }
        const ss = doc['StyleSheets'] as Record<string, any>;
        const existing = this.normalizeArray(ss['StyleSheet']);

        const hasStyle0 = existing.some((s: any) => s['@_ID'] === '0');
        if (!hasStyle0) {
            existing.unshift({
                '@_ID': '0',
                '@_Name': 'No Style',
                '@_NameU': 'No Style',
                '@_IsCustomName': '0',
                '@_IsCustomNameU': '0',
                Cell: [
                    { '@_N': 'EnableLineProps', '@_V': '1' },
                    { '@_N': 'EnableFillProps', '@_V': '1' },
                    { '@_N': 'EnableTextProps', '@_V': '1' },
                    { '@_N': 'HideForApply',   '@_V': '0' },
                ],
            });
        }

        const hasStyle1 = existing.some((s: any) => s['@_ID'] === '1');
        if (!hasStyle1) {
            existing.push({
                '@_ID': '1',
                '@_Name': 'Normal',
                '@_NameU': 'Normal',
                '@_IsCustomName': '0',
                '@_IsCustomNameU': '0',
                '@_LineStyle': '0',
                '@_FillStyle': '0',
                '@_TextStyle': '0',
            });
        }

        ss['StyleSheet'] = existing;
    }

    private nextId(existing: any[]): number {
        if (existing.length === 0) return 2;
        const max = existing.reduce((m: number, s: any) => {
            const id = parseInt(s['@_ID'], 10);
            return isNaN(id) ? m : Math.max(m, id);
        }, 0);
        return Math.max(max + 1, 2); // user styles always start at ≥ 2
    }

    /** Normalise a fast-xml-parser value that may be undefined, a single object, or an array. */
    private normalizeArray(val: any): any[] {
        if (!val) return [];
        return Array.isArray(val) ? val : [val];
    }

    private buildStyleSheetXml(id: number, name: string, props: StyleProps): any {
        const sheet: any = {
            '@_ID':           id.toString(),
            '@_Name':         name,
            '@_NameU':        name,
            '@_IsCustomName':  '1',
            '@_IsCustomNameU': '1',
            '@_LineStyle': (props.parentLineStyleId ?? 0).toString(),
            '@_FillStyle': (props.parentFillStyleId ?? 0).toString(),
            '@_TextStyle': (props.parentTextStyleId ?? 0).toString(),
        };

        const cells: any[] = [];
        const sections: any[] = [];

        // ── Fill ────────────────────────────────────────────────────────────
        if (props.fillColor !== undefined) {
            cells.push({ '@_N': 'FillForegnd', '@_V': props.fillColor, '@_F': hexToRgb(props.fillColor) });
            cells.push({ '@_N': 'FillPattern',  '@_V': '1' });
        }

        // ── Line ────────────────────────────────────────────────────────────
        if (props.lineColor !== undefined) {
            cells.push({ '@_N': 'LineColor', '@_V': props.lineColor, '@_F': hexToRgb(props.lineColor) });
        }
        if (props.lineWeight !== undefined) {
            cells.push({ '@_N': 'LineWeight', '@_V': (props.lineWeight / 72).toString(), '@_U': 'PT' });
        }
        if (props.linePattern !== undefined) {
            cells.push({ '@_N': 'LinePattern', '@_V': props.linePattern.toString() });
        }

        // ── TextBlock ───────────────────────────────────────────────────────
        if (props.verticalAlign !== undefined) {
            cells.push({ '@_N': 'VerticalAlign', '@_V': VERT_ALIGN[props.verticalAlign] });
        }
        if (props.textMarginTop    !== undefined) cells.push({ '@_N': 'TopMargin',    '@_V': props.textMarginTop.toString(),    '@_U': 'IN' });
        if (props.textMarginBottom !== undefined) cells.push({ '@_N': 'BottomMargin', '@_V': props.textMarginBottom.toString(), '@_U': 'IN' });
        if (props.textMarginLeft   !== undefined) cells.push({ '@_N': 'LeftMargin',   '@_V': props.textMarginLeft.toString(),   '@_U': 'IN' });
        if (props.textMarginRight  !== undefined) cells.push({ '@_N': 'RightMargin',  '@_V': props.textMarginRight.toString(),  '@_U': 'IN' });

        // ── Character section (no @_T on rows — stylesheet convention) ─────
        const charCells: any[] = [];

        const colorVal = props.fontColor ?? '#000000';
        if (props.fontColor !== undefined) {
            charCells.push({ '@_N': 'Color', '@_V': colorVal, '@_F': hexToRgb(colorVal) });
        }

        let styleVal = 0;
        if (props.bold)          styleVal |= 1;
        if (props.italic)        styleVal |= 2;
        if (props.underline)     styleVal |= 4;
        if (props.strikethrough) styleVal |= 8;
        if (styleVal > 0 || props.bold !== undefined || props.italic !== undefined
            || props.underline !== undefined || props.strikethrough !== undefined) {
            charCells.push({ '@_N': 'Style', '@_V': styleVal.toString() });
        }

        if (props.fontSize !== undefined) {
            charCells.push({ '@_N': 'Size', '@_V': (props.fontSize / 72).toString(), '@_U': 'PT' });
        }

        if (props.fontFamily !== undefined) {
            charCells.push({ '@_N': 'Font', '@_V': '0', '@_F': `FONT("${props.fontFamily}")` });
        }

        if (charCells.length > 0) {
            sections.push({
                '@_N': 'Character',
                Row: { '@_IX': '0', Cell: charCells },
            });
        }

        // ── Paragraph section ───────────────────────────────────────────────
        const paraCells: any[] = [];
        if (props.horzAlign  !== undefined) paraCells.push({ '@_N': 'HorzAlign', '@_V': HORZ_ALIGN[props.horzAlign] });
        if (props.spaceBefore !== undefined) paraCells.push({ '@_N': 'SpBefore',  '@_V': (props.spaceBefore / 72).toString(), '@_U': 'PT' });
        if (props.spaceAfter  !== undefined) paraCells.push({ '@_N': 'SpAfter',   '@_V': (props.spaceAfter  / 72).toString(), '@_U': 'PT' });
        if (props.lineSpacing !== undefined) paraCells.push({ '@_N': 'SpLine',    '@_V': (-props.lineSpacing).toString() });

        if (paraCells.length > 0) {
            sections.push({
                '@_N': 'Paragraph',
                Row: { '@_IX': '0', Cell: paraCells },
            });
        }

        if (cells.length    > 0) sheet.Cell    = cells;
        if (sections.length > 0) sheet.Section = sections;

        return sheet;
    }
}
