import { VisioPackage } from '../VisioPackage';
import { ColorEntry } from '../types/VisioTypes';
import { createXmlParser, createXmlBuilder, buildXml } from '../utils/XmlHelper';

/**
 * Manages the document-level color palette stored in `<Colors>` inside
 * `visio/document.xml`.
 *
 * The palette is a simple indexed table — Visio identifies colors by their
 * zero-based integer index (IX) as well as by their hex RGB value.
 *
 * Built-in entries (always present):
 *   - IX 0  →  #000000  (black)
 *   - IX 1  →  #FFFFFF  (white)
 *
 * User-added colors receive sequential indices starting at 2.
 */
export class ColorManager {
    private parser  = createXmlParser();
    private builder = createXmlBuilder();

    constructor(private pkg: VisioPackage) {}

    // ── Public API ────────────────────────────────────────────────────────────

    /**
     * Add a color to the document palette and return its integer index (IX).
     *
     * If the color is already in the palette the existing index is returned
     * without creating a duplicate. Built-in colors (black = 0, white = 1)
     * are returned directly.
     *
     * @param hex  CSS hex string — `'#4472C4'`, `'#abc'`, `'4472c4'` all accepted.
     * @returns    Integer IX that can be used as a color reference.
     *
     * @example
     * const ix = doc.addColor('#4472C4');  // → 2 (first user color)
     * await page.addShape({ text: 'Hi', fillColor: '#4472C4', ... });
     */
    addColor(hex: string): number {
        const normalized = normalizeHex(hex);
        const parsed     = this.getParsedDoc();
        const doc        = parsed['VisioDocument'] as Record<string, any>;

        this.ensureColors(doc);

        const colors  = doc['Colors'] as Record<string, any>;
        const entries = toArray(colors['ColorEntry']);

        // De-duplicate: return existing index if the color is already registered
        const existing = entries.find((e: any) => normalizeHex(e['@_RGB']) === normalized);
        if (existing) return parseInt(existing['@_IX'], 10);

        // Append at the next sequential index
        const nextIX = this.nextIndex(entries);
        entries.push({ '@_IX': nextIX.toString(), '@_RGB': normalized });
        colors['ColorEntry'] = entries;

        this.saveParsedDoc(parsed);
        return nextIX;
    }

    /**
     * Return all color entries currently in the document palette,
     * ordered by index ascending.
     */
    getColors(): ColorEntry[] {
        const parsed = this.getParsedDoc();
        const doc    = parsed['VisioDocument'] as Record<string, any>;
        const colors = doc?.['Colors'];
        if (!colors) return this.builtIns();

        return toArray(colors['ColorEntry'])
            .map((e: any) => ({
                index: parseInt(e['@_IX'], 10),
                rgb:   normalizeHex(e['@_RGB']),
            }))
            .sort((a, b) => a.index - b.index);
    }

    /**
     * Look up the index of a color in the palette by its hex value.
     * Returns `undefined` if the color has not been registered.
     *
     * @example
     * const ix = doc.getColorIndex('#4472C4');  // 2, or undefined if not added
     */
    getColorIndex(hex: string): number | undefined {
        const normalized = normalizeHex(hex);
        return this.getColors().find(e => e.rgb === normalized)?.index;
    }

    // ── Private helpers ───────────────────────────────────────────────────────

    private getParsedDoc(): Record<string, any> {
        const xml = this.pkg.getFileText('visio/document.xml');
        return this.parser.parse(xml) as Record<string, any>;
    }

    private saveParsedDoc(parsed: Record<string, any>): void {
        this.pkg.updateFile('visio/document.xml', buildXml(this.builder, parsed));
    }

    /** Ensure `<Colors>` exists and contains the two built-in entries. */
    private ensureColors(doc: Record<string, any>): void {
        if (!doc['Colors']) doc['Colors'] = {};

        const colors  = doc['Colors'] as Record<string, any>;
        const entries = toArray(colors['ColorEntry']);

        if (!entries.some((e: any) => e['@_IX'] === '0')) {
            entries.unshift({ '@_IX': '0', '@_RGB': '#000000' });
        }
        if (!entries.some((e: any) => e['@_IX'] === '1')) {
            const blackPos = entries.findIndex((e: any) => e['@_IX'] === '0');
            entries.splice(blackPos + 1, 0, { '@_IX': '1', '@_RGB': '#FFFFFF' });
        }

        colors['ColorEntry'] = entries;
    }

    /** Next IX = max existing index + 1, always at least 2. */
    private nextIndex(entries: any[]): number {
        const max = entries.reduce((m: number, e: any) => {
            const ix = parseInt(e['@_IX'], 10);
            return isNaN(ix) ? m : Math.max(m, ix);
        }, 1);
        return max + 1;
    }

    /** Minimal built-in palette returned when <Colors> is absent. */
    private builtIns(): ColorEntry[] {
        return [
            { index: 0, rgb: '#000000' },
            { index: 1, rgb: '#FFFFFF' },
        ];
    }
}

// ── Module-level helpers ──────────────────────────────────────────────────────

/** Normalise any hex string to uppercase `#RRGGBB`. */
function normalizeHex(hex: string): string {
    let h = hex.startsWith('#') ? hex.slice(1) : hex;
    if (h.length === 3) h = h[0]+h[0]+h[1]+h[1]+h[2]+h[2];
    return '#' + h.toUpperCase();
}

function toArray(val: any): any[] {
    if (!val) return [];
    return Array.isArray(val) ? val : [val];
}
