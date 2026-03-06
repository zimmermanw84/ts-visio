
import { ConnectionPointDef, ConnectionPointType, ConnectionTarget } from '../types/VisioTypes';

const TYPE_VALUES: Record<ConnectionPointType, string> = {
    inward:  '0',
    outward: '1',
    both:    '2',
};

/**
 * Builds and resolves Visio Connection section XML structures.
 *
 * Connection points in Visio are stored as rows in a `<Section N="Connection">` element.
 * Each row has an IX attribute (0-based) and an optional N (name) attribute.
 * X and Y cells store fractions via formula: `Width*{xFraction}` / `Height*{yFraction}`.
 *
 * When connecting to a named point, the Connect element uses:
 *   ToPart = 100 + IX   (instead of 3 for shape centre)
 *   ToCell = "Connections.X{IX+1}"  (1-based)
 */
export class ConnectionPointBuilder {

    /** Build the raw XML object for a Connection section from a list of point definitions. */
    static buildConnectionSection(points: ConnectionPointDef[], width = 1, height = 1): any {
        return {
            '@_N': 'Connection',
            Row: points.map((pt, ix) => this.buildRow(pt, ix, width, height)),
        };
    }

    /** Build a single Connection row XML object. */
    static buildRow(point: ConnectionPointDef, ix: number, width = 1, height = 1): any {
        const row: any = {
            '@_IX': ix.toString(),
            Cell: [
                { '@_N': 'X',       '@_V': (width  * point.xFraction).toString(), '@_F': `Width*${point.xFraction}`  },
                { '@_N': 'Y',       '@_V': (height * point.yFraction).toString(), '@_F': `Height*${point.yFraction}` },
                { '@_N': 'DirX',    '@_V': point.direction ? point.direction.x.toString() : '0' },
                { '@_N': 'DirY',    '@_V': point.direction ? point.direction.y.toString() : '0' },
                { '@_N': 'Type',    '@_V': point.type ? TYPE_VALUES[point.type] : '0' },
                { '@_N': 'AutoGen', '@_V': '0' },
            ],
        };
        if (point.name)   row['@_N'] = point.name;
        if (point.prompt) row.Cell.push({ '@_N': 'Prompt', '@_V': point.prompt });
        return row;
    }

    /**
     * Resolve a `ConnectionTarget` against a raw shape XML object.
     *
     * Returns the `ToPart` / `ToCell` strings for the Connect element, plus the
     * fractional X/Y position of the connection point (undefined when targeting
     * the shape centre, which falls back to edge-intersection logic).
     */
    static resolveTarget(
        target: ConnectionTarget,
        shape: any,
    ): { toPart: string; toCell: string; xFraction?: number; yFraction?: number } {
        if (target === 'center') {
            return { toPart: '3', toCell: 'PinX' };
        }

        if ('index' in target) {
            const ix = target.index;
            const pt = this.getPointByIx(shape, ix);
            return {
                toPart:    (100 + ix).toString(),
                toCell:    `Connections.X${ix + 1}`,
                xFraction: pt?.xFraction,
                yFraction: pt?.yFraction,
            };
        }

        // name lookup
        const ix = this.findIxByName(shape, target.name);
        if (ix === -1) {
            // Named point not found — fall back to centre
            return { toPart: '3', toCell: 'PinX' };
        }
        const pt = this.getPointByIx(shape, ix);
        return {
            toPart:    (100 + ix).toString(),
            toCell:    `Connections.X${ix + 1}`,
            xFraction: pt?.xFraction,
            yFraction: pt?.yFraction,
        };
    }

    // ── private helpers ────────────────────────────────────────────────────────

    private static getConnectionSection(shape: any): any | null {
        const sections = shape?.Section;
        if (!sections) return null;
        const arr = Array.isArray(sections) ? sections : [sections];
        return arr.find((s: any) => s['@_N'] === 'Connection') ?? null;
    }

    private static getRows(section: any): any[] {
        if (!section?.Row) return [];
        return Array.isArray(section.Row) ? section.Row : [section.Row];
    }

    private static findIxByName(shape: any, name: string): number {
        const section = this.getConnectionSection(shape);
        if (!section) return -1;
        const row = this.getRows(section).find((r: any) => r['@_N'] === name);
        if (!row) return -1;
        return parseInt(row['@_IX'], 10);
    }

    private static getPointByIx(
        shape: any,
        ix: number,
    ): { xFraction: number; yFraction: number } | undefined {
        const section = this.getConnectionSection(shape);
        if (!section) return undefined;

        const row = this.getRows(section).find((r: any) => parseInt(r['@_IX'], 10) === ix);
        if (!row) return undefined;

        const cells: any[] = Array.isArray(row.Cell) ? row.Cell : row.Cell ? [row.Cell] : [];
        const xCell = cells.find((c: any) => c['@_N'] === 'X');
        const yCell = cells.find((c: any) => c['@_N'] === 'Y');

        // Prefer formula parsing (works for both freshly-created and loaded shapes).
        // Fall back to @_V if no formula is present.
        const xFraction = this.parseFraction(xCell?.['@_F']) ?? parseFloat(xCell?.['@_V'] ?? '0');
        const yFraction = this.parseFraction(yCell?.['@_F']) ?? parseFloat(yCell?.['@_V'] ?? '0');

        return { xFraction, yFraction };
    }

    /**
     * Extract the numeric multiplier from a formula of the form `Width*0.5` or `Height*1`.
     * Returns `undefined` if the formula does not match that pattern.
     */
    private static parseFraction(formula: string | undefined): number | undefined {
        if (!formula) return undefined;
        const m = formula.match(/^(?:Width|Height)\*([0-9.]+)$/);
        return m ? parseFloat(m[1]) : undefined;
    }
}
