import { ShapeGeometry } from '../types/VisioTypes';

/**
 * For a 90° circular arc of radius r, the midpoint on the arc sits at 45° from
 * the corner. Its distance from each endpoint along the arc edge is:
 *   r * (1 - cos(45°)) = r * (1 - √2/2) ≈ r * 0.29289
 * This offset is used to compute the A, B control-point cells of EllipticalArcTo.
 */
const ARC_OFFSET = 1 - Math.SQRT2 / 2; // ≈ 0.29289321881345254

// ---------------------------------------------------------------------------
// Internal helpers
// ---------------------------------------------------------------------------

function cell(n: string, value: number | string, formula?: string): any {
    const obj: any = { '@_N': n, '@_V': typeof value === 'number' ? value.toString() : value };
    if (formula) obj['@_F'] = formula;
    return obj;
}

function moveTo(ix: number, x: number, y: number, xF?: string, yF?: string): any {
    return {
        '@_T': 'MoveTo',
        '@_IX': String(ix),
        Cell: [cell('X', x, xF), cell('Y', y, yF)],
    };
}

function lineTo(ix: number, x: number, y: number, xF?: string, yF?: string): any {
    return {
        '@_T': 'LineTo',
        '@_IX': String(ix),
        Cell: [cell('X', x, xF), cell('Y', y, yF)],
    };
}

function arcTo(ix: number, endX: number, endY: number, midX: number, midY: number): any {
    return {
        '@_T': 'EllipticalArcTo',
        '@_IX': String(ix),
        Cell: [
            cell('X', endX),   // arc endpoint X
            cell('Y', endY),   // arc endpoint Y
            cell('A', midX),   // midpoint-on-arc X
            cell('B', midY),   // midpoint-on-arc Y
            cell('C', '0'),    // major-axis angle (0 = no rotation)
            cell('D', '1'),    // eccentricity ratio (1 = circle)
        ],
    };
}

function section(rows: any[], noFill: string): any {
    return {
        '@_N': 'Geometry',
        '@_IX': '0',
        Cell: [
            { '@_N': 'NoShow', '@_V': '0' },
            { '@_N': 'NoFill', '@_V': noFill },
        ],
        Row: rows,
    };
}

// ---------------------------------------------------------------------------
// Public builder
// ---------------------------------------------------------------------------

export class GeometryBuilder {
    /**
     * Build a Geometry section object for the given props.
     * Defaults to 'rectangle' when geometry is omitted.
     */
    static build(props: {
        width: number;
        height: number;
        geometry?: ShapeGeometry;
        cornerRadius?: number;
        fillColor?: string;
    }): any {
        const { width: W, height: H, geometry, cornerRadius, fillColor } = props;
        const noFill = fillColor ? '0' : '1';

        switch (geometry) {
            case 'ellipse':
                return GeometryBuilder.ellipse(W, H, noFill);
            case 'diamond':
                return GeometryBuilder.diamond(W, H, noFill);
            case 'rounded-rectangle': {
                const r = cornerRadius ?? Math.min(W, H) * 0.1;
                return GeometryBuilder.roundedRectangle(W, H, r, noFill);
            }
            case 'triangle':
                return GeometryBuilder.triangle(W, H, noFill);
            case 'parallelogram':
                return GeometryBuilder.parallelogram(W, H, noFill);
            case 'rectangle':
            default:
                return GeometryBuilder.rectangle(W, H, noFill);
        }
    }

    // -----------------------------------------------------------------------
    // Individual geometry factories (also exported for direct use / testing)
    // -----------------------------------------------------------------------

    /** Standard rectangle: 4 LineTo rows starting at origin. */
    static rectangle(W: number, H: number, noFill: string): any {
        return section([
            moveTo(1, 0,  0),
            lineTo(2, W,  0, 'Width'),
            lineTo(3, W,  H, 'Width', 'Height'),
            lineTo(4, 0,  H, undefined, 'Height'),
            lineTo(5, 0,  0),
        ], noFill);
    }

    /**
     * Ellipse: a single Ellipse row. No MoveTo required.
     * X, Y = centre; A, B = rightmost point; C, D = topmost point.
     */
    static ellipse(W: number, H: number, noFill: string): any {
        return section([
            {
                '@_T': 'Ellipse',
                '@_IX': '1',
                Cell: [
                    cell('X', W / 2, 'Width*0.5'),    // centre X
                    cell('Y', H / 2, 'Height*0.5'),   // centre Y
                    cell('A', W,     'Width'),          // right X
                    cell('B', H / 2, 'Height*0.5'),   // right Y
                    cell('C', W / 2, 'Width*0.5'),    // top X
                    cell('D', H,     'Height'),         // top Y
                ],
            },
        ], noFill);
    }

    /**
     * Diamond: 4 LineTo rows, starting at top vertex and going clockwise
     * (top → right → bottom → left → top), matching Visio's built-in Decision shape.
     */
    static diamond(W: number, H: number, noFill: string): any {
        return section([
            moveTo(1, W / 2, H,     'Width*0.5', 'Height'),  // top
            lineTo(2, W,     H / 2, 'Width',     'Height*0.5'), // right
            lineTo(3, W / 2, 0,     'Width*0.5'),               // bottom
            lineTo(4, 0,     H / 2, undefined,   'Height*0.5'), // left
            lineTo(5, W / 2, H,     'Width*0.5', 'Height'),  // close
        ], noFill);
    }

    /**
     * Rounded rectangle with circular corner arcs of radius `r`.
     * Uses EllipticalArcTo with C=0 (no rotation) and D=1 (circle).
     * The A, B control cells are the arc midpoints: r*(1-√2/2) inset from each corner.
     */
    static roundedRectangle(W: number, H: number, r: number, noFill: string): any {
        const k = r * ARC_OFFSET; // midpoint inset ≈ r * 0.2929
        return section([
            moveTo(1, r,     0),
            lineTo(2, W - r, 0),
            arcTo( 3, W,     r,     W - k, k),       // bottom-right corner
            lineTo(4, W,     H - r),
            arcTo( 5, W - r, H,     W - k, H - k),  // top-right corner
            lineTo(6, r,     H),
            arcTo( 7, 0,     H - r, k,     H - k),  // top-left corner
            lineTo(8, 0,     r),
            arcTo( 9, r,     0,     k,     k),       // bottom-left corner (closes path)
        ], noFill);
    }

    /**
     * Right-pointing triangle (standard Visio flowchart orientation).
     * Vertices: bottom-left (0,0), apex-right (W, H/2), top-left (0, H).
     */
    static triangle(W: number, H: number, noFill: string): any {
        return section([
            moveTo(1, 0, 0),
            lineTo(2, W, H / 2, 'Width', 'Height*0.5'),
            lineTo(3, 0, H,     undefined, 'Height'),
            lineTo(4, 0, 0),
        ], noFill);
    }

    /**
     * Parallelogram with a rightward skew of 20% of the width.
     * Matches the proportions of Visio's built-in "Data" flowchart shape.
     */
    static parallelogram(W: number, H: number, noFill: string): any {
        const s = W * 0.2; // horizontal skew offset
        return section([
            moveTo(1, s,     0),
            lineTo(2, W,     0, 'Width'),
            lineTo(3, W - s, H),
            lineTo(4, 0,     H, undefined, 'Height'),
            lineTo(5, s,     0),
        ], noFill);
    }
}
