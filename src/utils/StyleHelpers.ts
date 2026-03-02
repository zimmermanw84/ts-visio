import type { HorzAlign, VertAlign } from '../types/VisioTypes';
export type { HorzAlign, VertAlign };

export interface VisioSection {
    '@_N': string;
    '@_IX'?: string;
    Row?: any[];
    Cell?: any[];
}

export const hexToRgb = (hex: string): string => {
    // Expand shorthand form (e.g. "03F") to full form (e.g. "0033FF")
    const shorthandRegex = /^#?([a-f\d])([a-f\d])([a-f\d])$/i;
    hex = hex.replace(shorthandRegex, (m, r, g, b) => r + r + g + g + b + b);

    const result = /^#?([a-f\d]{2})([a-f\d]{2})([a-f\d]{2})$/i.exec(hex);
    return result ? `RGB(${parseInt(result[1], 16)},${parseInt(result[2], 16)},${parseInt(result[3], 16)})` : 'RGB(0,0,0)';
};

export function createFillSection(hexColor: string): VisioSection {
    // Visio uses FillForegnd for the main background color.
    // Ideally we should sanitize hexColor to be #RRGGBB.
    const rgbFormula = hexToRgb(hexColor);
    return {
        '@_N': 'Fill',
        Cell: [
            { '@_N': 'FillForegnd', '@_V': hexColor, '@_F': rgbFormula },
            { '@_N': 'FillBkgnd', '@_V': '#FFFFFF' }, // Default background pattern color usually white
            { '@_N': 'FillPattern', '@_V': '1' }      // 1 = Solid fill
        ]
    };
}

export const ArrowHeads = {
    None: '0',
    Standard: '1',
    Open: '2',
    Stealth: '3',
    Diamond: '4',
    Oneway: '5',
    CrowsFoot: '29', // Visio "Many"
    One: '24',       // Visio "One" (Dash) - Approximate, or '26'
    // There are many variants, but 29 is the standard "Fork"
};

const HORZ_ALIGN_VALUES: Record<HorzAlign, string> = {
    left: '0',
    center: '1',
    right: '2',
    justify: '3',
};

const VERT_ALIGN_VALUES: Record<VertAlign, string> = {
    top: '0',
    middle: '1',
    bottom: '2',
};

export function horzAlignValue(align: HorzAlign): string {
    return HORZ_ALIGN_VALUES[align];
}

export function vertAlignValue(align: VertAlign): string {
    return VERT_ALIGN_VALUES[align];
}

export function createCharacterSection(props: {
    bold?: boolean;
    /** Italic text (Style bit 2). */
    italic?: boolean;
    /** Underline text (Style bit 4). */
    underline?: boolean;
    /** Strikethrough text (Style bit 8). */
    strikethrough?: boolean;
    color?: string;
    /** Font size in points (e.g. 12 for 12pt). Stored internally as inches (pt / 72). */
    fontSize?: number;
    /** Font family name (e.g. "Arial"). Uses FONT() formula for portability. */
    fontFamily?: string;
}): VisioSection {
    let styleVal = 0;
    if (props.bold)          styleVal |= 1;
    if (props.italic)        styleVal |= 2;
    if (props.underline)     styleVal |= 4;
    if (props.strikethrough) styleVal |= 8;

    const colorVal = props.color || '#000000';

    const cells: any[] = [
        { '@_N': 'Color', '@_V': colorVal, '@_F': hexToRgb(colorVal) },
        { '@_N': 'Style', '@_V': styleVal.toString() },
    ];

    if (props.fontSize !== undefined) {
        // Visio stores size in inches internally; @_U="PT" is a display hint for the ShapeSheet UI
        const sizeInInches = props.fontSize / 72;
        cells.push({ '@_N': 'Size', '@_V': sizeInInches.toString(), '@_U': 'PT' });
    }

    if (props.fontFamily !== undefined) {
        // FONT("name") formula lets Visio resolve the font by name at load time.
        // @_V="0" is a safe placeholder (document default font) used before Visio evaluates the formula.
        cells.push({ '@_N': 'Font', '@_V': '0', '@_F': `FONT("${props.fontFamily}")` });
    } else {
        cells.push({ '@_N': 'Font', '@_V': '1' }); // Default (Calibri)
    }

    return {
        '@_N': 'Character',
        Row: [
            {
                '@_T': 'Character',
                '@_IX': '0',
                Cell: cells,
            }
        ]
    };
}

export interface ParagraphProps {
    /** Horizontal text alignment within the paragraph. */
    horzAlign?: HorzAlign;
    /**
     * Space before each paragraph in **points**.
     * Converted to inches internally (pt / 72).
     */
    spaceBefore?: number;
    /**
     * Space after each paragraph in **points**.
     * Converted to inches internally (pt / 72).
     */
    spaceAfter?: number;
    /**
     * Line-height multiplier (1.0 = single, 1.5 = 1.5×, 2.0 = double).
     * Stored as a negative value in Visio's `SpLine` cell (negative means
     * proportional; positive means absolute in inches).
     */
    lineSpacing?: number;
}

export function createParagraphSection(props: ParagraphProps): VisioSection {
    const cells: any[] = [];

    if (props.horzAlign !== undefined) {
        cells.push({ '@_N': 'HorzAlign', '@_V': HORZ_ALIGN_VALUES[props.horzAlign] });
    }
    if (props.spaceBefore !== undefined) {
        cells.push({ '@_N': 'SpBefore', '@_V': (props.spaceBefore / 72).toString(), '@_U': 'PT' });
    }
    if (props.spaceAfter !== undefined) {
        cells.push({ '@_N': 'SpAfter', '@_V': (props.spaceAfter / 72).toString(), '@_U': 'PT' });
    }
    if (props.lineSpacing !== undefined) {
        // Negative value = proportional multiplier; positive = absolute (inches)
        cells.push({ '@_N': 'SpLine', '@_V': (-props.lineSpacing).toString() });
    }

    return {
        '@_N': 'Paragraph',
        Row: [
            {
                '@_T': 'Paragraph',
                '@_IX': '0',
                Cell: cells,
            }
        ]
    };
}

export interface TextBlockProps {
    /** Top text margin in inches. */
    topMargin?: number;
    /** Bottom text margin in inches. */
    bottomMargin?: number;
    /** Left text margin in inches. */
    leftMargin?: number;
    /** Right text margin in inches. */
    rightMargin?: number;
}

export function createTextBlockSection(props: TextBlockProps): VisioSection {
    const cells: any[] = [];
    if (props.topMargin    !== undefined) cells.push({ '@_N': 'TopMargin',    '@_V': props.topMargin.toString(),    '@_U': 'IN' });
    if (props.bottomMargin !== undefined) cells.push({ '@_N': 'BottomMargin', '@_V': props.bottomMargin.toString(), '@_U': 'IN' });
    if (props.leftMargin   !== undefined) cells.push({ '@_N': 'LeftMargin',   '@_V': props.leftMargin.toString(),   '@_U': 'IN' });
    if (props.rightMargin  !== undefined) cells.push({ '@_N': 'RightMargin',  '@_V': props.rightMargin.toString(),  '@_U': 'IN' });
    return {
        '@_N': 'TextBlock',
        Cell: cells,
    };
}

export function createLineSection(props: {
    color?: string;
    pattern?: string;
    weight?: string;
}): VisioSection {
    const cells: any[] = [
        { '@_N': 'LineColor', '@_V': props.color || '#000000' },
        { '@_N': 'LinePattern', '@_V': props.pattern || '1' }, // 1 = Solid
        { '@_N': 'LineWeight', '@_V': props.weight || '0.01', '@_U': 'IN' } // ~0.72pt
    ];

    // Add RGB Formula for custom colors
    if (props.color) {
        cells[0]['@_F'] = hexToRgb(props.color);
    }

    return {
        '@_N': 'Line',
        Cell: cells
    };
}
