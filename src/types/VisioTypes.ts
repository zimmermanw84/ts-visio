export interface VisioCell {
    N: string; // Name (e.g. "Width")
    V: string; // Value (e.g. "2.5")
    U?: string; // Unit (e.g. "IN")
    F?: string; // Formula (e.g. "Width*0.5")
}

export interface VisioRow {
    T?: string; // Type
    N?: string; // Name (e.g. Prop.Name)
    IX?: number; // Index
    Cells: { [name: string]: VisioCell }; // Named cells within the row
}

export interface VisioSection {
    N: string; // Name (e.g. "Geometry")
    Rows: VisioRow[];
    Cells?: { [name: string]: VisioCell }; // Direct cells for sections like Line/Fill
}

export interface VisioShape {
    ID: string;
    Name: string;
    NameU?: string; // Universal Name
    Type: string;   // e.g. "Shape" or "Group"
    Master?: string; // Master ID reference
    Text?: string;

    // ShapeSheet Data
    Cells: { [name: string]: VisioCell }; // Top-level cells
    Sections: { [name: string]: VisioSection };
}

export interface VisioConnect {
    FromSheet: string;
    FromCell: string;
    FromPart?: number;
    ToSheet: string;
    ToCell: string;
    ToPart?: number;
}

export interface VisioPage {
    ID: string;
    Name: string;
    NameU?: string;
    /** Resolved OPC part path (e.g. "visio/pages/page2.xml"). When present,
     *  this takes precedence over the ID-derived fallback path so that loaded
     *  files with non-sequential page filenames are handled correctly. */
    xmlPath?: string;
    Shapes: VisioShape[];
    Connects: VisioConnect[];
    isBackground?: boolean;
    backPageId?: string;

    // PageSheet
    PageSheet?: {
        Cells: { [name: string]: VisioCell };
        Sections: { [name: string]: VisioSection };
    };
}

export enum VisioPropType {
    String = 0,
    FixedList = 1,
    Number = 2,
    Boolean = 3,
    VariableList = 4,
    Date = 5,
    Duration = 6,
    Currency = 7
}

/** Document-level metadata that maps to `docProps/core.xml` and `docProps/app.xml`. */
export interface DocumentMetadata {
    /** Document title (`dc:title`). */
    title?: string;
    /** Author / creator (`dc:creator`). */
    author?: string;
    /** Short description (`dc:description`). */
    description?: string;
    /** Space-separated keywords (`cp:keywords`). */
    keywords?: string;
    /** Last-modified-by user (`cp:lastModifiedBy`). */
    lastModifiedBy?: string;
    /** Company name from `app.xml` `<Company>`. */
    company?: string;
    /** Manager name from `app.xml` `<Manager>`. */
    manager?: string;
    /** Document creation timestamp (`dcterms:created`). */
    created?: Date;
    /** Last-modified timestamp (`dcterms:modified`). */
    modified?: Date;
}

export type PageOrientation = 'portrait' | 'landscape';

/** Common paper sizes in inches (width × height in portrait orientation). */
export const PageSizes = {
    Letter:  { width: 8.5,    height: 11 },
    Legal:   { width: 8.5,    height: 14 },
    Tabloid: { width: 11,     height: 17 },
    A3:      { width: 11.693, height: 16.535 },
    A4:      { width: 8.268,  height: 11.693 },
    A5:      { width: 5.827,  height: 8.268 },
} as const;

export type PageSizeName = keyof typeof PageSizes;

/**
 * Physical length units supported for drawing-scale cells.
 * Imperial: `'in'` (inches), `'ft'` (feet), `'yd'` (yards), `'mi'` (miles)
 * Metric:   `'mm'`, `'cm'`, `'m'`, `'km'`
 */
export type LengthUnit = 'in' | 'ft' | 'yd' | 'mi' | 'mm' | 'cm' | 'm' | 'km';

/**
 * Drawing-scale description returned by `page.getDrawingScale()` and
 * accepted by `page.setDrawingScale()`.
 *
 * The ratio `drawingScale / pageScale` (after unit conversion to a common
 * base) is the factor by which one page unit maps to real-world distance.
 *
 * @example
 * // 1 inch on paper = 10 feet in the real world
 * { pageScale: 1, pageUnit: 'in', drawingScale: 10, drawingUnit: 'ft' }
 *
 * // 1:100 metric
 * { pageScale: 1, pageUnit: 'cm', drawingScale: 100, drawingUnit: 'cm' }
 */
export interface DrawingScaleInfo {
    /** Measurement on the paper (e.g. 1). */
    pageScale: number;
    /** Unit for the paper measurement. */
    pageUnit: LengthUnit;
    /** Corresponding real-world measurement (e.g. 10). */
    drawingScale: number;
    /** Unit for the real-world measurement. */
    drawingUnit: LengthUnit;
}

/** Connector line-routing algorithm. */
export type ConnectorRouting = 'straight' | 'orthogonal' | 'curved';

/**
 * Style options for a connector (dynamic connector shape).
 * All fields are optional; omitted fields retain their defaults.
 */
export interface ConnectorStyle {
    /** Line stroke color as a CSS hex string (e.g. `'#ff0000'`). */
    lineColor?: string;
    /**
     * Stroke weight in **points** (e.g. `1` for a 1 pt line).
     * Converted to inches internally (pt / 72).
     */
    lineWeight?: number;
    /**
     * Line pattern.  0 = no line, 1 = solid (default), 2 = dashed,
     * 3 = dotted, 4 = dash-dot, etc.  Matches Visio's LinePattern cell.
     */
    linePattern?: number;
    /** How Visio routes the connector between its endpoints. */
    routing?: ConnectorRouting;
}

/** Horizontal text alignment within a paragraph. */
export type HorzAlign = 'left' | 'center' | 'right' | 'justify';

/** Vertical text alignment within a shape's text block. */
export type VertAlign = 'top' | 'middle' | 'bottom';

/** Non-rectangular geometry variants supported by ShapeBuilder. */
export type ShapeGeometry =
    | 'rectangle'
    | 'ellipse'
    | 'diamond'
    | 'rounded-rectangle'
    | 'triangle'
    | 'parallelogram';

/**
 * A master shape definition — returned by `doc.getMasters()` and `doc.createMaster()`.
 * The `id` can be passed as `masterId` when calling `page.addShape()` to stamp an
 * instance of the master onto a page.
 */
export interface MasterRecord {
    /** String integer ID (matches `@_Master` attribute on shape instances). */
    id: string;
    /** Display name (locale-specific). */
    name: string;
    /** Universal (locale-independent) name. */
    nameU: string;
    /** OPC path to the individual master content file, e.g. `"visio/masters/master1.xml"`. */
    xmlPath: string;
}

/**
 * A single entry in the document-level color palette (`<Colors>` in `document.xml`).
 * Returned by `doc.getColors()` and created via `doc.addColor()`.
 */
export interface ColorEntry {
    /**
     * Zero-based integer index (IX). Can be passed as a color reference
     * anywhere a hex string is accepted (e.g. `fillColor`, `lineColor`).
     */
    index: number;
    /** Normalized CSS hex string, always uppercase `#RRGGBB`. */
    rgb: string;
}

/** A reference to a document-level stylesheet, returned by `doc.createStyle()`. */
export interface StyleRecord {
    /** Zero-based integer ID used as `LineStyle` / `FillStyle` / `TextStyle` on shapes. */
    id: number;
    /** Human-readable name given to the style. */
    name: string;
}

/**
 * Visual properties for a document-level stylesheet created via `doc.createStyle()`.
 * All fields are optional — omitted properties are inherited from the parent style (default: Style 0 "No Style").
 */
export interface StyleProps {
    /** Parent style for line property inheritance. Defaults to 0 ("No Style"). */
    parentLineStyleId?: number;
    /** Parent style for fill property inheritance. Defaults to 0 ("No Style"). */
    parentFillStyleId?: number;
    /** Parent style for text property inheritance. Defaults to 0 ("No Style"). */
    parentTextStyleId?: number;

    // ── Line ──────────────────────────────────────────────────────────────────
    /** Stroke colour as a CSS hex string (e.g. `'#cc0000'`). */
    lineColor?: string;
    /** Stroke weight in **points**. Stored internally as inches (pt / 72). */
    lineWeight?: number;
    /** Line pattern. 0 = none, 1 = solid, 2 = dash, 3 = dot, 4 = dash-dot. */
    linePattern?: number;

    // ── Fill ──────────────────────────────────────────────────────────────────
    /** Background fill colour as a CSS hex string. */
    fillColor?: string;

    // ── Character (text run) ──────────────────────────────────────────────────
    /** Text colour as a CSS hex string. */
    fontColor?: string;
    /** Font size in **points**. */
    fontSize?: number;
    /** Bold text. */
    bold?: boolean;
    /** Italic text. */
    italic?: boolean;
    /** Underline text. */
    underline?: boolean;
    /** Strikethrough text. */
    strikethrough?: boolean;
    /** Font family name (e.g. `'Calibri'`). */
    fontFamily?: string;

    // ── Paragraph ─────────────────────────────────────────────────────────────
    /** Horizontal text alignment within the paragraph. */
    horzAlign?: HorzAlign;
    /** Space before each paragraph in **points**. */
    spaceBefore?: number;
    /** Space after each paragraph in **points**. */
    spaceAfter?: number;
    /** Line-height multiplier (1.0 = single, 1.5 = 1.5×, 2.0 = double). */
    lineSpacing?: number;

    // ── TextBlock ─────────────────────────────────────────────────────────────
    /** Vertical text alignment within the shape. */
    verticalAlign?: VertAlign;
    /** Top text margin in inches. */
    textMarginTop?: number;
    /** Bottom text margin in inches. */
    textMarginBottom?: number;
    /** Left text margin in inches. */
    textMarginLeft?: number;
    /** Right text margin in inches. */
    textMarginRight?: number;
}

/** Whether a connection point accepts incoming glue, outgoing glue, or both. */
export type ConnectionPointType = 'inward' | 'outward' | 'both';

/**
 * Definition of a single connection point on a shape.
 * Positions are expressed as fractions of the shape's width/height
 * (0.0 = left/bottom, 1.0 = right/top in Visio's coordinate system).
 */
export interface ConnectionPointDef {
    /** Optional display name (e.g. `'Top'`, `'Right'`). Referenced when connecting by name. */
    name?: string;
    /** Horizontal position as a fraction of shape width. 0 = left edge, 1 = right edge. */
    xFraction: number;
    /** Vertical position as a fraction of shape height. 0 = bottom edge, 1 = top edge. */
    yFraction: number;
    /** Glue direction vector. Defaults to `{ x: 0, y: 0 }` (no preferred direction). */
    direction?: { x: number; y: number };
    /** Connection point type. Defaults to `'inward'`. */
    type?: ConnectionPointType;
    /** Optional tooltip shown in the Visio UI. */
    prompt?: string;
}

/**
 * Specifies which connection point to use when attaching a connector endpoint.
 * - `'center'`       → shape centre (default behaviour, `ToPart=3`)
 * - `{ name }`       → named connection point (e.g. `{ name: 'Top' }`)
 * - `{ index }`      → connection point by zero-based row index
 */
export type ConnectionTarget =
    | 'center'
    | { name: string }
    | { index: number };

/** Ready-made connection-point presets. */
export const StandardConnectionPoints: {
    /** Four cardinal points: Top, Right, Bottom, Left. */
    cardinal: ConnectionPointDef[];
    /** Eight points: four cardinal + four corners. */
    full: ConnectionPointDef[];
} = {
    cardinal: [
        { name: 'Top',    xFraction: 0.5, yFraction: 1.0, direction: { x:  0, y:  1 } },
        { name: 'Right',  xFraction: 1.0, yFraction: 0.5, direction: { x:  1, y:  0 } },
        { name: 'Bottom', xFraction: 0.5, yFraction: 0.0, direction: { x:  0, y: -1 } },
        { name: 'Left',   xFraction: 0.0, yFraction: 0.5, direction: { x: -1, y:  0 } },
    ],
    full: [
        { name: 'Top',         xFraction: 0.5, yFraction: 1.0, direction: { x:  0, y:  1 } },
        { name: 'Right',       xFraction: 1.0, yFraction: 0.5, direction: { x:  1, y:  0 } },
        { name: 'Bottom',      xFraction: 0.5, yFraction: 0.0, direction: { x:  0, y: -1 } },
        { name: 'Left',        xFraction: 0.0, yFraction: 0.5, direction: { x: -1, y:  0 } },
        { name: 'TopLeft',     xFraction: 0.0, yFraction: 1.0, direction: { x: -1, y:  1 } },
        { name: 'TopRight',    xFraction: 1.0, yFraction: 1.0, direction: { x:  1, y:  1 } },
        { name: 'BottomRight', xFraction: 1.0, yFraction: 0.0, direction: { x:  1, y: -1 } },
        { name: 'BottomLeft',  xFraction: 0.0, yFraction: 0.0, direction: { x: -1, y: -1 } },
    ],
};

export interface NewShapeProps {
    text: string;
    x: number;
    y: number;
    width: number;
    height: number;
    id?: string;
    fillColor?: string;
    fontColor?: string;
    bold?: boolean;
    /** Font size in points (e.g. 14 for 14pt). */
    fontSize?: number;
    /** Font family name (e.g. "Arial", "Times New Roman"). */
    fontFamily?: string;
    /** Horizontal text alignment within the shape. */
    horzAlign?: HorzAlign;
    /** Vertical text alignment within the shape. */
    verticalAlign?: VertAlign;
    /** Shape geometry. Defaults to 'rectangle'. */
    geometry?: ShapeGeometry;
    /** Corner radius in inches for 'rounded-rectangle'. Defaults to 10% of the smaller dimension. */
    cornerRadius?: number;
    type?: string;
    masterId?: string;
    imgRelId?: string;
    lineColor?: string;
    /** Line pattern. 0 = none, 1 = solid (default), 2 = dash, 3 = dot, 4 = dash-dot. */
    linePattern?: number;
    /** Italic text. */
    italic?: boolean;
    /** Underline text. */
    underline?: boolean;
    /** Strikethrough text. */
    strikethrough?: boolean;
    /** Space before each paragraph in **points**. */
    spaceBefore?: number;
    /** Space after each paragraph in **points**. */
    spaceAfter?: number;
    /** Line-height multiplier (1.0 = single, 1.5 = 1.5×, 2.0 = double). */
    lineSpacing?: number;
    /** Top text margin in inches. */
    textMarginTop?: number;
    /** Bottom text margin in inches. */
    textMarginBottom?: number;
    /** Left text margin in inches. */
    textMarginLeft?: number;
    /** Right text margin in inches. */
    textMarginRight?: number;
    /**
     * Connection points to add to the shape.
     * Use `StandardConnectionPoints.cardinal` for the four cardinal points,
     * or `StandardConnectionPoints.full` for eight points.
     */
    connectionPoints?: ConnectionPointDef[];

    /**
     * Apply a document-level stylesheet to this shape (sets `LineStyle`, `FillStyle`,
     * and `TextStyle` all to the same ID). Create styles via `doc.createStyle()`.
     * Takes precedence over `lineStyleId`, `fillStyleId`, and `textStyleId`.
     */
    styleId?: number;
    /** Apply a stylesheet only for line properties (`LineStyle` attribute). */
    lineStyleId?: number;
    /** Apply a stylesheet only for fill properties (`FillStyle` attribute). */
    fillStyleId?: number;
    /** Apply a stylesheet only for text properties (`TextStyle` attribute). */
    textStyleId?: number;
}

/**
 * Visual style properties applied to an existing shape via `shape.setStyle()`.
 * All fields are optional; omitted properties are left unchanged.
 */
export interface ShapeStyle {
    fillColor?: string;
    /** Border/stroke colour as a CSS hex string (e.g. `'#cc0000'`). */
    lineColor?: string;
    /** Stroke weight in **points**. Stored internally as inches (pt / 72). */
    lineWeight?: number;
    /** Line pattern. 0 = none, 1 = solid (default), 2 = dash, 3 = dot, 4 = dash-dot. */
    linePattern?: number;
    fontColor?: string;
    bold?: boolean;
    /** Italic text. */
    italic?: boolean;
    /** Underline text. */
    underline?: boolean;
    /** Strikethrough text. */
    strikethrough?: boolean;
    /** Font size in points (e.g. 14 for 14pt). */
    fontSize?: number;
    /** Font family name (e.g. "Arial"). */
    fontFamily?: string;
    /** Horizontal text alignment. */
    horzAlign?: HorzAlign;
    /** Vertical text alignment. */
    verticalAlign?: VertAlign;
    /** Space before each paragraph in **points**. */
    spaceBefore?: number;
    /** Space after each paragraph in **points**. */
    spaceAfter?: number;
    /** Line-height multiplier (1.0 = single, 1.5 = 1.5×, 2.0 = double). */
    lineSpacing?: number;
    /** Top text margin in inches. */
    textMarginTop?: number;
    /** Bottom text margin in inches. */
    textMarginBottom?: number;
    /** Left text margin in inches. */
    textMarginLeft?: number;
    /** Right text margin in inches. */
    textMarginRight?: number;
}
