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

/** Non-rectangular geometry variants supported by ShapeBuilder. */
export type ShapeGeometry =
    | 'rectangle'
    | 'ellipse'
    | 'diamond'
    | 'rounded-rectangle'
    | 'triangle'
    | 'parallelogram';

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
    horzAlign?: 'left' | 'center' | 'right' | 'justify';
    /** Vertical text alignment within the shape. */
    verticalAlign?: 'top' | 'middle' | 'bottom';
    /** Shape geometry. Defaults to 'rectangle'. */
    geometry?: ShapeGeometry;
    /** Corner radius in inches for 'rounded-rectangle'. Defaults to 10% of the smaller dimension. */
    cornerRadius?: number;
    type?: string;
    masterId?: string;
    imgRelId?: string;
    lineColor?: string;
}
