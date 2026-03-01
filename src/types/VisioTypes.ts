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
    type?: string;
    masterId?: string;
    imgRelId?: string;
    lineColor?: string;
}
