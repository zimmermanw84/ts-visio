import { XMLParser, XMLBuilder } from 'fast-xml-parser';
import { VisioPackage } from './VisioPackage';
import { createFillSection, createCharacterSection, createLineSection } from './utils/StyleHelpers';

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
    type?: string;
}

export class ShapeModifier {
    private parser: XMLParser;
    private builder: XMLBuilder;

    constructor(private pkg: VisioPackage) {
        this.parser = new XMLParser({
            ignoreAttributes: false,
            attributeNamePrefix: "@_"
        });
        this.builder = new XMLBuilder({
            ignoreAttributes: false,
            attributeNamePrefix: "@_",
            format: true
        });
    }

    private getPagePath(pageId: string): string {
        return `visio/pages/page${pageId}.xml`;
    }

    private getAllShapes(parsed: any): any[] {
        let topLevelShapes = parsed.PageContents.Shapes ? parsed.PageContents.Shapes.Shape : [];
        if (!Array.isArray(topLevelShapes)) {
            topLevelShapes = topLevelShapes ? [topLevelShapes] : [];
        }

        const gather = (shapeList: any[]): any[] => {
            let all: any[] = [];
            for (const s of shapeList) {
                all.push(s);
                if (s.Shapes && s.Shapes.Shape) {
                    const children = Array.isArray(s.Shapes.Shape) ? s.Shapes.Shape : [s.Shapes.Shape];
                    all = all.concat(gather(children));
                }
            }
            return all;
        };

        return gather(topLevelShapes);
    }

    private getNextId(parsed: any): string {
        const allShapes = this.getAllShapes(parsed);
        let maxId = 0;
        for (const s of allShapes) {
            const id = parseInt(s['@_ID']);
            if (!isNaN(id) && id > maxId) maxId = id;
        }
        return (maxId + 1).toString();
    }

    async addConnector(pageId: string, fromShapeId: string, toShapeId: string, beginArrow?: string, endArrow?: string): Promise<string> {
        const pagePath = `visio/pages/page${pageId}.xml`;
        let content = '';

        try {
            content = this.pkg.getFileText(pagePath);
        } catch {
            throw new Error(`Could not find page file for ID ${pageId}. Expected at ${pagePath}`);
        }

        const parsed = this.parser.parse(content);

        // Ensure Shapes collection exists
        if (!parsed.PageContents.Shapes) {
            parsed.PageContents.Shapes = { Shape: [] };
        }
        if (!Array.isArray(parsed.PageContents.Shapes.Shape)) {
            parsed.PageContents.Shapes.Shape = [parsed.PageContents.Shapes.Shape];
        }

        // Generate ID
        const newId = this.getNextId(parsed);

        // Recursive Find Helper for Source/Target (since they might be inside Groups)
        const allShapes = this.getAllShapes(parsed);
        const findShape = (id: string) => allShapes.find((s: any) => s['@_ID'] == id);

        const sourceShape = findShape(fromShapeId);
        const targetShape = findShape(toShapeId);

        let beginX = '0';
        let beginY = '0';
        let endX = '0';
        let endY = '0';

        const getCellVal = (shape: any, name: string) => {
            if (!shape || !shape.Cell) return '0';
            const cell = shape.Cell.find((c: any) => c['@_N'] === name);
            return cell ? cell['@_V'] : '0';
        };

        if (sourceShape) {
            beginX = getCellVal(sourceShape, 'PinX');
            beginY = getCellVal(sourceShape, 'PinY');
        }
        if (targetShape) {
            endX = getCellVal(targetShape, 'PinX');
            endY = getCellVal(targetShape, 'PinY');
        }

        // 1. Create Connector Shape
        const connectorShape: any = {
            '@_ID': newId,
            '@_NameU': 'Dynamic connector',
            '@_Name': 'Dynamic connector',
            '@_Type': 'Shape',
            // '@_Master': '2', // Removed: We don't have masters in blank templates yet
            Cell: [
                { '@_N': 'BeginX', '@_V': beginX },
                { '@_N': 'BeginY', '@_V': beginY },
                { '@_N': 'EndX', '@_V': endX },
                { '@_N': 'EndY', '@_V': endY },
                { '@_N': 'PinX', '@_V': '0', '@_F': '(BeginX+EndX)/2' },
                { '@_N': 'PinY', '@_V': '0', '@_F': '(BeginY+EndY)/2' },
                { '@_N': 'BeginArrow', '@_V': beginArrow || '0' },
                { '@_N': 'EndArrow', '@_V': endArrow || '0' },
                // 1D Transform requires standard cells
                { '@_N': 'Width', '@_V': '0', '@_F': 'SQRT((EndX-BeginX)^2+(EndY-BeginY)^2)' },
                { '@_N': 'Height', '@_V': '0' },
                { '@_N': 'Angle', '@_V': '0', '@_F': 'ATAN2(EndY-BeginY,EndX-BeginX)' },
                { '@_N': 'LocPinX', '@_V': '0', '@_F': 'Width*0.5' }, // Center pivot
                { '@_N': 'LocPinY', '@_V': '0', '@_F': 'Height*0.5' },
                { '@_N': 'ObjType', '@_V': '2' }, // 1D Shape
                { '@_N': 'ShapePermeableX', '@_V': '0' }, // Recommended for connectors
                { '@_N': 'ShapePermeableY', '@_V': '0' },
                { '@_N': 'ShapeRouteStyle', '@_V': '1' }, // Right-Angle
                { '@_N': 'ConFixedCode', '@_V': '0' }
            ],
            Section: [
                createLineSection({ color: '#000000', weight: '0.01' }),
                {
                    '@_N': 'Geometry',
                    '@_IX': '0',
                    Row: [
                        { '@_T': 'MoveTo', '@_IX': '1', Cell: [{ '@_N': 'X', '@_V': '0' }, { '@_N': 'Y', '@_V': '0' }] },
                        { '@_T': 'LineTo', '@_IX': '2', Cell: [{ '@_N': 'X', '@_V': '0', '@_F': 'Width' }, { '@_N': 'Y', '@_V': '0', '@_F': 'Height*0' }] }
                    ]
                }
            ]
        };

        const topLevelShapes = parsed.PageContents.Shapes.Shape; // Always array due to Ensure Shapes above
        topLevelShapes.push(connectorShape);

        // 2. Add to Connects collection
        if (!parsed.PageContents.Connects) {
            parsed.PageContents.Connects = { Connect: [] };
        }

        let connectCollection = parsed.PageContents.Connects.Connect;
        // Ensure it's an array if it was a single object or undefined
        if (!Array.isArray(connectCollection)) {
            // If it was valid object but not array, wrap it. Else init empty.
            connectCollection = connectCollection ? [connectCollection] : [];
            parsed.PageContents.Connects.Connect = connectCollection;
        }

        // Add Tail Connection (BeginX -> FromShape)
        connectCollection.push({
            '@_FromSheet': newId,
            '@_FromCell': 'BeginX',
            '@_FromPart': '9', // constant for BeginX
            '@_ToSheet': fromShapeId,
            '@_ToCell': 'PinX', // Walking glue
            '@_ToPart': '3'     // constant for PinX connection
        });

        // Add Head Connection (EndX -> ToShape)
        connectCollection.push({
            '@_FromSheet': newId,
            '@_FromCell': 'EndX',
            '@_FromPart': '12', // constant for EndX
            '@_ToSheet': toShapeId,
            '@_ToCell': 'PinX',
            '@_ToPart': '3'
        });

        // Save back
        const builder = new XMLBuilder({
            ignoreAttributes: false,
            attributeNamePrefix: "@_",
            format: true
        });
        const newXml = builder.build(parsed);
        this.pkg.updateFile(pagePath, newXml);

        return newId;
    }

    async addShape(pageId: string, props: NewShapeProps, parentId?: string): Promise<string> {
        const pagePath = this.getPagePath(pageId);
        let content: string;
        try {
            content = this.pkg.getFileText(pagePath);
        } catch {
            throw new Error(`Could not find page file for ID ${pageId}. Expected at ${pagePath}`);
        }

        const parsed = this.parser.parse(content);

        // Ensure Shapes container exists
        if (!parsed.PageContents.Shapes) {
            parsed.PageContents.Shapes = { Shape: [] };
        }
        let topLevelShapes = parsed.PageContents.Shapes.Shape;
        if (!Array.isArray(topLevelShapes)) {
            topLevelShapes = topLevelShapes ? [topLevelShapes] : [];
            parsed.PageContents.Shapes.Shape = topLevelShapes;
        }

        const allShapes = this.getAllShapes(parsed);

        // Auto-generate ID if not provided
        let newId = props.id;
        if (!newId) {
            newId = this.getNextId(parsed);
        }

        const newShape: any = {
            '@_ID': newId,
            '@_Name': `Sheet.${newId}`,
            '@_Type': props.type || 'Shape', // Allow specifying Group type
            Cell: [
                { '@_N': 'PinX', '@_V': props.x.toString() },
                { '@_N': 'PinY', '@_V': props.y.toString() },
                { '@_N': 'Width', '@_V': props.width.toString() },
                { '@_N': 'Height', '@_V': props.height.toString() },
                { '@_N': 'LocPinX', '@_V': (props.width / 2).toString() },
                { '@_N': 'LocPinY', '@_V': (props.height / 2).toString() }
            ],
            Text: { '#text': props.text },
            Section: [
                {
                    '@_N': 'Geometry',
                    '@_IX': '0',
                    Row: [
                        { '@_T': 'MoveTo', '@_IX': '1', Cell: [{ '@_N': 'X', '@_V': '0' }, { '@_N': 'Y', '@_V': '0' }] },
                        { '@_T': 'LineTo', '@_IX': '2', Cell: [{ '@_N': 'X', '@_V': props.width.toString() }, { '@_N': 'Y', '@_V': '0' }] },
                        { '@_T': 'LineTo', '@_IX': '3', Cell: [{ '@_N': 'X', '@_V': props.width.toString() }, { '@_N': 'Y', '@_V': props.height.toString() }] },
                        { '@_T': 'LineTo', '@_IX': '4', Cell: [{ '@_N': 'X', '@_V': '0' }, { '@_N': 'Y', '@_V': props.height.toString() }] },
                        { '@_T': 'LineTo', '@_IX': '5', Cell: [{ '@_N': 'X', '@_V': '0' }, { '@_N': 'Y', '@_V': '0' }] }
                    ]
                }
            ]
        };

        if (props.fillColor) {
            // Add Fill Section
            newShape.Section.push(createFillSection(props.fillColor));
        }

        if (props.fontColor || props.bold) {
            newShape.Section.push(createCharacterSection({
                bold: props.bold,
                color: props.fontColor
            }));
        }

        if (parentId) {
            // Add to Parent Group
            const parent = allShapes.find((s: any) => s['@_ID'] == parentId);
            if (!parent) {
                throw new Error(`Parent shape ${parentId} not found`);
            }

            // Ensure Parent has Shapes collection
            if (!parent.Shapes) {
                parent.Shapes = { Shape: [] };
            }
            if (!Array.isArray(parent.Shapes.Shape)) {
                parent.Shapes.Shape = parent.Shapes.Shape ? [parent.Shapes.Shape] : [];
            }

            // Mark parent as Group if not already
            if (parent['@_Type'] !== 'Group') {
                parent['@_Type'] = 'Group';
            }

            parent.Shapes.Shape.push(newShape);
        } else {
            // Add to Page
            topLevelShapes.push(newShape);
        }

        const newXml = this.builder.build(parsed);
        this.pkg.updateFile(pagePath, newXml);

        return newId;
    }

    async updateShapeText(pageId: string, shapeId: string, newText: string): Promise<void> {
        const pagePath = this.getPagePath(pageId);
        let content: string;

        try {
            content = this.pkg.getFileText(pagePath);
        } catch {
            // Fallback: This is a simplification.
            // In a real robust app, we should use the page mapping logic to find file path.
            // For now, if exact ID match fails, we throw.
            // A more complex lookup would require re-accessing PageManager or iterating keys.
            throw new Error(`Could not find page file for ID ${pageId}. Expected at ${pagePath}`);
        }

        const parsed = this.parser.parse(content);
        let found = false;

        // Helper to recursively find and update shape
        const findAndUpdate = (shapes: any[]) => {
            for (const shape of shapes) {
                if (shape['@_ID'] == shapeId) {
                    shape.Text = {
                        '#text': newText
                    };
                    found = true;
                    return;
                }
                // If shapes can be nested (Groups), check for sub-shapes - future improvement
                // But typically basic Text is on top level or group level.
            }
        };

        const shapesData = parsed.PageContents?.Shapes?.Shape;
        if (shapesData) {
            const shapesArray = Array.isArray(shapesData) ? shapesData : [shapesData];
            findAndUpdate(shapesArray);
        }

        if (!found) {
            throw new Error(`Shape ${shapeId} not found on page ${pageId}`);
        }

        const newXml = this.builder.build(parsed);
        this.pkg.updateFile(pagePath, newXml);
    }
    async updateShapeStyle(pageId: string, shapeId: string, style: ShapeStyle): Promise<void> {
        const pagePath = this.getPagePath(pageId);
        let content: string;
        try {
            content = this.pkg.getFileText(pagePath);
        } catch {
            throw new Error(`Could not find page file for ID ${pageId}`);
        }

        const parsed = this.parser.parse(content);
        let found = false;

        const findAndUpdate = (shapes: any[]) => {
            for (const shape of shapes) {
                if (shape['@_ID'] == shapeId) {
                    found = true;
                    // Ensure Section array exists
                    if (!shape.Section) {
                        shape.Section = [];
                    } else if (!Array.isArray(shape.Section)) {
                        shape.Section = [shape.Section];
                    }

                    // Update/Add Fill
                    if (style.fillColor) {
                        // Remove existing Fill section if any (simplified: assuming IX=0)
                        shape.Section = shape.Section.filter((s: any) => s['@_N'] !== 'Fill');
                        shape.Section.push(createFillSection(style.fillColor));
                    }

                    // Update/Add Character (Font/Text Style)
                    if (style.fontColor || style.bold !== undefined) {
                        // Remove existing Character section if any
                        shape.Section = shape.Section.filter((s: any) => s['@_N'] !== 'Character');
                        shape.Section.push(createCharacterSection({
                            bold: style.bold,
                            color: style.fontColor
                        }));
                    }
                    return;
                }
            }
        };

        const shapesData = parsed.PageContents?.Shapes?.Shape;
        if (shapesData) {
            const shapesArray = Array.isArray(shapesData) ? shapesData : [shapesData];
            findAndUpdate(shapesArray);
        }

        if (!found) {
            throw new Error(`Shape ${shapeId} not found on page ${pageId}`);
        }

        const newXml = this.builder.build(parsed);
        this.pkg.updateFile(pagePath, newXml);
    }
    async updateShapePosition(pageId: string, shapeId: string, x: number, y: number): Promise<void> {
        const pagePath = this.getPagePath(pageId);
        let content: string;
        try {
            content = this.pkg.getFileText(pagePath);
        } catch {
            throw new Error(`Could not find page file for ID ${pageId}`);
        }

        const parsed = this.parser.parse(content);
        let found = false;

        const findAndUpdate = (shapes: any[]) => {
            for (const shape of shapes) {
                if (shape['@_ID'] == shapeId) {
                    found = true;
                    // Ensure Cell array exists
                    if (!shape.Cell) {
                        shape.Cell = [];
                    } else if (!Array.isArray(shape.Cell)) {
                        shape.Cell = [shape.Cell];
                    }

                    // Helper to update specific cell
                    const updateCell = (name: string, value: string) => {
                        const cell = shape.Cell.find((c: any) => c['@_N'] === name);
                        if (cell) {
                            cell['@_V'] = value;
                        } else {
                            shape.Cell.push({ '@_N': name, '@_V': value });
                        }
                    };

                    updateCell('PinX', x.toString());
                    updateCell('PinY', y.toString());
                    return;
                }
            }
        };

        const shapesData = parsed.PageContents?.Shapes?.Shape;
        if (shapesData) {
            const shapesArray = Array.isArray(shapesData) ? shapesData : [shapesData];
            findAndUpdate(shapesArray);
        }

        if (!found) {
            throw new Error(`Shape ${shapeId} not found on page ${pageId}`);
        }

        const newXml = this.builder.build(parsed);
        this.pkg.updateFile(pagePath, newXml);
    }
}

export interface ShapeStyle {
    fillColor?: string;
    fontColor?: string;
    bold?: boolean;
}
