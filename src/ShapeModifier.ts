import { XMLParser, XMLBuilder } from 'fast-xml-parser';
import { VisioPackage } from './VisioPackage';
import { RelsManager } from './core/RelsManager';
import { createFillSection, createCharacterSection, createLineSection } from './utils/StyleHelpers';
import { RELATIONSHIP_TYPES } from './core/VisioConstants';
import { NewShapeProps } from './types/VisioTypes';
import { ForeignShapeBuilder } from './shapes/ForeignShapeBuilder';
import { ShapeBuilder } from './shapes/ShapeBuilder';
import { ConnectorBuilder } from './shapes/ConnectorBuilder';
import { ContainerBuilder } from './shapes/ContainerBuilder';

export class ShapeModifier {
    // ...
    async addContainer(pageId: string, props: NewShapeProps): Promise<string> {

        const pagePath = this.getPagePath(pageId);
        let content = this.pkg.getFileText(pagePath);
        const parsed = this.parser.parse(content);

        // Ensure Shapes container...
        if (!parsed.PageContents.Shapes) parsed.PageContents.Shapes = { Shape: [] };
        let topLevelShapes = parsed.PageContents.Shapes.Shape;
        if (!Array.isArray(topLevelShapes)) {
            topLevelShapes = topLevelShapes ? [topLevelShapes] : [];
            parsed.PageContents.Shapes.Shape = topLevelShapes;
        }

        let newId = props.id || this.getNextId(parsed);
        const containerShape = ContainerBuilder.createContainerShape(newId, props);

        topLevelShapes.push(containerShape);

        const newXml = this.builder.build(parsed);
        this.pkg.updateFile(pagePath, newXml);
        return newId;
    }
    private parser: XMLParser;
    private builder: XMLBuilder;
    private relsManager: RelsManager;
    private pageCache: Map<string, { content: string, parsed: any }> = new Map();
    private dirtyPages: Set<string> = new Set();
    public autoSave: boolean = true;

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
        this.relsManager = new RelsManager(pkg);
    }

    private getPagePath(pageId: string): string {
        return `visio/pages/page${pageId}.xml`;
    }

    private getAllShapes(parsed: any): any[] {
        let topLevelShapes = parsed.PageContents.Shapes ? parsed.PageContents.Shapes.Shape : [];
        if (!Array.isArray(topLevelShapes)) {
            topLevelShapes = topLevelShapes ? [topLevelShapes] : [];
        }

        const all: any[] = [];
        const gather = (shapeList: any[]): void => {
            for (const s of shapeList) {
                all.push(s);
                if (s.Shapes && s.Shapes.Shape) {
                    const children = Array.isArray(s.Shapes.Shape) ? s.Shapes.Shape : [s.Shapes.Shape];
                    gather(children);
                }
            }
        };

        gather(topLevelShapes);
        return all;
    }

    private getNextId(parsed: any): string {
        // Updates PageSheet.NextShapeID to prevent ID conflicts.
        // Calculates the next ID from existing shapes and increments the counter.

        const allShapes = this.getAllShapes(parsed);
        let maxId = 0;
        for (const s of allShapes) {
            const id = parseInt(s['@_ID']);
            if (!isNaN(id) && id > maxId) maxId = id;
        }
        const nextId = maxId + 1;

        // Update PageSheet so that NextShapeID always points to the next available shape ID (store nextId + 1)
        this.updateNextShapeId(parsed, nextId + 1);

        return nextId.toString();
    }

    private ensurePageSheet(parsed: any) {
        if (!parsed.PageContents.PageSheet) {
            parsed.PageContents.PageSheet = { Cell: [] };
        }
        if (!Array.isArray(parsed.PageContents.PageSheet.Cell)) {
            parsed.PageContents.PageSheet.Cell = parsed.PageContents.PageSheet.Cell ? [parsed.PageContents.PageSheet.Cell] : [];
        }
    }

    private updateNextShapeId(parsed: any, nextVal: number) {
        this.ensurePageSheet(parsed);
        const cells = parsed.PageContents.PageSheet.Cell;
        const cell = cells.find((c: any) => c['@_N'] === 'NextShapeID');
        if (cell) {
            cell['@_V'] = nextVal.toString();
        } else {
            cells.push({ '@_N': 'NextShapeID', '@_V': nextVal.toString() });
        }
    }

    private getParsed(pageId: string): any {
        const pagePath = this.getPagePath(pageId);
        let content: string;
        try {
            content = this.pkg.getFileText(pagePath);
        } catch {
            throw new Error(`Could not find page file for ID ${pageId}. Expected at ${pagePath}`);
        }

        const cached = this.pageCache.get(pagePath);
        if (cached && cached.content === content) {
            return cached.parsed;
        }

        const parsed = this.parser.parse(content);
        this.pageCache.set(pagePath, { content, parsed });
        return parsed;
    }

    private saveParsed(pageId: string, parsed: any): void {
        const pagePath = this.getPagePath(pageId);

        if (!this.autoSave) {
            this.dirtyPages.add(pagePath);
            return;
        }

        this.performSave(pagePath, parsed);
    }

    private performSave(pagePath: string, parsed: any): void {
        const newXml = this.builder.build(parsed);
        this.pkg.updateFile(pagePath, newXml);
        this.pageCache.set(pagePath, { content: newXml, parsed });
    }

    public flush(): void {
        for (const pagePath of this.dirtyPages) {
            const cached = this.pageCache.get(pagePath);
            if (cached && cached.parsed) {
                this.performSave(pagePath, cached.parsed);
            }
        }
        this.dirtyPages.clear();
    }

    async addConnector(pageId: string, fromShapeId: string, toShapeId: string, beginArrow?: string, endArrow?: string): Promise<string> {
        const parsed = this.getParsed(pageId);

        // Ensure Shapes collection exists
        if (!parsed.PageContents.Shapes) {
            parsed.PageContents.Shapes = { Shape: [] };
        }
        if (!Array.isArray(parsed.PageContents.Shapes.Shape)) {
            parsed.PageContents.Shapes.Shape = [parsed.PageContents.Shapes.Shape];
        }

        const newId = this.getNextId(parsed);
        const shapeHierarchy = ConnectorBuilder.buildShapeHierarchy(parsed);

        const layout = ConnectorBuilder.calculateConnectorLayout(fromShapeId, toShapeId, shapeHierarchy);
        const connectorShape = ConnectorBuilder.createConnectorShapeObject(newId, layout, beginArrow, endArrow);

        const topLevelShapes = parsed.PageContents.Shapes.Shape;
        topLevelShapes.push(connectorShape);

        ConnectorBuilder.addConnectorToConnects(parsed, newId, fromShapeId, toShapeId);

        this.saveParsed(pageId, parsed);

        return newId;
    }




    async addShape(pageId: string, props: NewShapeProps, parentId?: string): Promise<string> {
        const parsed = this.getParsed(pageId);

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

        let newShape: any;

        if (props.type === 'Foreign' && props.imgRelId) {
            newShape = ForeignShapeBuilder.createImageShapeObject(newId, props.imgRelId, props);
            // Text for foreign shapes? Usually none, but we can support it.
            if (props.text !== undefined && props.text !== null) {
                newShape.Text = { '#text': props.text };
            }
        } else {
            // Standard Shape creation logic
            newShape = ShapeBuilder.createStandardShape(newId, props);

            if (props.masterId) {
                // Phase 3: Ensure Relationship
                await this.relsManager.ensureRelationship(
                    `visio/pages/page${pageId}.xml`,
                    '../masters/masters.xml',
                    RELATIONSHIP_TYPES.MASTERS
                );
            }
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

        this.saveParsed(pageId, parsed);

        return newId;
    }

    async updateShapeText(pageId: string, shapeId: string, newText: string): Promise<void> {
        const parsed = this.getParsed(pageId);
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

        this.saveParsed(pageId, parsed);
    }
    async updateShapeStyle(pageId: string, shapeId: string, style: ShapeStyle): Promise<void> {
        const parsed = this.getParsed(pageId);
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

        this.saveParsed(pageId, parsed);
    }
    async updateShapePosition(pageId: string, shapeId: string, x: number, y: number): Promise<void> {
        const parsed = this.getParsed(pageId);
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

        this.saveParsed(pageId, parsed);
    }
    addPropertyDefinition(pageId: string, shapeId: string, name: string, type: number, options: { label?: string, invisible?: boolean } = {}): void {
        const parsed = this.getParsed(pageId);
        let found = false;

        const findAndUpdate = (shapes: any[]) => {
            for (const shape of shapes) {
                if (shape['@_ID'] == shapeId) {
                    found = true;
                    // Ensure Section array exists
                    if (!shape.Section) shape.Section = [];
                    if (!Array.isArray(shape.Section)) shape.Section = [shape.Section];

                    // Find or Create Property Section
                    let propSection = shape.Section.find((s: any) => s['@_N'] === 'Property');
                    if (!propSection) {
                        propSection = { '@_N': 'Property', Row: [] };
                        shape.Section.push(propSection);
                    }

                    // Ensure Row array exists
                    if (!propSection.Row) propSection.Row = [];
                    if (!Array.isArray(propSection.Row)) propSection.Row = [propSection.Row];

                    // Check if property already exists
                    const existingRow = propSection.Row.find((r: any) => r['@_N'] === `Prop.${name}`);
                    if (existingRow) {
                        // Update existing Definition
                        const updateCell = (n: string, v: string) => {
                            let c = existingRow.Cell.find((x: any) => x['@_N'] === n);
                            if (c) c['@_V'] = v;
                            else existingRow.Cell.push({ '@_N': n, '@_V': v });
                        };
                        if (options.label !== undefined) updateCell('Label', options.label);
                        updateCell('Type', type.toString());
                        if (options.invisible !== undefined) updateCell('Invisible', options.invisible ? '1' : '0');
                    } else {
                        // Create New Row
                        propSection.Row.push({
                            '@_N': `Prop.${name}`,
                            Cell: [
                                { '@_N': 'Label', '@_V': options.label || name }, // Default label to name
                                { '@_N': 'Type', '@_V': type.toString() },
                                { '@_N': 'Invisible', '@_V': options.invisible ? '1' : '0' },
                                { '@_N': 'Value', '@_V': '0' } // Initialize with default
                            ]
                        });
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

        this.saveParsed(pageId, parsed);
    }
    private dateToVisioString(date: Date): string {
        // Visio typically accepts ISO 8601 strings for Type 5
        // Example: 2022-01-01T00:00:00
        return date.toISOString().split('.')[0]; // remove milliseconds
    }

    setPropertyValue(pageId: string, shapeId: string, name: string, value: string | number | boolean | Date): void {
        const parsed = this.getParsed(pageId);
        let found = false;

        const findAndUpdate = (shapes: any[]) => {
            for (const shape of shapes) {
                if (shape['@_ID'] == shapeId) {
                    found = true;
                    // Ensure Section array exists
                    const sections = shape.Section ? (Array.isArray(shape.Section) ? shape.Section : [shape.Section]) : [];
                    const propSection = sections.find((s: any) => s['@_N'] === 'Property');

                    if (!propSection) {
                        throw new Error(`Property definition 'Prop.${name}' does not exist on shape ${shapeId}. Call addPropertyDefinition first.`);
                    }

                    const rows = propSection.Row ? (Array.isArray(propSection.Row) ? propSection.Row : [propSection.Row]) : [];
                    const row = rows.find((r: any) => r['@_N'] === `Prop.${name}`);

                    if (!row) {
                        throw new Error(`Property definition 'Prop.${name}' does not exist on shape ${shapeId}. Call addPropertyDefinition first.`);
                    }

                    // Determine Visio Value String
                    let visioValue = '';
                    if (value instanceof Date) {
                        visioValue = this.dateToVisioString(value);
                    } else if (typeof value === 'boolean') {
                        visioValue = value ? '1' : '0'; // Should boolean be V='TRUE' or 1? Standard practice is often 1/0 or TRUE/FALSE. Cells are formulaic.
                        // However, if the Type is 3 (Boolean), Visio often expects 0/1 or TRUE/FALSE.
                        // Let's stick to '1'/'0' for safety in formulas if generic.
                    } else {
                        visioValue = value.toString();
                    }

                    // Update or Add Value Cell
                    // Note: If Type is String (0), V="String". If Number (2), V="123".
                    // Visio often puts string values in formulae as "String", but in XML V attribute it's raw text?
                    // Actually, for String props, V usually contains the string.

                    let valCell = row.Cell.find((c: any) => c['@_N'] === 'Value');
                    if (valCell) {
                        valCell['@_V'] = visioValue;
                    } else {
                        row.Cell.push({ '@_N': 'Value', '@_V': visioValue });
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

        this.saveParsed(pageId, parsed);
    }
}

export interface ShapeStyle {
    fillColor?: string;
    fontColor?: string;
    bold?: boolean;
}
