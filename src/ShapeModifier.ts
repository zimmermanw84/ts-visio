import { XMLParser, XMLBuilder } from 'fast-xml-parser';
import { VisioPackage } from './VisioPackage';
import { RelsManager } from './core/RelsManager';
import { createFillSection, createCharacterSection, createLineSection, createParagraphSection, vertAlignValue, HorzAlign, VertAlign } from './utils/StyleHelpers';
import { RELATIONSHIP_TYPES } from './core/VisioConstants';
import { NewShapeProps, VisioPropType, ConnectorStyle } from './types/VisioTypes';
import type { ShapeData, ShapeHyperlink } from './Shape';
import { ForeignShapeBuilder } from './shapes/ForeignShapeBuilder';
import { ShapeBuilder } from './shapes/ShapeBuilder';
import { ConnectorBuilder } from './shapes/ConnectorBuilder';
import { ContainerBuilder } from './shapes/ContainerBuilder';
import { createXmlParser, createXmlBuilder, buildXml } from './utils/XmlHelper';

export class ShapeModifier {
    // ...
    async addContainer(pageId: string, props: NewShapeProps): Promise<string> {
        const parsed = this.getParsed(pageId);

        if (!parsed.PageContents.Shapes) parsed.PageContents.Shapes = { Shape: [] };
        let topLevelShapes = parsed.PageContents.Shapes.Shape;
        if (!Array.isArray(topLevelShapes)) {
            topLevelShapes = topLevelShapes ? [topLevelShapes] : [];
            parsed.PageContents.Shapes.Shape = topLevelShapes;
        }

        const newId = props.id || this.getNextId(parsed);
        const containerShape = ContainerBuilder.createContainerShape(newId, props);

        topLevelShapes.push(containerShape);
        this.getShapeMap(parsed).set(newId, containerShape);

        this.saveParsed(pageId, parsed);
        return newId;
    }

    async addList(pageId: string, props: NewShapeProps, direction: 'vertical' | 'horizontal' = 'vertical'): Promise<string> {
        const parsed = this.getParsed(pageId);

        if (!parsed.PageContents.Shapes) parsed.PageContents.Shapes = { Shape: [] };
        let topLevelShapes = parsed.PageContents.Shapes.Shape;
        if (!Array.isArray(topLevelShapes)) {
            topLevelShapes = topLevelShapes ? [topLevelShapes] : [];
            parsed.PageContents.Shapes.Shape = topLevelShapes;
        }

        const newId = props.id || this.getNextId(parsed);
        const listShape = ContainerBuilder.createContainerShape(newId, props);
        ContainerBuilder.makeList(listShape, direction);

        topLevelShapes.push(listShape);
        this.getShapeMap(parsed).set(newId, listShape);

        this.saveParsed(pageId, parsed);
        return newId;
    }

    private parser: XMLParser;
    private builder: XMLBuilder;
    private relsManager: RelsManager;
    private pageCache: Map<string, { content: string, parsed: any }> = new Map();
    private dirtyPages: Set<string> = new Set();
    private shapeCache = new WeakMap<object, Map<string, any>>();
    private pagePathRegistry = new Map<string, string>();
    public autoSave: boolean = true;

    /**
     * Register the resolved OPC part path for a page ID.
     * Must be called before any operation on a loaded file to ensure the
     * correct file is targeted rather than the ID-derived fallback name.
     */
    registerPage(pageId: string, xmlPath: string): void {
        this.pagePathRegistry.set(pageId, xmlPath);
    }

    constructor(private pkg: VisioPackage) {
        this.parser = createXmlParser();
        this.builder = createXmlBuilder();
        this.relsManager = new RelsManager(pkg);
    }

    private getPagePath(pageId: string): string {
        return this.pagePathRegistry.get(pageId) ?? `visio/pages/page${pageId}.xml`;
    }

    private getShapeMap(parsed: any): Map<string, any> {
        if (!this.shapeCache.has(parsed)) {
            const map = new Map<string, any>();
            let topLevelShapes = parsed.PageContents.Shapes ? parsed.PageContents.Shapes.Shape : [];
            if (!Array.isArray(topLevelShapes)) {
                topLevelShapes = topLevelShapes ? [topLevelShapes] : [];
            }

            const gather = (shapeList: any[]): void => {
                for (const s of shapeList) {
                    map.set(s['@_ID'], s);
                    if (s.Shapes && s.Shapes.Shape) {
                        const children = Array.isArray(s.Shapes.Shape) ? s.Shapes.Shape : [s.Shapes.Shape];
                        gather(children);
                    }
                }
            };

            gather(topLevelShapes);
            this.shapeCache.set(parsed, map);
        }
        return this.shapeCache.get(parsed)!;
    }

    private getAllShapes(parsed: any): any[] {
        return Array.from(this.getShapeMap(parsed).values());
    }

    private getNextId(parsed: any): string {
        // Updates PageSheet.NextShapeID to prevent ID conflicts.
        // Calculates the next ID from existing shapes and increments the counter.

        const shapeMap = this.getShapeMap(parsed);
        let maxId = 0;
        for (const s of shapeMap.values()) {
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
            // Enforce order: PageSheet, Shapes, Connects
            // Save existing elements
            const shapes = parsed.PageContents.Shapes;
            const connects = parsed.PageContents.Connects;
            const rels = parsed.PageContents.Relationships;

            // Delete to re-insert in order
            if (shapes) delete parsed.PageContents.Shapes;
            if (connects) delete parsed.PageContents.Connects;
            if (rels) delete parsed.PageContents.Relationships;

            // Insert PageSheet first
            parsed.PageContents.PageSheet = { Cell: [] };

            // Restore others in order
            if (shapes) parsed.PageContents.Shapes = shapes;
            if (connects) parsed.PageContents.Connects = connects;
            if (rels) parsed.PageContents.Relationships = rels;
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
        const newXml = buildXml(this.builder, parsed);
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

    async addConnector(pageId: string, fromShapeId: string, toShapeId: string, beginArrow?: string, endArrow?: string, style?: ConnectorStyle): Promise<string> {
        const parsed = this.getParsed(pageId);

        // Ensure Shapes collection exists
        if (!parsed.PageContents.Shapes) {
            parsed.PageContents.Shapes = { Shape: [] };
        }
        if (!Array.isArray(parsed.PageContents.Shapes.Shape)) {
            parsed.PageContents.Shapes.Shape = parsed.PageContents.Shapes.Shape ? [parsed.PageContents.Shapes.Shape] : [];
        }

        const newId = this.getNextId(parsed);
        const shapeHierarchy = ConnectorBuilder.buildShapeHierarchy(parsed);

        // Validate arrow values (Visio supports 0-45)
        const validateArrow = (val?: string): string => {
            if (!val) return '0';
            const num = parseInt(val);
            if (isNaN(num) || num < 0 || num > 45) return '0';
            return val;
        };

        const layout = ConnectorBuilder.calculateConnectorLayout(fromShapeId, toShapeId, shapeHierarchy);
        const connectorShape = ConnectorBuilder.createConnectorShapeObject(newId, layout, validateArrow(beginArrow), validateArrow(endArrow), style);

        const topLevelShapes = parsed.PageContents.Shapes.Shape;
        topLevelShapes.push(connectorShape);
        this.getShapeMap(parsed).set(newId, connectorShape);

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
            const parent = this.getShapeMap(parsed).get(parentId);
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
        this.getShapeMap(parsed).set(newId, newShape);

        this.saveParsed(pageId, parsed);

        return newId;
    }

    async deleteShape(pageId: string, shapeId: string): Promise<void> {
        const parsed = this.getParsed(pageId);

        // 1. Remove shape from the shape tree (handles top-level and nested groups)
        const removed = this.removeShapeFromTree(parsed.PageContents.Shapes, shapeId);
        if (!removed) {
            throw new Error(`Shape ${shapeId} not found on page ${pageId}`);
        }

        // 2. Remove any Connect elements referencing this shape
        if (parsed.PageContents.Connects?.Connect) {
            let connects = parsed.PageContents.Connects.Connect;
            if (!Array.isArray(connects)) connects = [connects];
            const filtered = connects.filter(
                (c: any) => c['@_FromSheet'] !== shapeId && c['@_ToSheet'] !== shapeId
            );
            parsed.PageContents.Connects.Connect = filtered;
        }

        // 3. Remove any Relationship elements referencing this shape (container membership, etc.)
        if (parsed.PageContents.Relationships?.Relationship) {
            let rels = parsed.PageContents.Relationships.Relationship;
            if (!Array.isArray(rels)) rels = [rels];
            parsed.PageContents.Relationships.Relationship = rels.filter(
                (r: any) => r['@_ShapeID'] !== shapeId && r['@_RelatedShapeID'] !== shapeId
            );
        }

        // 4. Invalidate the shape cache so the map is rebuilt on next access
        this.shapeCache.delete(parsed);

        this.saveParsed(pageId, parsed);
    }

    private removeShapeFromTree(shapesContainer: any, shapeId: string): boolean {
        if (!shapesContainer?.Shape) return false;

        let shapes = shapesContainer.Shape;
        if (!Array.isArray(shapes)) shapes = [shapes];

        const idx = shapes.findIndex((s: any) => s['@_ID'] === shapeId);
        if (idx !== -1) {
            shapes.splice(idx, 1);
            shapesContainer.Shape = shapes;
            return true;
        }

        // Recurse into nested group children
        for (const shape of shapes) {
            if (shape.Shapes && this.removeShapeFromTree(shape.Shapes, shapeId)) {
                return true;
            }
        }

        return false;
    }

    async updateShapeText(pageId: string, shapeId: string, newText: string): Promise<void> {
        const parsed = this.getParsed(pageId);
        const shape = this.getShapeMap(parsed).get(shapeId);

        if (!shape) {
            throw new Error(`Shape ${shapeId} not found on page ${pageId}`);
        }

        shape.Text = {
            '#text': newText
        };

        this.saveParsed(pageId, parsed);
    }
    async updateShapeStyle(pageId: string, shapeId: string, style: ShapeStyle): Promise<void> {
        const parsed = this.getParsed(pageId);
        const shape = this.getShapeMap(parsed).get(shapeId);

        if (!shape) {
            throw new Error(`Shape ${shapeId} not found on page ${pageId}`);
        }

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
        if (style.fontColor || style.bold !== undefined || style.fontSize !== undefined || style.fontFamily !== undefined) {
            shape.Section = shape.Section.filter((s: any) => s['@_N'] !== 'Character');
            shape.Section.push(createCharacterSection({
                bold: style.bold,
                color: style.fontColor,
                fontSize: style.fontSize,
                fontFamily: style.fontFamily,
            }));
        }

        // Update/Add Paragraph (Horizontal Alignment)
        if (style.horzAlign !== undefined) {
            shape.Section = shape.Section.filter((s: any) => s['@_N'] !== 'Paragraph');
            shape.Section.push(createParagraphSection(style.horzAlign));
        }

        // Update/Add VerticalAlign (top-level shape Cell)
        if (style.verticalAlign !== undefined) {
            if (!shape.Cell) shape.Cell = [];
            if (!Array.isArray(shape.Cell)) shape.Cell = [shape.Cell];
            const existingCell = shape.Cell.find((c: any) => c['@_N'] === 'VerticalAlign');
            if (existingCell) {
                existingCell['@_V'] = vertAlignValue(style.verticalAlign);
            } else {
                shape.Cell.push({ '@_N': 'VerticalAlign', '@_V': vertAlignValue(style.verticalAlign) });
            }
        }

        this.saveParsed(pageId, parsed);
    }
    async updateShapeDimensions(pageId: string, shapeId: string, w: number, h: number): Promise<void> {
        const parsed = this.getParsed(pageId);
        const shape = this.getShapeMap(parsed).get(shapeId);

        if (!shape) throw new Error(`Shape ${shapeId} not found`);

        // Ensure Cell array
        if (!shape.Cell) shape.Cell = [];
        if (!Array.isArray(shape.Cell)) shape.Cell = [shape.Cell];

        const updateCell = (name: string, val: string) => {
            const cell = shape.Cell.find((c: any) => c['@_N'] === name);
            if (cell) cell['@_V'] = val;
            else shape.Cell.push({ '@_N': name, '@_V': val });
        };

        updateCell('Width', w.toString());
        updateCell('Height', h.toString());

        this.saveParsed(pageId, parsed);
    }

    /**
     * Set the rotation angle of a shape. Degrees are converted to radians
     * for storage in the Angle cell (Visio's native unit).
     */
    async rotateShape(pageId: string, shapeId: string, degrees: number): Promise<void> {
        const parsed = this.getParsed(pageId);
        const shape = this.getShapeMap(parsed).get(shapeId);
        if (!shape) throw new Error(`Shape ${shapeId} not found on page ${pageId}`);

        if (!shape.Cell) shape.Cell = [];
        if (!Array.isArray(shape.Cell)) shape.Cell = [shape.Cell];

        const radians = (degrees * Math.PI) / 180;
        const existing = shape.Cell.find((c: any) => c['@_N'] === 'Angle');
        if (existing) existing['@_V'] = radians.toString();
        else shape.Cell.push({ '@_N': 'Angle', '@_V': radians.toString() });

        this.saveParsed(pageId, parsed);
    }

    /**
     * Set the flip state for a shape along the X or Y axis.
     * FlipX mirrors left-to-right; FlipY mirrors top-to-bottom.
     */
    setShapeFlip(pageId: string, shapeId: string, axis: 'x' | 'y', enabled: boolean): void {
        const parsed = this.getParsed(pageId);
        const shape = this.getShapeMap(parsed).get(shapeId);
        if (!shape) throw new Error(`Shape ${shapeId} not found on page ${pageId}`);

        if (!shape.Cell) shape.Cell = [];
        if (!Array.isArray(shape.Cell)) shape.Cell = [shape.Cell];

        const cellName = axis === 'x' ? 'FlipX' : 'FlipY';
        const value = enabled ? '1' : '0';
        const existing = shape.Cell.find((c: any) => c['@_N'] === cellName);
        if (existing) existing['@_V'] = value;
        else shape.Cell.push({ '@_N': cellName, '@_V': value });

        this.saveParsed(pageId, parsed);
    }

    /**
     * Resize a shape, keeping it centred on its current PinX/PinY.
     * Updates Width, Height, LocPinX, LocPinY, and the cached @_V on any
     * Geometry cells whose @_F formula references Width or Height, so that
     * non-Visio renderers see consistent values.
     */
    async resizeShape(pageId: string, shapeId: string, width: number, height: number): Promise<void> {
        const parsed = this.getParsed(pageId);
        const shape = this.getShapeMap(parsed).get(shapeId);
        if (!shape) throw new Error(`Shape ${shapeId} not found on page ${pageId}`);

        if (!shape.Cell) shape.Cell = [];
        if (!Array.isArray(shape.Cell)) shape.Cell = [shape.Cell];

        const upsert = (name: string, val: string) => {
            const cell = shape.Cell.find((c: any) => c['@_N'] === name);
            if (cell) cell['@_V'] = val;
            else shape.Cell.push({ '@_N': name, '@_V': val });
        };

        upsert('Width',   width.toString());
        upsert('Height',  height.toString());
        upsert('LocPinX', (width  / 2).toString());
        upsert('LocPinY', (height / 2).toString());

        // Keep cached @_V consistent for renderers that don't evaluate formulas
        if (shape.Section) {
            const sections = Array.isArray(shape.Section) ? shape.Section : [shape.Section];
            for (const section of sections) {
                if (section['@_N'] !== 'Geometry' || !section.Row) continue;
                const rows = Array.isArray(section.Row) ? section.Row : [section.Row];
                for (const row of rows) {
                    if (!row.Cell) continue;
                    const cells = Array.isArray(row.Cell) ? row.Cell : [row.Cell];
                    for (const cell of cells) {
                        if (cell['@_F'] === 'Width')  cell['@_V'] = width.toString();
                        if (cell['@_F'] === 'Height') cell['@_V'] = height.toString();
                    }
                }
            }
        }

        this.saveParsed(pageId, parsed);
    }

    async updateShapePosition(pageId: string, shapeId: string, x: number, y: number): Promise<void> {
        const parsed = this.getParsed(pageId);
        const shape = this.getShapeMap(parsed).get(shapeId);

        if (!shape) {
            throw new Error(`Shape ${shapeId} not found on page ${pageId}`);
        }

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

        this.saveParsed(pageId, parsed);
    }
    addPropertyDefinition(pageId: string, shapeId: string, name: string, type: number, options: { label?: string, invisible?: boolean } = {}): void {
        const parsed = this.getParsed(pageId);
        const shape = this.getShapeMap(parsed).get(shapeId);

        if (!shape) {
            throw new Error(`Shape ${shapeId} not found on page ${pageId}`);
        }

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

        this.saveParsed(pageId, parsed);
    }
    private dateToVisioString(date: Date): string {
        // Visio typically accepts ISO 8601 strings for Type 5
        // Example: 2022-01-01T00:00:00
        return date.toISOString().split('.')[0]; // remove milliseconds
    }

    setPropertyValue(pageId: string, shapeId: string, name: string, value: string | number | boolean | Date): void {
        const parsed = this.getParsed(pageId);
        const shape = this.getShapeMap(parsed).get(shapeId);

        if (!shape) {
            throw new Error(`Shape ${shapeId} not found on page ${pageId}`);
        }

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

        this.saveParsed(pageId, parsed);
    }
    getShapeGeometry(pageId: string, shapeId: string): { x: number, y: number, width: number, height: number } {
        const parsed = this.getParsed(pageId);
        const shape = this.getShapeMap(parsed).get(shapeId);

        if (!shape) throw new Error(`Shape ${shapeId} not found`);

        const getCellVal = (name: string) => {
            // Ensure Cell is array
            if (!shape.Cell) return 0;
            const cells = Array.isArray(shape.Cell) ? shape.Cell : [shape.Cell];
            const c = cells.find((cell: any) => cell['@_N'] === name);
            return c ? Number(c['@_V']) : 0;
        };

        return {
            x: getCellVal('PinX'),
            y: getCellVal('PinY'),
            width: getCellVal('Width'),
            height: getCellVal('Height')
        };
    }

    async addRelationship(pageId: string, shapeId: string, relatedShapeId: string, type: string): Promise<void> {
        const parsed = this.getParsed(pageId);

        // Ensure Relationships collection exists in PageContents
        if (!parsed.PageContents.Relationships) {
            parsed.PageContents.Relationships = { Relationship: [] };
        }
        // Ensure Relationship is an array (Standard robustness pattern)
        if (!Array.isArray(parsed.PageContents.Relationships.Relationship)) {
            parsed.PageContents.Relationships.Relationship = parsed.PageContents.Relationships.Relationship
                ? [parsed.PageContents.Relationships.Relationship]
                : [];
        }

        const relationships = parsed.PageContents.Relationships.Relationship;

        // Check definition: Type, ShapeID (Container), RelatedShapeID (Member)
        // Avoid duplicates?
        const exists = relationships.find((r: any) =>
            r['@_Type'] === type &&
            r['@_ShapeID'] === shapeId &&
            r['@_RelatedShapeID'] === relatedShapeId
        );

        if (!exists) {
            relationships.push({
                '@_Type': type,
                '@_ShapeID': shapeId,
                '@_RelatedShapeID': relatedShapeId
            });
            this.saveParsed(pageId, parsed);
        }
    }

    getContainerMembers(pageId: string, containerId: string): string[] {
        const parsed = this.getParsed(pageId);
        const rels = parsed.PageContents?.Relationships?.Relationship;
        if (!rels) return [];

        const relsArray = Array.isArray(rels) ? rels : [rels];

        return relsArray
            .filter((r: any) => r['@_Type'] === 'Container' && r['@_ShapeID'] === containerId)
            .map((r: any) => r['@_RelatedShapeID']);
    }

    async reorderShape(pageId: string, shapeId: string, position: 'front' | 'back'): Promise<void> {
        const parsed = this.getParsed(pageId);
        const shapesContainer = parsed.PageContents?.Shapes;
        if (!shapesContainer || !shapesContainer.Shape) return;

        let shapes = shapesContainer.Shape;
        if (!Array.isArray(shapes)) shapes = [shapes];

        const idx = shapes.findIndex((s: any) => s['@_ID'] == shapeId);
        if (idx === -1) return;

        const shape = shapes[idx];
        shapes.splice(idx, 1); // Remove

        if (position === 'back') {
            shapes.unshift(shape); // Add to start (Back of Z-Order)
        } else {
            shapes.push(shape); // Add to end (Front of Z-Order)
        }

        // Update array in object
        shapesContainer.Shape = shapes;
        this.saveParsed(pageId, parsed);
    }

    async addListItem(pageId: string, listId: string, itemId: string): Promise<void> {
        // 1. Get List Properties (Direction, Spacing)
        const parsed = this.getParsed(pageId);
        const listShape = this.getShapeMap(parsed).get(listId);
        if (!listShape) throw new Error(`List ${listId} not found`);

        const getUserVal = (name: string, def: string) => {
            if (!listShape.Section) return def;
            const userSec = listShape.Section.find((s: any) => s['@_N'] === 'User');
            if (!userSec || !userSec.Row) return def;
            const rows = Array.isArray(userSec.Row) ? userSec.Row : [userSec.Row];
            const row = rows.find((r: any) => r['@_N'] === name);
            if (!row || !row.Cell) return def;
            // Value cell
            const valCell = Array.isArray(row.Cell) ? row.Cell.find((c: any) => c['@_N'] === 'Value') : row.Cell;
            return valCell ? valCell['@_V'] : def;
        };

        const direction = parseInt(getUserVal('msvSDListDirection', '1')); // 1=Vert, 0=Horiz
        const spacing = parseFloat(getUserVal('msvSDListSpacing', '0.125').replace(/[^0-9.]/g, '')); // Crude parse if unit included

        // 2. Determine Position
        const memberIds = this.getContainerMembers(pageId, listId);
        const itemGeo = this.getShapeGeometry(pageId, itemId);
        const listGeo = this.getShapeGeometry(pageId, listId);

        let newX = listGeo.x;
        let newY = listGeo.y;

        if (memberIds.length === 0) {
            // First Item: Place at Top/Left of Container (with some internal margin/padding)
            // For simplicity, center on Container center or rely on resizeToFit to adjust container AROUND it later.
            // Let's place it at current container PinX/PinY
            newX = listGeo.x;
            newY = listGeo.y;
        } else {
            const lastId = memberIds[memberIds.length - 1];
            const lastGeo = this.getShapeGeometry(pageId, lastId);

            if (direction === 1) { // Vertical (Stack Down)
                // Last Bottom - Spacing - ItemHalfHeight
                const lastBottom = lastGeo.y - (lastGeo.height / 2);
                newY = lastBottom - spacing - (itemGeo.height / 2);
                newX = lastGeo.x; // Align Centers
            } else { // Horizontal (Stack Right)
                // Last Right + Spacing + ItemHalfWidth
                const lastRight = lastGeo.x + (lastGeo.width / 2);
                newX = lastRight + spacing + (itemGeo.width / 2);
                newY = lastGeo.y; // Align Centers
            }
        }

        // 3. Update Item Position
        await this.updateShapePosition(pageId, itemId, newX, newY);

        // 4. Add Relationship
        await this.addRelationship(pageId, listId, itemId, 'Container');

        // 5. Resize List Container
        await this.resizeContainerToFit(pageId, listId, 0.25);
    }

    async resizeContainerToFit(pageId: string, containerId: string, padding: number = 0.25): Promise<void> {
        const memberIds = this.getContainerMembers(pageId, containerId);
        if (memberIds.length === 0) return;

        // Calculate Bounding Box
        let minX = Infinity, minY = Infinity;
        let maxX = -Infinity, maxY = -Infinity;

        for (const mid of memberIds) {
            const geo = this.getShapeGeometry(pageId, mid);
            // Visio PinX/PinY is center. Bounding box needs Left/Bottom/Right/Top
            const left = geo.x - (geo.width / 2);
            const right = geo.x + (geo.width / 2);
            const bottom = geo.y - (geo.height / 2);
            const top = geo.y + (geo.height / 2);

            if (left < minX) minX = left;
            if (right > maxX) maxX = right;
            if (bottom < minY) minY = bottom;
            if (top > maxY) maxY = top;
        }

        // Apply Padding
        minX -= padding;
        maxX += padding;
        minY -= padding;
        maxY += padding;

        const newWidth = maxX - minX;
        const newHeight = maxY - minY;
        const newPinX = minX + (newWidth / 2);
        const newPinY = minY + (newHeight / 2);

        // Update Geometry
        await this.updateShapePosition(pageId, containerId, newPinX, newPinY);
        await this.updateShapeDimensions(pageId, containerId, newWidth, newHeight);

        // Update Z-Order (Send to Back)
        await this.reorderShape(pageId, containerId, 'back');
    }

    async addHyperlink(pageId: string, shapeId: string, details: { address?: string, subAddress?: string, description?: string }): Promise<void> {
        const parsed = this.getParsed(pageId);
        const shape = this.getShapeMap(parsed).get(shapeId);

        if (!shape) throw new Error(`Shape ${shapeId} not found`);

        // Ensure Section array
        if (!shape.Section) shape.Section = [];
        if (!Array.isArray(shape.Section)) shape.Section = [shape.Section];

        // Find or Create Hyperlink Section
        let linkSection = shape.Section.find((s: any) => s['@_N'] === 'Hyperlink');
        if (!linkSection) {
            linkSection = { '@_N': 'Hyperlink', Row: [] };
            shape.Section.push(linkSection);
        }

        // Ensure Row array
        if (!linkSection.Row) linkSection.Row = [];
        if (!Array.isArray(linkSection.Row)) linkSection.Row = [linkSection.Row];

        // Determine next Row ID (Hyperlink.Row_1, Hyperlink.Row_2, etc.)
        const nextIdx = linkSection.Row.length + 1;
        const rowName = `Hyperlink.Row_${nextIdx}`;

        const newRow: any = {
            '@_N': rowName,
            Cell: []
        };

        if (details.address !== undefined) {
            // XMLBuilder handles XML escaping automatically
            newRow.Cell.push({ '@_N': 'Address', '@_V': details.address });
        }

        if (details.subAddress !== undefined) {
            newRow.Cell.push({ '@_N': 'SubAddress', '@_V': details.subAddress });
        }

        if (details.description !== undefined) {
            newRow.Cell.push({ '@_N': 'Description', '@_V': details.description });
        }

        // Default NewWindow to 0 (false)
        newRow.Cell.push({ '@_N': 'NewWindow', '@_V': '0' });

        linkSection.Row.push(newRow);

        this.saveParsed(pageId, parsed);
    }

    async addLayer(pageId: string, name: string, options: { visible?: boolean, lock?: boolean, print?: boolean } = {}): Promise<{ name: string, index: number }> {
        const parsed = this.getParsed(pageId);

        // Ensure PageSheet
        this.ensurePageSheet(parsed);
        const pageSheet = parsed.PageContents.PageSheet;

        // Ensure Section array
        if (!pageSheet.Section) pageSheet.Section = [];
        if (!Array.isArray(pageSheet.Section)) pageSheet.Section = [pageSheet.Section];

        // Find or Create Layer Section
        let layerSection = pageSheet.Section.find((s: any) => s['@_N'] === 'Layer');
        if (!layerSection) {
            layerSection = { '@_N': 'Layer', Row: [] };
            pageSheet.Section.push(layerSection);
        }

        // Ensure Row array
        if (!layerSection.Row) layerSection.Row = [];
        if (!Array.isArray(layerSection.Row)) layerSection.Row = [layerSection.Row];

        // Verify name uniqueness (Visio allows duplicates but it's bad practice)
        // For simplicity, we create a new layer even if name matches.

        // Determine Index (IX)
        let maxIx = -1;
        for (const row of layerSection.Row) {
            const ix = parseInt(row['@_IX']);
            if (!isNaN(ix) && ix > maxIx) maxIx = ix;
        }
        const newIndex = maxIx + 1;

        const newRow = {
            '@_IX': newIndex.toString(),
            Cell: [
                { '@_N': 'Name', '@_V': name },
                { '@_N': 'Visible', '@_V': (options.visible ?? true) ? '1' : '0' },
                { '@_N': 'Lock', '@_V': (options.lock ?? false) ? '1' : '0' },
                { '@_N': 'Print', '@_V': (options.print ?? true) ? '1' : '0' }
            ]
        };

        layerSection.Row.push(newRow);
        this.saveParsed(pageId, parsed);

        return { name, index: newIndex };
    }

    async assignLayer(pageId: string, shapeId: string, layerIndex: number): Promise<void> {
        const parsed = this.getParsed(pageId);
        const shape = this.getShapeMap(parsed).get(shapeId);

        if (!shape) throw new Error(`Shape ${shapeId} not found`);

        // Ensure Section array
        if (!shape.Section) shape.Section = [];
        if (!Array.isArray(shape.Section)) shape.Section = [shape.Section];

        // Find or Create LayerMem Section
        let memSection = shape.Section.find((s: any) => s['@_N'] === 'LayerMem');
        if (!memSection) {
            memSection = { '@_N': 'LayerMem', Row: [] };
            shape.Section.push(memSection);
        }

        // Ensure Row array
        if (!memSection.Row) memSection.Row = [];
        if (!Array.isArray(memSection.Row)) memSection.Row = [memSection.Row];

        // Ensure Row exists (LayerMem usually has 1 row)
        if (memSection.Row.length === 0) {
            memSection.Row.push({ Cell: [] });
        }
        const row = memSection.Row[0];

        // Ensure Cell array
        if (!row.Cell) row.Cell = [];
        if (!Array.isArray(row.Cell)) row.Cell = [row.Cell];

        // Find LayerMember Cell
        let cell = row.Cell.find((c: any) => c['@_N'] === 'LayerMember');
        if (!cell) {
            cell = { '@_N': 'LayerMember', '@_V': '' };
            row.Cell.push(cell);
        }

        // Update Value
        const currentVal = cell['@_V'] || '';
        const indices = currentVal.split(';').filter((s: string) => s.length > 0);
        const idxStr = layerIndex.toString();

        if (!indices.includes(idxStr)) {
            indices.push(idxStr);
            // Sort optionally? Visio doesn't strictly require sorting but it's cleaner.
            // Let's keep insertion order or sort numeric. Visio usually semicolon separates.
            cell['@_V'] = indices.join(';');
            this.saveParsed(pageId, parsed);
        }
    }

    async updateLayerProperty(pageId: string, layerIndex: number, propName: string, value: string): Promise<void> {
        const parsed = this.getParsed(pageId);

        this.ensurePageSheet(parsed);
        const pageSheet = parsed.PageContents.PageSheet;

        // Find Layer Section
        if (!pageSheet.Section) return;
        const sections = Array.isArray(pageSheet.Section) ? pageSheet.Section : [pageSheet.Section];
        const layerSection = sections.find((s: any) => s['@_N'] === 'Layer');
        if (!layerSection || !layerSection.Row) return;

        const rows = Array.isArray(layerSection.Row) ? layerSection.Row : [layerSection.Row];
        const row = rows.find((r: any) => r['@_IX'] == layerIndex.toString());
        if (!row) return;

        // Ensure Cell array
        if (!row.Cell) row.Cell = [];
        if (!Array.isArray(row.Cell)) row.Cell = [row.Cell];

        // Find or Create Cell
        let cell = row.Cell.find((c: any) => c['@_N'] === propName);
        if (!cell) {
            cell = { '@_N': propName, '@_V': value };
            row.Cell.push(cell);
        } else {
            cell['@_V'] = value;
        }

        this.saveParsed(pageId, parsed);
    }

    /**
     * Read back all custom property (shape data) entries for a shape.
     * Returns a map of property key → ShapeData, with values coerced to
     * the declared type (Number, Boolean, Date, or String).
     */
    getShapeProperties(pageId: string, shapeId: string): Record<string, ShapeData> {
        const parsed = this.getParsed(pageId);
        const shape = this.getShapeMap(parsed).get(shapeId);
        if (!shape) throw new Error(`Shape ${shapeId} not found on page ${pageId}`);

        const result: Record<string, ShapeData> = {};
        if (!shape.Section) return result;

        const sections = Array.isArray(shape.Section) ? shape.Section : [shape.Section];
        const propSection = sections.find((s: any) => s['@_N'] === 'Property');
        if (!propSection?.Row) return result;

        const rows = Array.isArray(propSection.Row) ? propSection.Row : [propSection.Row];
        for (const row of rows) {
            // Row names are "Prop.KeyName" — strip the prefix to recover the user-facing key
            const rawKey: string = row['@_N'] ?? '';
            const key = rawKey.startsWith('Prop.') ? rawKey.slice(5) : rawKey;
            if (!key) continue;

            const cells: any[] = Array.isArray(row.Cell) ? row.Cell : (row.Cell ? [row.Cell] : []);
            const getCell = (name: string): string | undefined =>
                cells.find((c: any) => c['@_N'] === name)?.['@_V'];

            const rawValue = getCell('Value') ?? '';
            const type = parseInt(getCell('Type') ?? '0') as VisioPropType;
            const label = getCell('Label');
            const hidden = getCell('Invisible') === '1';

            let value: string | number | boolean | Date;
            switch (type) {
                case VisioPropType.Number:
                case VisioPropType.Currency:
                case VisioPropType.Duration:
                    value = parseFloat(rawValue) || 0;
                    break;
                case VisioPropType.Boolean:
                    value = rawValue === '1' || rawValue.toLowerCase() === 'true';
                    break;
                case VisioPropType.Date:
                    value = new Date(rawValue);
                    break;
                default:
                    value = rawValue;
            }

            result[key] = { value, label, hidden, type };
        }

        return result;
    }

    /**
     * Set the page canvas size. Writes PageWidth / PageHeight into the PageSheet
     * and sets DrawingSizeType=0 (Custom) so Visio does not override the values.
     */
    setPageSize(pageId: string, width: number, height: number): void {
        if (width <= 0 || height <= 0) throw new Error('Page dimensions must be positive');
        const parsed = this.getParsed(pageId);
        this.ensurePageSheet(parsed);
        const ps = parsed.PageContents.PageSheet;

        const upsert = (name: string, value: string) => {
            const existing = ps.Cell.find((c: any) => c['@_N'] === name);
            if (existing) existing['@_V'] = value;
            else ps.Cell.push({ '@_N': name, '@_V': value });
        };

        upsert('PageWidth',       width.toString());
        upsert('PageHeight',      height.toString());
        upsert('DrawingSizeType', '0');
        upsert('PageDrawSizeType', '0');

        this.saveParsed(pageId, parsed);
    }

    /**
     * Read the current page canvas dimensions.
     * Returns 8.5 × 11 (US Letter) if no PageSheet cells are present.
     */
    getPageDimensions(pageId: string): { width: number; height: number } {
        const parsed = this.getParsed(pageId);
        const ps = parsed.PageContents?.PageSheet;
        if (!ps?.Cell) return { width: 8.5, height: 11 };
        const cells: any[] = Array.isArray(ps.Cell) ? ps.Cell : [ps.Cell];
        const getVal = (name: string, def: number): number => {
            const c = cells.find((cell: any) => cell['@_N'] === name);
            return c ? parseFloat(c['@_V']) : def;
        };
        return { width: getVal('PageWidth', 8.5), height: getVal('PageHeight', 11) };
    }

    /**
     * Read back all hyperlinks attached to a shape.
     */
    getShapeHyperlinks(pageId: string, shapeId: string): ShapeHyperlink[] {
        const parsed = this.getParsed(pageId);
        const shape = this.getShapeMap(parsed).get(shapeId);
        if (!shape) throw new Error(`Shape ${shapeId} not found on page ${pageId}`);

        const result: ShapeHyperlink[] = [];
        if (!shape.Section) return result;

        const sections = Array.isArray(shape.Section) ? shape.Section : [shape.Section];
        const linkSection = sections.find((s: any) => s['@_N'] === 'Hyperlink');
        if (!linkSection?.Row) return result;

        const rows = Array.isArray(linkSection.Row) ? linkSection.Row : [linkSection.Row];
        for (const row of rows) {
            const cells: any[] = Array.isArray(row.Cell) ? row.Cell : (row.Cell ? [row.Cell] : []);
            const getCell = (name: string): string | undefined =>
                cells.find((c: any) => c['@_N'] === name)?.['@_V'];

            result.push({
                address: getCell('Address'),
                subAddress: getCell('SubAddress'),
                description: getCell('Description'),
                newWindow: getCell('NewWindow') === '1',
            });
        }

        return result;
    }

    /**
     * Read back the layer indices a shape is assigned to.
     * Returns an empty array if the shape has no layer assignment.
     */
    getShapeLayerIndices(pageId: string, shapeId: string): number[] {
        const parsed = this.getParsed(pageId);
        const shape = this.getShapeMap(parsed).get(shapeId);
        if (!shape) throw new Error(`Shape ${shapeId} not found on page ${pageId}`);

        if (!shape.Section) return [];
        const sections = Array.isArray(shape.Section) ? shape.Section : [shape.Section];
        const memSection = sections.find((s: any) => s['@_N'] === 'LayerMem');
        if (!memSection?.Row) return [];

        const rows = Array.isArray(memSection.Row) ? memSection.Row : [memSection.Row];
        const row = rows[0];
        if (!row) return [];

        const cells: any[] = Array.isArray(row.Cell) ? row.Cell : (row.Cell ? [row.Cell] : []);
        const memberCell = cells.find((c: any) => c['@_N'] === 'LayerMember');
        if (!memberCell?.['@_V']) return [];

        return (memberCell['@_V'] as string)
            .split(';')
            .filter((s: string) => s.length > 0)
            .map((s: string) => parseInt(s))
            .filter((n: number) => !isNaN(n));
    }
}


export interface ShapeStyle {
    fillColor?: string;
    fontColor?: string;
    bold?: boolean;
    /** Font size in points (e.g. 14 for 14pt). */
    fontSize?: number;
    /** Font family name (e.g. "Arial"). */
    fontFamily?: string;
    /** Horizontal text alignment. */
    horzAlign?: HorzAlign;
    /** Vertical text alignment. */
    verticalAlign?: VertAlign;
}
