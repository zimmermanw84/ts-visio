import { XMLParser, XMLBuilder } from 'fast-xml-parser';
import { VisioPackage } from './VisioPackage';
import { RelsManager } from './core/RelsManager';
import { createFillSection, createCharacterSection, createLineSection, createParagraphSection, createTextBlockSection, vertAlignValue, HorzAlign, VertAlign } from './utils/StyleHelpers';
import { RELATIONSHIP_TYPES, SECTION_NAMES, SHAPE_TYPES, STRUCT_RELATIONSHIP_TYPES, LENGTH_UNIT_TO_VISIO, VISIO_TO_LENGTH_UNIT } from './core/VisioConstants';
import { NewShapeProps, VisioPropType, ConnectorStyle, ConnectionTarget, ConnectionPointDef, DrawingScaleInfo, LengthUnit } from './types/VisioTypes';
import { ConnectionPointBuilder } from './shapes/ConnectionPointBuilder';
import { ShapeReader } from './ShapeReader';
import type { ShapeData, ShapeHyperlink } from './Shape';
import type { VisioShape } from './types/VisioTypes';
import { ForeignShapeBuilder } from './shapes/ForeignShapeBuilder';
import { ShapeBuilder } from './shapes/ShapeBuilder';
import { ConnectorBuilder } from './shapes/ConnectorBuilder';
import { ContainerBuilder } from './shapes/ContainerBuilder';
import { createXmlParser, createXmlBuilder, buildXml } from './utils/XmlHelper';

export class ShapeModifier {
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
            // Enforce element order: PageSheet must precede Shapes and Connects in the XML.
            const shapes = parsed.PageContents.Shapes;
            const connects = parsed.PageContents.Connects;
            const rels = parsed.PageContents.Relationships;

            if (shapes) delete parsed.PageContents.Shapes;
            if (connects) delete parsed.PageContents.Connects;
            if (rels) delete parsed.PageContents.Relationships;

            parsed.PageContents.PageSheet = { Cell: [] };

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

    async addConnector(
        pageId: string,
        fromShapeId: string,
        toShapeId: string,
        beginArrow?: string,
        endArrow?: string,
        style?: ConnectorStyle,
        fromPort?: ConnectionTarget,
        toPort?: ConnectionTarget,
    ): Promise<string> {
        const parsed = this.getParsed(pageId);

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

        const layout = ConnectorBuilder.calculateConnectorLayout(fromShapeId, toShapeId, shapeHierarchy, fromPort, toPort);
        const connectorShape = ConnectorBuilder.createConnectorShapeObject(newId, layout, validateArrow(beginArrow), validateArrow(endArrow), style);

        const topLevelShapes = parsed.PageContents.Shapes.Shape;
        topLevelShapes.push(connectorShape);
        this.getShapeMap(parsed).set(newId, connectorShape);

        ConnectorBuilder.addConnectorToConnects(parsed, newId, fromShapeId, toShapeId, shapeHierarchy, fromPort, toPort);

        this.saveParsed(pageId, parsed);

        return newId;
    }

    /**
     * Add a single connection point to an existing shape.
     * Returns the zero-based IX (row index) of the newly added point.
     */
    addConnectionPoint(pageId: string, shapeId: string, point: ConnectionPointDef): number {
        const parsed = this.getParsed(pageId);
        const shapeMap = this.getShapeMap(parsed);
        const shape = shapeMap.get(shapeId);
        if (!shape) throw new Error(`Shape ${shapeId} not found on page ${pageId}`);

        if (!shape.Section) shape.Section = [];
        if (!Array.isArray(shape.Section)) shape.Section = [shape.Section];

        let connSection = shape.Section.find((s: any) => s['@_N'] === SECTION_NAMES.Connection);
        if (!connSection) {
            connSection = { '@_N': SECTION_NAMES.Connection, Row: [] };
            shape.Section.push(connSection);
        }
        if (!connSection.Row) connSection.Row = [];
        if (!Array.isArray(connSection.Row)) connSection.Row = [connSection.Row];

        const ix = connSection.Row.length;
        connSection.Row.push(ConnectionPointBuilder.buildRow(point, ix));

        this.saveParsed(pageId, parsed);
        return ix;
    }

    /**
     * Apply a document-level stylesheet to an existing shape by setting its
     * `LineStyle`, `FillStyle`, and/or `TextStyle` attributes.
     *
     * @param which  `'all'` (default) sets all three; `'line'`, `'fill'`, or `'text'` sets only one.
     */
    applyStyle(
        pageId: string,
        shapeId: string,
        styleId: number,
        which: 'all' | 'line' | 'fill' | 'text' = 'all',
    ): void {
        const parsed   = this.getParsed(pageId);
        const shapeMap = this.getShapeMap(parsed);
        const shape    = shapeMap.get(shapeId);
        if (!shape) throw new Error(`Shape ${shapeId} not found on page ${pageId}`);

        const sid = styleId.toString();
        if (which === 'all'  || which === 'line') shape['@_LineStyle'] = sid;
        if (which === 'all'  || which === 'fill') shape['@_FillStyle'] = sid;
        if (which === 'all'  || which === 'text') shape['@_TextStyle'] = sid;

        this.saveParsed(pageId, parsed);
    }

    async addShape(pageId: string, props: NewShapeProps, parentId?: string): Promise<string> {
        const parsed = this.getParsed(pageId);

        if (!parsed.PageContents.Shapes) {
            parsed.PageContents.Shapes = { Shape: [] };
        }
        let topLevelShapes = parsed.PageContents.Shapes.Shape;
        if (!Array.isArray(topLevelShapes)) {
            topLevelShapes = topLevelShapes ? [topLevelShapes] : [];
            parsed.PageContents.Shapes.Shape = topLevelShapes;
        }

        const newId = props.id || this.getNextId(parsed);

        let newShape: any;

        if (props.type === SHAPE_TYPES.Foreign && props.imgRelId) {
            newShape = ForeignShapeBuilder.createImageShapeObject(newId, props.imgRelId, props);
            if (props.text !== undefined && props.text !== null) {
                newShape.Text = { '#text': props.text };
            }
        } else {
            newShape = ShapeBuilder.createStandardShape(newId, props);

            if (props.masterId) {
                await this.relsManager.ensureRelationship(
                    `visio/pages/page${pageId}.xml`,
                    '../masters/masters.xml',
                    RELATIONSHIP_TYPES.MASTERS
                );
            }
        }

        if (parentId) {
            const parent = this.getShapeMap(parsed).get(parentId);
            if (!parent) {
                throw new Error(`Parent shape ${parentId} not found`);
            }

            if (!parent.Shapes) {
                parent.Shapes = { Shape: [] };
            }
            if (!Array.isArray(parent.Shapes.Shape)) {
                parent.Shapes.Shape = parent.Shapes.Shape ? [parent.Shapes.Shape] : [];
            }

            if (parent['@_Type'] !== SHAPE_TYPES.Group) {
                parent['@_Type'] = SHAPE_TYPES.Group;
            }

            parent.Shapes.Shape.push(newShape);
        } else {
            topLevelShapes.push(newShape);
        }
        this.getShapeMap(parsed).set(newId, newShape);

        this.saveParsed(pageId, parsed);

        return newId;
    }

    async deleteShape(pageId: string, shapeId: string): Promise<void> {
        const parsed = this.getParsed(pageId);

        const removed = this.removeShapeFromTree(parsed.PageContents.Shapes, shapeId);
        if (!removed) {
            throw new Error(`Shape ${shapeId} not found on page ${pageId}`);
        }

        if (parsed.PageContents.Connects?.Connect) {
            let connects = parsed.PageContents.Connects.Connect;
            if (!Array.isArray(connects)) connects = [connects];
            const filtered = connects.filter(
                (c: any) => c['@_FromSheet'] !== shapeId && c['@_ToSheet'] !== shapeId
            );
            parsed.PageContents.Connects.Connect = filtered;
        }

        if (parsed.PageContents.Relationships?.Relationship) {
            let rels = parsed.PageContents.Relationships.Relationship;
            if (!Array.isArray(rels)) rels = [rels];
            parsed.PageContents.Relationships.Relationship = rels.filter(
                (r: any) => r['@_ShapeID'] !== shapeId && r['@_RelatedShapeID'] !== shapeId
            );
        }

        // Invalidate the shape cache so the map is rebuilt on next access.
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

        if (!shape.Section) {
            shape.Section = [];
        } else if (!Array.isArray(shape.Section)) {
            shape.Section = [shape.Section];
        }

        if (style.fillColor) {
            shape.Section = shape.Section.filter((s: any) => s['@_N'] !== SECTION_NAMES.Fill);
            shape.Section.push(createFillSection(style.fillColor));
        }

        const hasLineProps = style.lineColor !== undefined
            || style.lineWeight !== undefined
            || style.linePattern !== undefined;

        if (hasLineProps) {
            shape.Section = shape.Section.filter((s: any) => s['@_N'] !== SECTION_NAMES.Line);
            shape.Section.push(createLineSection({
                color:   style.lineColor,
                weight:  style.lineWeight !== undefined ? (style.lineWeight / 72).toString() : undefined,
                pattern: style.linePattern !== undefined ? style.linePattern.toString() : undefined,
            }));
        }

        const hasCharProps = style.fontColor !== undefined
            || style.bold !== undefined
            || style.italic !== undefined
            || style.underline !== undefined
            || style.strikethrough !== undefined
            || style.fontSize !== undefined
            || style.fontFamily !== undefined;

        if (hasCharProps) {
            shape.Section = shape.Section.filter((s: any) => s['@_N'] !== SECTION_NAMES.Character);
            shape.Section.push(createCharacterSection({
                bold: style.bold,
                italic: style.italic,
                underline: style.underline,
                strikethrough: style.strikethrough,
                color: style.fontColor,
                fontSize: style.fontSize,
                fontFamily: style.fontFamily,
            }));
        }

        const hasParagraphProps = style.horzAlign !== undefined
            || style.spaceBefore !== undefined
            || style.spaceAfter !== undefined
            || style.lineSpacing !== undefined;

        if (hasParagraphProps) {
            shape.Section = shape.Section.filter((s: any) => s['@_N'] !== SECTION_NAMES.Paragraph);
            shape.Section.push(createParagraphSection({
                horzAlign: style.horzAlign,
                spaceBefore: style.spaceBefore,
                spaceAfter: style.spaceAfter,
                lineSpacing: style.lineSpacing,
            }));
        }

        const hasTextBlockProps = style.textMarginTop !== undefined
            || style.textMarginBottom !== undefined
            || style.textMarginLeft !== undefined
            || style.textMarginRight !== undefined;

        if (hasTextBlockProps) {
            shape.Section = shape.Section.filter((s: any) => s['@_N'] !== SECTION_NAMES.TextBlock);
            shape.Section.push(createTextBlockSection({
                topMargin:    style.textMarginTop,
                bottomMargin: style.textMarginBottom,
                leftMargin:   style.textMarginLeft,
                rightMargin:  style.textMarginRight,
            }));
        }

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
                if (section['@_N'] !== SECTION_NAMES.Geometry || !section.Row) continue;
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

        if (!shape.Cell) {
            shape.Cell = [];
        } else if (!Array.isArray(shape.Cell)) {
            shape.Cell = [shape.Cell];
        }

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

        if (!shape.Section) shape.Section = [];
        if (!Array.isArray(shape.Section)) shape.Section = [shape.Section];

        let propSection = shape.Section.find((s: any) => s['@_N'] === SECTION_NAMES.Property);
        if (!propSection) {
            propSection = { '@_N': SECTION_NAMES.Property, Row: [] };
            shape.Section.push(propSection);
        }

        if (!propSection.Row) propSection.Row = [];
        if (!Array.isArray(propSection.Row)) propSection.Row = [propSection.Row];

        const existingRow = propSection.Row.find((r: any) => r['@_N'] === `Prop.${name}`);
        if (existingRow) {
            const updateCell = (n: string, v: string) => {
                let c = existingRow.Cell.find((x: any) => x['@_N'] === n);
                if (c) c['@_V'] = v;
                else existingRow.Cell.push({ '@_N': n, '@_V': v });
            };
            if (options.label !== undefined) updateCell('Label', options.label);
            updateCell('Type', type.toString());
            if (options.invisible !== undefined) updateCell('Invisible', options.invisible ? '1' : '0');
        } else {
            propSection.Row.push({
                '@_N': `Prop.${name}`,
                Cell: [
                    { '@_N': 'Label', '@_V': options.label || name },
                    { '@_N': 'Type', '@_V': type.toString() },
                    { '@_N': 'Invisible', '@_V': options.invisible ? '1' : '0' },
                    { '@_N': 'Value', '@_V': '0' }
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

        const sections = shape.Section ? (Array.isArray(shape.Section) ? shape.Section : [shape.Section]) : [];
        const propSection = sections.find((s: any) => s['@_N'] === SECTION_NAMES.Property);

        if (!propSection) {
            throw new Error(`Property definition 'Prop.${name}' does not exist on shape ${shapeId}. Call addPropertyDefinition first.`);
        }

        const rows = propSection.Row ? (Array.isArray(propSection.Row) ? propSection.Row : [propSection.Row]) : [];
        const row = rows.find((r: any) => r['@_N'] === `Prop.${name}`);

        if (!row) {
            throw new Error(`Property definition 'Prop.${name}' does not exist on shape ${shapeId}. Call addPropertyDefinition first.`);
        }

        let visioValue = '';
        if (value instanceof Date) {
            visioValue = this.dateToVisioString(value);
        } else if (typeof value === 'boolean') {
            visioValue = value ? '1' : '0';
        } else {
            visioValue = value.toString();
        }

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

        if (!parsed.PageContents.Relationships) {
            parsed.PageContents.Relationships = { Relationship: [] };
        }
        if (!Array.isArray(parsed.PageContents.Relationships.Relationship)) {
            parsed.PageContents.Relationships.Relationship = parsed.PageContents.Relationships.Relationship
                ? [parsed.PageContents.Relationships.Relationship]
                : [];
        }

        const relationships = parsed.PageContents.Relationships.Relationship;

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
            .filter((r: any) => r['@_Type'] === STRUCT_RELATIONSHIP_TYPES.Container && r['@_ShapeID'] === containerId)
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
        shapes.splice(idx, 1);

        if (position === 'back') {
            shapes.unshift(shape); // Back of Z-Order
        } else {
            shapes.push(shape); // Front of Z-Order
        }

        shapesContainer.Shape = shapes;
        this.saveParsed(pageId, parsed);
    }

    async addListItem(pageId: string, listId: string, itemId: string): Promise<void> {
        const parsed = this.getParsed(pageId);
        const listShape = this.getShapeMap(parsed).get(listId);
        if (!listShape) throw new Error(`List ${listId} not found`);

        const getUserVal = (name: string, def: string) => {
            if (!listShape.Section) return def;
            const userSec = listShape.Section.find((s: any) => s['@_N'] === SECTION_NAMES.User);
            if (!userSec || !userSec.Row) return def;
            const rows = Array.isArray(userSec.Row) ? userSec.Row : [userSec.Row];
            const row = rows.find((r: any) => r['@_N'] === name);
            if (!row || !row.Cell) return def;
            const valCell = Array.isArray(row.Cell) ? row.Cell.find((c: any) => c['@_N'] === 'Value') : row.Cell;
            return valCell ? valCell['@_V'] : def;
        };

        const direction = parseInt(getUserVal('msvSDListDirection', '1')); // 1=Vert, 0=Horiz
        const spacing = parseFloat(getUserVal('msvSDListSpacing', '0.125').replace(/[^0-9.]/g, ''));

        const memberIds = this.getContainerMembers(pageId, listId);
        const itemGeo = this.getShapeGeometry(pageId, itemId);
        const listGeo = this.getShapeGeometry(pageId, listId);

        let newX = listGeo.x;
        let newY = listGeo.y;

        if (memberIds.length === 0) {
            newX = listGeo.x;
            newY = listGeo.y;
        } else {
            const lastId = memberIds[memberIds.length - 1];
            const lastGeo = this.getShapeGeometry(pageId, lastId);

            if (direction === 1) { // Vertical (Stack Down)
                const lastBottom = lastGeo.y - (lastGeo.height / 2);
                newY = lastBottom - spacing - (itemGeo.height / 2);
                newX = lastGeo.x; // Align Centers
            } else { // Horizontal (Stack Right)
                const lastRight = lastGeo.x + (lastGeo.width / 2);
                newX = lastRight + spacing + (itemGeo.width / 2);
                newY = lastGeo.y; // Align Centers
            }
        }

        await this.updateShapePosition(pageId, itemId, newX, newY);
        await this.addRelationship(pageId, listId, itemId, STRUCT_RELATIONSHIP_TYPES.Container);
        await this.resizeContainerToFit(pageId, listId, 0.25);
    }

    async resizeContainerToFit(pageId: string, containerId: string, padding: number = 0.25): Promise<void> {
        const memberIds = this.getContainerMembers(pageId, containerId);
        if (memberIds.length === 0) return;

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

        minX -= padding;
        maxX += padding;
        minY -= padding;
        maxY += padding;

        const newWidth = maxX - minX;
        const newHeight = maxY - minY;
        const newPinX = minX + (newWidth / 2);
        const newPinY = minY + (newHeight / 2);

        await this.updateShapePosition(pageId, containerId, newPinX, newPinY);
        await this.updateShapeDimensions(pageId, containerId, newWidth, newHeight);
        await this.reorderShape(pageId, containerId, 'back');
    }

    async addHyperlink(pageId: string, shapeId: string, details: { address?: string, subAddress?: string, description?: string }): Promise<void> {
        const parsed = this.getParsed(pageId);
        const shape = this.getShapeMap(parsed).get(shapeId);

        if (!shape) throw new Error(`Shape ${shapeId} not found`);

        if (!shape.Section) shape.Section = [];
        if (!Array.isArray(shape.Section)) shape.Section = [shape.Section];

        let linkSection = shape.Section.find((s: any) => s['@_N'] === 'Hyperlink');
        if (!linkSection) {
            linkSection = { '@_N': 'Hyperlink', Row: [] };
            shape.Section.push(linkSection);
        }

        if (!linkSection.Row) linkSection.Row = [];
        if (!Array.isArray(linkSection.Row)) linkSection.Row = [linkSection.Row];

        const nextIdx = linkSection.Row.length + 1;
        const rowName = `Hyperlink.Row_${nextIdx}`;

        const newRow: any = {
            '@_N': rowName,
            Cell: []
        };

        if (details.address !== undefined) {
            newRow.Cell.push({ '@_N': 'Address', '@_V': details.address });
        }

        if (details.subAddress !== undefined) {
            newRow.Cell.push({ '@_N': 'SubAddress', '@_V': details.subAddress });
        }

        if (details.description !== undefined) {
            newRow.Cell.push({ '@_N': 'Description', '@_V': details.description });
        }

        newRow.Cell.push({ '@_N': 'NewWindow', '@_V': '0' });

        linkSection.Row.push(newRow);

        this.saveParsed(pageId, parsed);
    }

    async addLayer(pageId: string, name: string, options: { visible?: boolean, lock?: boolean, print?: boolean } = {}): Promise<{ name: string, index: number }> {
        const parsed = this.getParsed(pageId);

        this.ensurePageSheet(parsed);
        const pageSheet = parsed.PageContents.PageSheet;

        if (!pageSheet.Section) pageSheet.Section = [];
        if (!Array.isArray(pageSheet.Section)) pageSheet.Section = [pageSheet.Section];

        let layerSection = pageSheet.Section.find((s: any) => s['@_N'] === SECTION_NAMES.Layer);
        if (!layerSection) {
            layerSection = { '@_N': SECTION_NAMES.Layer, Row: [] };
            pageSheet.Section.push(layerSection);
        }

        if (!layerSection.Row) layerSection.Row = [];
        if (!Array.isArray(layerSection.Row)) layerSection.Row = [layerSection.Row];

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

        if (!shape.Section) shape.Section = [];
        if (!Array.isArray(shape.Section)) shape.Section = [shape.Section];

        let memSection = shape.Section.find((s: any) => s['@_N'] === SECTION_NAMES.LayerMem);
        if (!memSection) {
            memSection = { '@_N': SECTION_NAMES.LayerMem, Row: [] };
            shape.Section.push(memSection);
        }

        if (!memSection.Row) memSection.Row = [];
        if (!Array.isArray(memSection.Row)) memSection.Row = [memSection.Row];

        // LayerMem contains a single row with a semicolon-separated LayerMember cell.
        if (memSection.Row.length === 0) {
            memSection.Row.push({ Cell: [] });
        }
        const row = memSection.Row[0];

        if (!row.Cell) row.Cell = [];
        if (!Array.isArray(row.Cell)) row.Cell = [row.Cell];

        let cell = row.Cell.find((c: any) => c['@_N'] === 'LayerMember');
        if (!cell) {
            cell = { '@_N': 'LayerMember', '@_V': '' };
            row.Cell.push(cell);
        }

        const currentVal = cell['@_V'] || '';
        const indices = currentVal.split(';').filter((s: string) => s.length > 0);
        const idxStr = layerIndex.toString();

        if (!indices.includes(idxStr)) {
            indices.push(idxStr);
            cell['@_V'] = indices.join(';');
            this.saveParsed(pageId, parsed);
        }
    }

    async updateLayerProperty(pageId: string, layerIndex: number, propName: string, value: string): Promise<void> {
        const parsed = this.getParsed(pageId);

        this.ensurePageSheet(parsed);
        const pageSheet = parsed.PageContents.PageSheet;

        if (!pageSheet.Section) return;
        const sections = Array.isArray(pageSheet.Section) ? pageSheet.Section : [pageSheet.Section];
        const layerSection = sections.find((s: any) => s['@_N'] === SECTION_NAMES.Layer);
        if (!layerSection || !layerSection.Row) return;

        const rows = Array.isArray(layerSection.Row) ? layerSection.Row : [layerSection.Row];
        const row = rows.find((r: any) => r['@_IX'] == layerIndex.toString());
        if (!row) return;

        if (!row.Cell) row.Cell = [];
        if (!Array.isArray(row.Cell)) row.Cell = [row.Cell];

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
     * Return all layers defined in the page's PageSheet as plain objects.
     */
    getPageLayers(pageId: string): Array<{ name: string; index: number; visible: boolean; locked: boolean }> {
        const parsed = this.getParsed(pageId);
        const pageSheet = parsed.PageContents?.PageSheet;
        if (!pageSheet?.Section) return [];

        const sections = Array.isArray(pageSheet.Section) ? pageSheet.Section : [pageSheet.Section];
        const layerSection = sections.find((s: any) => s['@_N'] === SECTION_NAMES.Layer);
        if (!layerSection?.Row) return [];

        const rows = Array.isArray(layerSection.Row) ? layerSection.Row : [layerSection.Row];
        return rows.map((row: any) => {
            const cells: any[] = Array.isArray(row.Cell) ? row.Cell : (row.Cell ? [row.Cell] : []);
            const getVal = (name: string) => cells.find((c: any) => c['@_N'] === name)?.['@_V'];
            return {
                name:    getVal('Name')    ?? '',
                index:   parseInt(row['@_IX'], 10),
                visible: getVal('Visible') !== '0',
                locked:  getVal('Lock')    === '1',
            };
        });
    }

    /**
     * Delete a layer by index and remove it from all shape LayerMember cells.
     */
    deleteLayer(pageId: string, layerIndex: number): void {
        const parsed = this.getParsed(pageId);

        const pageSheet = parsed.PageContents?.PageSheet;
        if (pageSheet?.Section) {
            const sections = Array.isArray(pageSheet.Section) ? pageSheet.Section : [pageSheet.Section];
            const layerSection = sections.find((s: any) => s['@_N'] === SECTION_NAMES.Layer);
            if (layerSection?.Row) {
                const rows = Array.isArray(layerSection.Row) ? layerSection.Row : [layerSection.Row];
                layerSection.Row = rows.filter((r: any) => r['@_IX'] !== layerIndex.toString());
            }
        }

        // Remove this layer index from every shape's LayerMember cell
        const idxStr = layerIndex.toString();
        for (const [, shape] of this.getShapeMap(parsed)) {
            if (!shape.Section) continue;
            const sections: any[] = Array.isArray(shape.Section) ? shape.Section : [shape.Section];
            const memSec = sections.find((s: any) => s['@_N'] === SECTION_NAMES.LayerMem);
            if (!memSec?.Row) continue;

            const rows = Array.isArray(memSec.Row) ? memSec.Row : [memSec.Row];
            const row = rows[0];
            if (!row?.Cell) continue;

            const cells: any[] = Array.isArray(row.Cell) ? row.Cell : [row.Cell];
            const memberCell = cells.find((c: any) => c['@_N'] === 'LayerMember');
            if (!memberCell?.['@_V']) continue;

            const remaining = memberCell['@_V']
                .split(';')
                .filter((s: string) => s.length > 0 && s !== idxStr);
            memberCell['@_V'] = remaining.join(';');
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
        const propSection = sections.find((s: any) => s['@_N'] === SECTION_NAMES.Property);
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
     * Read the drawing scale from the PageSheet.
     * Returns `null` when no custom scale is set (i.e. 1:1 / no-scale default).
     */
    getDrawingScale(pageId: string): DrawingScaleInfo | null {
        const parsed = this.getParsed(pageId);
        const ps = parsed.PageContents?.PageSheet;
        if (!ps?.Cell) return null;
        const cells: any[] = Array.isArray(ps.Cell) ? ps.Cell : [ps.Cell];
        const getCell = (name: string) => cells.find((c: any) => c['@_N'] === name);

        const psCell = getCell('PageScale');
        const dsCell = getCell('DrawingScale');
        if (!psCell && !dsCell) return null;

        const psUnit  = psCell?.['@_Unit'] ?? 'MSG';
        const dsUnit  = dsCell?.['@_Unit'] ?? 'MSG';
        // "MSG" is Visio's sentinel for "no real unit" — treat as 1:1
        if (psUnit === 'MSG' && dsUnit === 'MSG') return null;

        const toUserUnit = (v: string): LengthUnit =>
            (VISIO_TO_LENGTH_UNIT[v] as LengthUnit | undefined) ?? 'in';

        return {
            pageScale:   parseFloat(psCell?.['@_V'] ?? '1'),
            pageUnit:    toUserUnit(psUnit),
            drawingScale: parseFloat(dsCell?.['@_V'] ?? '1'),
            drawingUnit:  toUserUnit(dsUnit),
        };
    }

    /**
     * Set a custom drawing scale on the PageSheet.
     * @param pageScale   Measurement on paper (e.g. 1).
     * @param pageUnit    Unit for the paper measurement (e.g. `'in'`).
     * @param drawingScale Real-world measurement (e.g. 10).
     * @param drawingUnit  Unit for the real-world measurement (e.g. `'ft'`).
     */
    setDrawingScale(
        pageId: string,
        pageScale: number, pageUnit: LengthUnit,
        drawingScale: number, drawingUnit: LengthUnit
    ): void {
        if (pageScale <= 0 || drawingScale <= 0) {
            throw new Error('Drawing scale values must be positive');
        }
        const parsed = this.getParsed(pageId);
        this.ensurePageSheet(parsed);
        const ps = parsed.PageContents.PageSheet;

        const upsertWithUnit = (name: string, value: string, unit: string) => {
            const existing = ps.Cell.find((c: any) => c['@_N'] === name);
            if (existing) {
                existing['@_V'] = value;
                existing['@_Unit'] = unit;
            } else {
                ps.Cell.push({ '@_N': name, '@_V': value, '@_Unit': unit });
            }
        };

        upsertWithUnit('PageScale',   pageScale.toString(),   LENGTH_UNIT_TO_VISIO[pageUnit]);
        upsertWithUnit('DrawingScale', drawingScale.toString(), LENGTH_UNIT_TO_VISIO[drawingUnit]);

        this.saveParsed(pageId, parsed);
    }

    /**
     * Reset the drawing scale to 1:1 (no custom scale).
     * Restores `PageScale` and `DrawingScale` to `V="1" Unit="MSG"`.
     */
    clearDrawingScale(pageId: string): void {
        const parsed = this.getParsed(pageId);
        this.ensurePageSheet(parsed);
        const ps = parsed.PageContents.PageSheet;

        const reset = (name: string) => {
            const existing = ps.Cell.find((c: any) => c['@_N'] === name);
            if (existing) {
                existing['@_V'] = '1';
                existing['@_Unit'] = 'MSG';
            } else {
                ps.Cell.push({ '@_N': name, '@_V': '1', '@_Unit': 'MSG' });
            }
        };

        reset('PageScale');
        reset('DrawingScale');
        this.saveParsed(pageId, parsed);
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
        const memSection = sections.find((s: any) => s['@_N'] === SECTION_NAMES.LayerMem);
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

    /**
     * Return the direct child shapes of a group or container shape.
     * Returns an empty array for non-group shapes or shapes with no children.
     */
    getShapeChildren(pageId: string, shapeId: string): VisioShape[] {
        const pagePath = this.getPagePath(pageId);
        const reader = new ShapeReader(this.pkg);
        return reader.readChildShapes(pagePath, shapeId);
    }
}


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
