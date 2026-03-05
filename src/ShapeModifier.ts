import { VisioPackage } from './VisioPackage';
import { RelsManager } from './core/RelsManager';
import { PageXmlCache } from './core/PageXmlCache';
import { PageSheetEditor } from './core/PageSheetEditor';
import { LayerEditor } from './core/LayerEditor';
import { ContainerEditor } from './core/ContainerEditor';
import { ConnectorEditor } from './core/ConnectorEditor';
import { createFillSection, createCharacterSection, createLineSection, createParagraphSection, createTextBlockSection, vertAlignValue, horzAlignValue, hexToRgb } from './utils/StyleHelpers';
import { RELATIONSHIP_TYPES, SECTION_NAMES, SHAPE_TYPES } from './core/VisioConstants';
import { NewShapeProps, VisioPropType, ConnectorStyle, ConnectionTarget, ConnectionPointDef, DrawingScaleInfo, LengthUnit, ShapeStyle } from './types/VisioTypes';
import { ConnectionPointBuilder } from './shapes/ConnectionPointBuilder';
import { ShapeReader } from './ShapeReader';
import type { ShapeData, ShapeHyperlink } from './Shape';
import type { VisioShape } from './types/VisioTypes';
import { ForeignShapeBuilder } from './shapes/ForeignShapeBuilder';
import { ShapeBuilder } from './shapes/ShapeBuilder';

export class ShapeModifier {
    private cache: PageXmlCache;
    private relsManager: RelsManager;
    private pageSheet: PageSheetEditor;
    private layers: LayerEditor;
    private containers: ContainerEditor;
    private connectors: ConnectorEditor;

    constructor(private pkg: VisioPackage) {
        this.cache        = new PageXmlCache(pkg);
        this.relsManager  = new RelsManager(pkg);
        this.pageSheet    = new PageSheetEditor(this.cache);
        this.layers       = new LayerEditor(this.cache);
        this.containers   = new ContainerEditor(this.cache);
        this.connectors   = new ConnectorEditor(this.cache, this.relsManager);
    }

    // ── Infrastructure pass-throughs ─────────────────────────────────────────

    get autoSave(): boolean { return this.cache.autoSave; }
    set autoSave(v: boolean) { this.cache.autoSave = v; }

    registerPage(pageId: string, xmlPath: string): void {
        this.cache.registerPage(pageId, xmlPath);
    }

    flush(): void { this.cache.flush(); }

    /** @internal Exposed for test introspection. */
    getParsed(pageId: string): any { return this.cache.getParsed(pageId); }

    /** @internal Exposed for test introspection. */
    getAllShapes(parsed: any): any[] { return this.cache.getAllShapes(parsed); }

    // ── Connector ─────────────────────────────────────────────────────────────

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
        return this.connectors.addConnector(pageId, fromShapeId, toShapeId, beginArrow, endArrow, style, fromPort, toPort);
    }

    // ── Container / List ──────────────────────────────────────────────────────

    async addContainer(pageId: string, props: NewShapeProps): Promise<string> {
        return this.containers.addContainer(pageId, props);
    }

    async addList(pageId: string, props: NewShapeProps, direction: 'vertical' | 'horizontal' = 'vertical'): Promise<string> {
        return this.containers.addList(pageId, props, direction);
    }

    async addRelationship(pageId: string, shapeId: string, relatedShapeId: string, type: string): Promise<void> {
        return this.containers.addRelationship(pageId, shapeId, relatedShapeId, type);
    }

    getContainerMembers(pageId: string, containerId: string): string[] {
        return this.containers.getContainerMembers(pageId, containerId);
    }

    async reorderShape(pageId: string, shapeId: string, position: 'front' | 'back'): Promise<void> {
        return this.containers.reorderShape(pageId, shapeId, position);
    }

    async addListItem(pageId: string, listId: string, itemId: string): Promise<void> {
        return this.containers.addListItem(pageId, listId, itemId);
    }

    async resizeContainerToFit(pageId: string, containerId: string, padding: number = 0.25): Promise<void> {
        return this.containers.resizeContainerToFit(pageId, containerId, padding);
    }

    // ── Layers ────────────────────────────────────────────────────────────────

    async addLayer(pageId: string, name: string, options: { visible?: boolean; lock?: boolean; print?: boolean } = {}): Promise<{ name: string; index: number }> {
        return this.layers.addLayer(pageId, name, options);
    }

    async assignLayer(pageId: string, shapeId: string, layerIndex: number): Promise<void> {
        return this.layers.assignLayer(pageId, shapeId, layerIndex);
    }

    async updateLayerProperty(pageId: string, layerIndex: number, propName: string, value: string): Promise<void> {
        return this.layers.updateLayerProperty(pageId, layerIndex, propName, value);
    }

    getPageLayers(pageId: string): Array<{ name: string; index: number; visible: boolean; locked: boolean }> {
        return this.layers.getPageLayers(pageId);
    }

    deleteLayer(pageId: string, layerIndex: number): void {
        return this.layers.deleteLayer(pageId, layerIndex);
    }

    getShapeLayerIndices(pageId: string, shapeId: string): number[] {
        return this.layers.getShapeLayerIndices(pageId, shapeId);
    }

    // ── Page sheet ────────────────────────────────────────────────────────────

    /**
     * Set the page canvas size. Writes PageWidth / PageHeight into the PageSheet
     * and sets DrawingSizeType=0 (Custom) so Visio does not override the values.
     */
    setPageSize(pageId: string, width: number, height: number): void {
        return this.pageSheet.setPageSize(pageId, width, height);
    }

    /**
     * Read the current page canvas dimensions.
     * Returns 8.5 × 11 (US Letter) if no PageSheet cells are present.
     */
    getPageDimensions(pageId: string): { width: number; height: number } {
        return this.pageSheet.getPageDimensions(pageId);
    }

    /**
     * Read the drawing scale from the PageSheet.
     * Returns `null` when no custom scale is set (i.e. 1:1 / no-scale default).
     */
    getDrawingScale(pageId: string): DrawingScaleInfo | null {
        return this.pageSheet.getDrawingScale(pageId);
    }

    /**
     * Set a custom drawing scale on the PageSheet.
     * @param pageScale    Measurement on paper (e.g. 1).
     * @param pageUnit     Unit for the paper measurement (e.g. `'in'`).
     * @param drawingScale Real-world measurement (e.g. 10).
     * @param drawingUnit  Unit for the real-world measurement (e.g. `'ft'`).
     */
    setDrawingScale(
        pageId: string,
        pageScale: number, pageUnit: LengthUnit,
        drawingScale: number, drawingUnit: LengthUnit,
    ): void {
        return this.pageSheet.setDrawingScale(pageId, pageScale, pageUnit, drawingScale, drawingUnit);
    }

    /**
     * Reset the drawing scale to 1:1 (no custom scale).
     * Restores `PageScale` and `DrawingScale` to `V="1" Unit="MSG"`.
     */
    clearDrawingScale(pageId: string): void {
        return this.pageSheet.clearDrawingScale(pageId);
    }

    // ── Shape geometry (delegates to cache) ──────────────────────────────────

    getShapeGeometry(pageId: string, shapeId: string): { x: number; y: number; width: number; height: number } {
        return this.cache.getShapeGeometry(pageId, shapeId);
    }

    async updateShapePosition(pageId: string, shapeId: string, x: number, y: number): Promise<void> {
        return this.cache.updateShapePosition(pageId, shapeId, x, y);
    }

    async updateShapeDimensions(pageId: string, shapeId: string, w: number, h: number): Promise<void> {
        return this.cache.updateShapeDimensions(pageId, shapeId, w, h);
    }

    // ── Shape CRUD ────────────────────────────────────────────────────────────

    /**
     * Add a single connection point to an existing shape.
     * Returns the zero-based IX (row index) of the newly added point.
     */
    addConnectionPoint(pageId: string, shapeId: string, point: ConnectionPointDef): number {
        const parsed = this.cache.getParsed(pageId);
        const shape  = this.cache.getShapeMap(parsed).get(shapeId);
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
        const shapeCells: any[] = Array.isArray(shape.Cell) ? shape.Cell : shape.Cell ? [shape.Cell] : [];
        const shapeWidth  = parseFloat(shapeCells.find((c: any) => c['@_N'] === 'Width' )?.['@_V'] ?? '1');
        const shapeHeight = parseFloat(shapeCells.find((c: any) => c['@_N'] === 'Height')?.['@_V'] ?? '1');
        connSection.Row.push(ConnectionPointBuilder.buildRow(point, ix, shapeWidth, shapeHeight));

        this.cache.saveParsed(pageId, parsed);
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
        const parsed   = this.cache.getParsed(pageId);
        const shapeMap = this.cache.getShapeMap(parsed);
        const shape    = shapeMap.get(shapeId);
        if (!shape) throw new Error(`Shape ${shapeId} not found on page ${pageId}`);

        const sid = styleId.toString();
        if (which === 'all' || which === 'line') shape['@_LineStyle'] = sid;
        if (which === 'all' || which === 'fill') shape['@_FillStyle'] = sid;
        if (which === 'all' || which === 'text') shape['@_TextStyle'] = sid;

        this.cache.saveParsed(pageId, parsed);
    }

    async addShape(pageId: string, props: NewShapeProps, parentId?: string): Promise<string> {
        const parsed = this.cache.getParsed(pageId);

        if (!parsed.PageContents.Shapes) {
            parsed.PageContents.Shapes = { Shape: [] };
        }
        let topLevelShapes = parsed.PageContents.Shapes.Shape;
        if (!Array.isArray(topLevelShapes)) {
            topLevelShapes = topLevelShapes ? [topLevelShapes] : [];
            parsed.PageContents.Shapes.Shape = topLevelShapes;
        }

        const newId = props.id || this.cache.getNextId(parsed);

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
                    this.cache.getPagePath(pageId),
                    '../masters/masters.xml',
                    RELATIONSHIP_TYPES.MASTERS,
                );
            }
        }

        if (parentId) {
            const parent = this.cache.getShapeMap(parsed).get(parentId);
            if (!parent) throw new Error(`Parent shape ${parentId} not found`);

            if (!parent.Shapes) parent.Shapes = { Shape: [] };
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

        this.cache.getShapeMap(parsed).set(newId, newShape);
        this.cache.saveParsed(pageId, parsed);
        return newId;
    }

    async deleteShape(pageId: string, shapeId: string): Promise<void> {
        const parsed = this.cache.getParsed(pageId);

        // Collect the deleted shape's ID and all descendant IDs before removal
        // so Connect/Relationship entries for child shapes are also cleaned up.
        const shapeNode = this.findShapeInTree(parsed.PageContents.Shapes, shapeId);
        if (!shapeNode) throw new Error(`Shape ${shapeId} not found on page ${pageId}`);
        const removedIds = this.collectShapeIds(shapeNode);

        this.removeShapeFromTree(parsed.PageContents.Shapes, shapeId);

        if (parsed.PageContents.Connects?.Connect) {
            let connects = parsed.PageContents.Connects.Connect;
            if (!Array.isArray(connects)) connects = [connects];
            parsed.PageContents.Connects.Connect = connects.filter(
                (c: any) => !removedIds.has(c['@_FromSheet']) && !removedIds.has(c['@_ToSheet']),
            );
        }

        if (parsed.PageContents.Relationships?.Relationship) {
            let rels = parsed.PageContents.Relationships.Relationship;
            if (!Array.isArray(rels)) rels = [rels];
            parsed.PageContents.Relationships.Relationship = rels.filter(
                (r: any) => !removedIds.has(r['@_ShapeID']) && !removedIds.has(r['@_RelatedShapeID']),
            );
        }

        // Invalidate the shape cache so the map is rebuilt on next access.
        this.cache.shapeCache.delete(parsed);
        this.cache.saveParsed(pageId, parsed);
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

    private findShapeInTree(shapesContainer: any, shapeId: string): any | null {
        if (!shapesContainer?.Shape) return null;
        const shapes = Array.isArray(shapesContainer.Shape) ? shapesContainer.Shape : [shapesContainer.Shape];
        for (const s of shapes) {
            if (s['@_ID'] === shapeId) return s;
            if (s.Shapes) {
                const found = this.findShapeInTree(s.Shapes, shapeId);
                if (found) return found;
            }
        }
        return null;
    }

    private collectShapeIds(shape: any): Set<string> {
        const ids = new Set<string>();
        const gather = (s: any) => {
            ids.add(s['@_ID']);
            if (s.Shapes?.Shape) {
                const children = Array.isArray(s.Shapes.Shape) ? s.Shapes.Shape : [s.Shapes.Shape];
                for (const child of children) gather(child);
            }
        };
        gather(shape);
        return ids;
    }

    async updateShapeText(pageId: string, shapeId: string, newText: string): Promise<void> {
        const parsed = this.cache.getParsed(pageId);
        const shape  = this.cache.getShapeMap(parsed).get(shapeId);
        if (!shape) throw new Error(`Shape ${shapeId} not found on page ${pageId}`);

        shape.Text = { '#text': newText };
        this.cache.saveParsed(pageId, parsed);
    }

    async updateShapeStyle(pageId: string, shapeId: string, style: ShapeStyle): Promise<void> {
        const parsed = this.cache.getParsed(pageId);
        const shape  = this.cache.getShapeMap(parsed).get(shapeId);
        if (!shape) throw new Error(`Shape ${shapeId} not found on page ${pageId}`);

        if (!shape.Section) {
            shape.Section = [];
        } else if (!Array.isArray(shape.Section)) {
            shape.Section = [shape.Section];
        }

        // Upsert a single cell by name inside a flat Cell[].
        const upsertCell = (cells: any[], name: string, val: string, extra?: Record<string, string>) => {
            const existing = cells.find((c: any) => c['@_N'] === name);
            if (existing) {
                existing['@_V'] = val;
                if (extra) Object.assign(existing, extra);
            } else {
                cells.push({ '@_N': name, '@_V': val, ...extra });
            }
        };

        // Normalise and return the Cell[] of a flat-cell section.
        const ensureCells = (section: any): any[] => {
            if (!section.Cell) section.Cell = [];
            else if (!Array.isArray(section.Cell)) section.Cell = [section.Cell];
            return section.Cell as any[];
        };

        // Return (creating if absent) the Cell[] of Row[0] in a row-based section.
        const getOrCreateRow0Cells = (section: any, rowType: string): any[] => {
            if (!section.Row) section.Row = [];
            else if (!Array.isArray(section.Row)) section.Row = [section.Row];
            let row0 = section.Row.find((r: any) => r['@_IX'] === '0' || r['@_IX'] === 0);
            if (!row0) {
                row0 = { '@_T': rowType, '@_IX': '0', Cell: [] };
                section.Row.push(row0);
            }
            if (!row0.Cell) row0.Cell = [];
            else if (!Array.isArray(row0.Cell)) row0.Cell = [row0.Cell];
            return row0.Cell as any[];
        };

        if (style.fillColor) {
            const existing = shape.Section.find((s: any) => s['@_N'] === SECTION_NAMES.Fill);
            if (existing) {
                upsertCell(ensureCells(existing), 'FillForegnd', style.fillColor, { '@_F': hexToRgb(style.fillColor) });
            } else {
                shape.Section.push(createFillSection(style.fillColor));
            }
        }

        const hasLineProps = style.lineColor !== undefined
            || style.lineWeight !== undefined
            || style.linePattern !== undefined;

        if (hasLineProps) {
            const existing = shape.Section.find((s: any) => s['@_N'] === SECTION_NAMES.Line);
            if (existing) {
                const cells = ensureCells(existing);
                if (style.lineColor   !== undefined) upsertCell(cells, 'LineColor',   style.lineColor,                           { '@_F': hexToRgb(style.lineColor) });
                if (style.lineWeight  !== undefined) upsertCell(cells, 'LineWeight',  (style.lineWeight / 72).toString(),        { '@_U': 'IN' });
                if (style.linePattern !== undefined) upsertCell(cells, 'LinePattern', style.linePattern.toString());
            } else {
                shape.Section.push(createLineSection({
                    color:   style.lineColor,
                    weight:  style.lineWeight  !== undefined ? (style.lineWeight / 72).toString() : undefined,
                    pattern: style.linePattern !== undefined ? style.linePattern.toString()       : undefined,
                }));
            }
        }

        const hasCharProps = style.fontColor !== undefined
            || style.bold !== undefined
            || style.italic !== undefined
            || style.underline !== undefined
            || style.strikethrough !== undefined
            || style.fontSize !== undefined
            || style.fontFamily !== undefined;

        if (hasCharProps) {
            const existing = shape.Section.find((s: any) => s['@_N'] === SECTION_NAMES.Character);
            if (existing) {
                const cells = getOrCreateRow0Cells(existing, 'Character');
                const hasStyleBits = style.bold !== undefined || style.italic !== undefined
                    || style.underline !== undefined || style.strikethrough !== undefined;
                if (hasStyleBits) {
                    const styleCell = cells.find((c: any) => c['@_N'] === 'Style');
                    let styleVal = styleCell ? (parseInt(styleCell['@_V'] || '0') || 0) : 0;
                    if (style.bold          !== undefined) { if (style.bold)          styleVal |= 1; else styleVal &= ~1; }
                    if (style.italic        !== undefined) { if (style.italic)        styleVal |= 2; else styleVal &= ~2; }
                    if (style.underline     !== undefined) { if (style.underline)     styleVal |= 4; else styleVal &= ~4; }
                    if (style.strikethrough !== undefined) { if (style.strikethrough) styleVal |= 8; else styleVal &= ~8; }
                    upsertCell(cells, 'Style', styleVal.toString());
                }
                if (style.fontColor  !== undefined) upsertCell(cells, 'Color', style.fontColor,                    { '@_F': hexToRgb(style.fontColor) });
                if (style.fontSize   !== undefined) upsertCell(cells, 'Size',  (style.fontSize / 72).toString(),   { '@_U': 'PT' });
                if (style.fontFamily !== undefined) upsertCell(cells, 'Font',  '0',                                { '@_F': `FONT("${style.fontFamily}")` });
            } else {
                shape.Section.push(createCharacterSection({
                    bold:          style.bold,
                    italic:        style.italic,
                    underline:     style.underline,
                    strikethrough: style.strikethrough,
                    color:         style.fontColor,
                    fontSize:      style.fontSize,
                    fontFamily:    style.fontFamily,
                }));
            }
        }

        const hasParagraphProps = style.horzAlign !== undefined
            || style.spaceBefore !== undefined
            || style.spaceAfter !== undefined
            || style.lineSpacing !== undefined;

        if (hasParagraphProps) {
            const existing = shape.Section.find((s: any) => s['@_N'] === SECTION_NAMES.Paragraph);
            if (existing) {
                const cells = getOrCreateRow0Cells(existing, 'Paragraph');
                if (style.horzAlign   !== undefined) upsertCell(cells, 'HorzAlign', horzAlignValue(style.horzAlign));
                if (style.spaceBefore !== undefined) upsertCell(cells, 'SpBefore',  (style.spaceBefore / 72).toString(), { '@_U': 'PT' });
                if (style.spaceAfter  !== undefined) upsertCell(cells, 'SpAfter',   (style.spaceAfter  / 72).toString(), { '@_U': 'PT' });
                if (style.lineSpacing !== undefined) upsertCell(cells, 'SpLine',    (-style.lineSpacing).toString());
            } else {
                shape.Section.push(createParagraphSection({
                    horzAlign:   style.horzAlign,
                    spaceBefore: style.spaceBefore,
                    spaceAfter:  style.spaceAfter,
                    lineSpacing: style.lineSpacing,
                }));
            }
        }

        const hasTextBlockProps = style.textMarginTop !== undefined
            || style.textMarginBottom !== undefined
            || style.textMarginLeft !== undefined
            || style.textMarginRight !== undefined;

        if (hasTextBlockProps) {
            const existing = shape.Section.find((s: any) => s['@_N'] === SECTION_NAMES.TextBlock);
            if (existing) {
                const cells = ensureCells(existing);
                if (style.textMarginTop    !== undefined) upsertCell(cells, 'TopMargin',    style.textMarginTop.toString(),    { '@_U': 'IN' });
                if (style.textMarginBottom !== undefined) upsertCell(cells, 'BottomMargin', style.textMarginBottom.toString(), { '@_U': 'IN' });
                if (style.textMarginLeft   !== undefined) upsertCell(cells, 'LeftMargin',   style.textMarginLeft.toString(),   { '@_U': 'IN' });
                if (style.textMarginRight  !== undefined) upsertCell(cells, 'RightMargin',  style.textMarginRight.toString(),  { '@_U': 'IN' });
            } else {
                shape.Section.push(createTextBlockSection({
                    topMargin:    style.textMarginTop,
                    bottomMargin: style.textMarginBottom,
                    leftMargin:   style.textMarginLeft,
                    rightMargin:  style.textMarginRight,
                }));
            }
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

        this.cache.saveParsed(pageId, parsed);
    }

    /**
     * Set the rotation angle of a shape. Degrees are converted to radians
     * for storage in the Angle cell (Visio's native unit).
     */
    async rotateShape(pageId: string, shapeId: string, degrees: number): Promise<void> {
        const parsed = this.cache.getParsed(pageId);
        const shape  = this.cache.getShapeMap(parsed).get(shapeId);
        if (!shape) throw new Error(`Shape ${shapeId} not found on page ${pageId}`);

        if (!shape.Cell) shape.Cell = [];
        if (!Array.isArray(shape.Cell)) shape.Cell = [shape.Cell];

        const radians  = (degrees * Math.PI) / 180;
        const existing = shape.Cell.find((c: any) => c['@_N'] === 'Angle');
        if (existing) existing['@_V'] = radians.toString();
        else shape.Cell.push({ '@_N': 'Angle', '@_V': radians.toString() });

        this.cache.saveParsed(pageId, parsed);
    }

    /**
     * Set the flip state for a shape along the X or Y axis.
     * FlipX mirrors left-to-right; FlipY mirrors top-to-bottom.
     */
    setShapeFlip(pageId: string, shapeId: string, axis: 'x' | 'y', enabled: boolean): void {
        const parsed = this.cache.getParsed(pageId);
        const shape  = this.cache.getShapeMap(parsed).get(shapeId);
        if (!shape) throw new Error(`Shape ${shapeId} not found on page ${pageId}`);

        if (!shape.Cell) shape.Cell = [];
        if (!Array.isArray(shape.Cell)) shape.Cell = [shape.Cell];

        const cellName = axis === 'x' ? 'FlipX' : 'FlipY';
        const value    = enabled ? '1' : '0';
        const existing = shape.Cell.find((c: any) => c['@_N'] === cellName);
        if (existing) existing['@_V'] = value;
        else shape.Cell.push({ '@_N': cellName, '@_V': value });

        this.cache.saveParsed(pageId, parsed);
    }

    /**
     * Resize a shape, keeping it centred on its current PinX/PinY.
     * Updates Width, Height, LocPinX, LocPinY, and the cached @_V on any
     * Geometry cells whose @_F formula references Width or Height, so that
     * non-Visio renderers see consistent values.
     */
    async resizeShape(pageId: string, shapeId: string, width: number, height: number): Promise<void> {
        const parsed = this.cache.getParsed(pageId);
        const shape  = this.cache.getShapeMap(parsed).get(shapeId);
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

        // Keep cached @_V consistent for renderers that don't evaluate formulas.
        // Evaluate any formula that references Width or Height (e.g. 'Width*0.5').
        if (shape.Section) {
            const sections = Array.isArray(shape.Section) ? shape.Section : [shape.Section];
            for (const section of sections) {
                if (section['@_N'] !== SECTION_NAMES.Geometry || !section.Row) continue;
                const rows = Array.isArray(section.Row) ? section.Row : [section.Row];
                for (const row of rows) {
                    if (!row.Cell) continue;
                    const cells = Array.isArray(row.Cell) ? row.Cell : [row.Cell];
                    for (const cell of cells) {
                        const formula: unknown = cell['@_F'];
                        if (typeof formula !== 'string') continue;
                        if (!formula.includes('Width') && !formula.includes('Height')) continue;
                        const expr = formula
                            .replace(/\bWidth\b/g, width.toString())
                            .replace(/\bHeight\b/g, height.toString());
                        // Validate expression contains only safe arithmetic characters
                        if (!/^[\d\s.+\-*/()]+$/.test(expr)) continue;
                        try {
                            // eslint-disable-next-line no-new-func
                            const result = Function(`"use strict"; return (${expr})`)() as number;
                            cell['@_V'] = result.toString();
                        } catch { /* ignore malformed formulas */ }
                    }
                }
            }
        }

        this.cache.saveParsed(pageId, parsed);
    }

    addPropertyDefinition(
        pageId: string,
        shapeId: string,
        name: string,
        type: number,
        options: { label?: string; invisible?: boolean } = {},
    ): void {
        const parsed = this.cache.getParsed(pageId);
        const shape  = this.cache.getShapeMap(parsed).get(shapeId);
        if (!shape) throw new Error(`Shape ${shapeId} not found on page ${pageId}`);

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
                    { '@_N': 'Label',     '@_V': options.label || name },
                    { '@_N': 'Type',      '@_V': type.toString() },
                    { '@_N': 'Invisible', '@_V': options.invisible ? '1' : '0' },
                    { '@_N': 'Value',     '@_V': '0' },
                ],
            });
        }

        this.cache.saveParsed(pageId, parsed);
    }

    private dateToVisioString(date: Date): string {
        // Visio accepts ISO 8601 strings for date properties (Type 5).
        return date.toISOString().split('.')[0]; // strip milliseconds
    }

    setPropertyValue(pageId: string, shapeId: string, name: string, value: string | number | boolean | Date): void {
        const parsed = this.cache.getParsed(pageId);
        const shape  = this.cache.getShapeMap(parsed).get(shapeId);
        if (!shape) throw new Error(`Shape ${shapeId} not found on page ${pageId}`);

        const sections   = shape.Section ? (Array.isArray(shape.Section) ? shape.Section : [shape.Section]) : [];
        const propSection = sections.find((s: any) => s['@_N'] === SECTION_NAMES.Property);
        if (!propSection) {
            throw new Error(`Property definition 'Prop.${name}' does not exist on shape ${shapeId}. Call addPropertyDefinition first.`);
        }

        const rows = propSection.Row ? (Array.isArray(propSection.Row) ? propSection.Row : [propSection.Row]) : [];
        const row  = rows.find((r: any) => r['@_N'] === `Prop.${name}`);
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

        this.cache.saveParsed(pageId, parsed);
    }

    /**
     * Read back all custom property (shape data) entries for a shape.
     * Returns a map of property key → ShapeData, with values coerced to
     * the declared type (Number, Boolean, Date, or String).
     */
    getShapeProperties(pageId: string, shapeId: string): Record<string, ShapeData> {
        const parsed = this.cache.getParsed(pageId);
        const shape  = this.cache.getShapeMap(parsed).get(shapeId);
        if (!shape) throw new Error(`Shape ${shapeId} not found on page ${pageId}`);

        const result: Record<string, ShapeData> = {};
        if (!shape.Section) return result;

        const sections    = Array.isArray(shape.Section) ? shape.Section : [shape.Section];
        const propSection = sections.find((s: any) => s['@_N'] === SECTION_NAMES.Property);
        if (!propSection?.Row) return result;

        const rows = Array.isArray(propSection.Row) ? propSection.Row : [propSection.Row];
        for (const row of rows) {
            // Row names are "Prop.KeyName" — strip the prefix to recover the user-facing key.
            const rawKey: string = row['@_N'] ?? '';
            const key = rawKey.startsWith('Prop.') ? rawKey.slice(5) : rawKey;
            if (!key) continue;

            const cells: any[] = Array.isArray(row.Cell) ? row.Cell : (row.Cell ? [row.Cell] : []);
            const getCell = (n: string): string | undefined =>
                cells.find((c: any) => c['@_N'] === n)?.['@_V'];

            const rawValue = getCell('Value') ?? '';
            const type     = parseInt(getCell('Type') ?? '0') as VisioPropType;
            const label    = getCell('Label');
            const hidden   = getCell('Invisible') === '1';

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

    async addHyperlink(
        pageId: string,
        shapeId: string,
        details: { address?: string; subAddress?: string; description?: string },
    ): Promise<void> {
        const parsed = this.cache.getParsed(pageId);
        const shape  = this.cache.getShapeMap(parsed).get(shapeId);
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
        const newRow: any = { '@_N': `Hyperlink.Row_${nextIdx}`, Cell: [] };

        if (details.address    !== undefined) newRow.Cell.push({ '@_N': 'Address',     '@_V': details.address });
        if (details.subAddress !== undefined) newRow.Cell.push({ '@_N': 'SubAddress',   '@_V': details.subAddress });
        if (details.description !== undefined) newRow.Cell.push({ '@_N': 'Description', '@_V': details.description });
        newRow.Cell.push({ '@_N': 'NewWindow', '@_V': '0' });

        linkSection.Row.push(newRow);
        this.cache.saveParsed(pageId, parsed);
    }

    /**
     * Read back all hyperlinks attached to a shape.
     */
    getShapeHyperlinks(pageId: string, shapeId: string): ShapeHyperlink[] {
        const parsed = this.cache.getParsed(pageId);
        const shape  = this.cache.getShapeMap(parsed).get(shapeId);
        if (!shape) throw new Error(`Shape ${shapeId} not found on page ${pageId}`);

        const result: ShapeHyperlink[] = [];
        if (!shape.Section) return result;

        const sections    = Array.isArray(shape.Section) ? shape.Section : [shape.Section];
        const linkSection = sections.find((s: any) => s['@_N'] === 'Hyperlink');
        if (!linkSection?.Row) return result;

        const rows = Array.isArray(linkSection.Row) ? linkSection.Row : [linkSection.Row];
        for (const row of rows) {
            const cells: any[] = Array.isArray(row.Cell) ? row.Cell : (row.Cell ? [row.Cell] : []);
            const getCell = (n: string): string | undefined =>
                cells.find((c: any) => c['@_N'] === n)?.['@_V'];

            result.push({
                address:     getCell('Address'),
                subAddress:  getCell('SubAddress'),
                description: getCell('Description'),
                newWindow:   getCell('NewWindow') === '1',
            });
        }

        return result;
    }

    /**
     * Return the direct child shapes of a group or container shape.
     * Returns an empty array for non-group shapes or shapes with no children.
     */
    getShapeChildren(pageId: string, shapeId: string): VisioShape[] {
        const pagePath = this.cache.getPagePath(pageId);
        const reader   = new ShapeReader(this.pkg);
        return reader.readChildShapes(pagePath, shapeId);
    }
}

