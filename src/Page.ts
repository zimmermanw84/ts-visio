import { VisioPage, ConnectorStyle, PageOrientation, PageSizes, PageSizeName, ConnectionTarget, DrawingScaleInfo, LengthUnit } from './types/VisioTypes';
import { Connector } from './Connector';
import { VisioPackage } from './VisioPackage';
import { ShapeReader } from './ShapeReader';
import { ShapeModifier } from './ShapeModifier';
import { NewShapeProps } from './types/VisioTypes';
import { Shape } from './Shape';
import { MediaManager } from './core/MediaManager';
import { RelsManager } from './core/RelsManager';
import { createVisioShapeStub } from './utils/StubHelpers';
import { Layer } from './Layer';
import { TablePattern } from './diagrams/TablePattern';
import { SwimlanePattern } from './diagrams/SwimlanePattern';

/**
 * Represents a single page (tab) inside a Visio document.
 *
 * Obtain a `Page` instance from {@link VisioDocument.pages},
 * {@link VisioDocument.addPage}, or {@link VisioDocument.getPage}.
 *
 * @example
 * ```typescript
 * const doc  = await VisioDocument.create();
 * const page = doc.pages[0];
 * await page.addShape({ text: 'Hello', x: 1, y: 1, width: 2, height: 1 });
 * ```
 *
 * @category Pages
 */
export class Page {
    private media: MediaManager;
    private rels: RelsManager;
    private modifier: ShapeModifier;
    /** Resolved OPC part path for this page's XML file. */
    private pagePath: string;

    constructor(
        private internalPage: VisioPage,
        private pkg: VisioPackage,
        media?: MediaManager,
        rels?: RelsManager,
        modifier?: ShapeModifier
    ) {
        // Prefer the relationship-resolved path over the ID-derived fallback so
        // that loaded files with non-sequential page filenames work correctly.
        this.pagePath = internalPage.xmlPath ?? `visio/pages/page${internalPage.ID}.xml`;
        this.media = media || new MediaManager(pkg);
        this.rels = rels || new RelsManager(pkg);
        this.modifier = modifier || new ShapeModifier(pkg);
        this.modifier.registerPage(internalPage.ID, this.pagePath);
    }

    get id(): string {
        return this.internalPage.ID;
    }

    get name(): string {
        return this.internalPage.Name;
    }

    /** Width of the page canvas in inches. */
    get pageWidth(): number {
        return this.modifier.getPageDimensions(this.id).width;
    }

    /** Height of the page canvas in inches. */
    get pageHeight(): number {
        return this.modifier.getPageDimensions(this.id).height;
    }

    /** Current page orientation derived from the canvas dimensions. */
    get orientation(): PageOrientation {
        return this.pageWidth > this.pageHeight ? 'landscape' : 'portrait';
    }

    /**
     * Set the page canvas size in inches.
     * @example page.setSize(11, 8.5)  // landscape letter
     */
    setSize(width: number, height: number): this {
        this.modifier.setPageSize(this.id, width, height);
        return this;
    }

    /**
     * Convenience: change the page to a named standard size (portrait by default).
     * @example page.setNamedSize('A4')
     */
    setNamedSize(sizeName: PageSizeName, orientation: PageOrientation = 'portrait'): this {
        const { width, height } = PageSizes[sizeName];
        return orientation === 'landscape'
            ? this.setSize(height, width)
            : this.setSize(width, height);
    }

    /**
     * Rotate the canvas between portrait and landscape without changing the paper size.
     * Swaps width and height when the current orientation does not match the requested one.
     */
    setOrientation(orientation: PageOrientation): this {
        const w = this.pageWidth;
        const h = this.pageHeight;
        if (orientation === 'landscape' && h > w) {
            this.modifier.setPageSize(this.id, h, w);
        } else if (orientation === 'portrait' && w > h) {
            this.modifier.setPageSize(this.id, h, w);
        }
        return this;
    }

    /**
     * Return all top-level shapes on the page.
     * Group children are not included; use {@link findShapes} to search the entire shape tree.
     */
    getShapes(): Shape[] {
        const reader = new ShapeReader(this.pkg);
        try {
            const internalShapes = reader.readShapes(this.pagePath, this.modifier.getPageXml(this.id));
            return internalShapes.map(s => new Shape(s, this.id, this.pkg, this.modifier));
        } catch (e) {
            console.warn(`Could not read shapes for page ${this.id}:`, e);
            return [];
        }
    }

    /**
     * Find a shape by its ID anywhere on the page, including shapes nested inside groups.
     * Returns undefined if no shape with that ID exists.
     */
    getShapeById(id: string): Shape | undefined {
        const reader = new ShapeReader(this.pkg);
        const internal = reader.readShapeById(this.pagePath, id, this.modifier.getPageXml(this.id));
        if (!internal) return undefined;
        return new Shape(internal, this.id, this.pkg, this.modifier);
    }

    /**
     * Return all shapes on the page (including nested group children) that satisfy
     * the predicate. Equivalent to getAllShapes().filter(predicate).
     */
    findShapes(predicate: (shape: Shape) => boolean): Shape[] {
        const reader = new ShapeReader(this.pkg);
        const all = reader.readAllShapes(this.pagePath, this.modifier.getPageXml(this.id));
        return all
            .map(s => new Shape(s, this.id, this.pkg, this.modifier))
            .filter(predicate);
    }

    /**
     * Add a new shape to the page and return a {@link Shape} handle to it.
     *
     * @param props    Visual and text properties for the new shape.
     *                 All geometry fields (`x`, `y`, `width`, `height`) are in inches.
     * @param parentId Optional ID of an existing group shape to nest the new shape inside.
     *
     * @example
     * ```typescript
     * // Plain rectangle
     * const box = await page.addShape({ text: 'Server', x: 2, y: 3, width: 2, height: 1 });
     *
     * // Ellipse with styling
     * await page.addShape({
     *     text: 'Start', x: 1, y: 1, width: 1.5, height: 1.5,
     *     geometry: 'ellipse', fillColor: '#4472C4', fontColor: '#ffffff',
     * });
     *
     * // Master instance (shape defined by a reusable master)
     * const m = doc.createMaster('Router', 'ellipse');
     * await page.addShape({ text: '', x: 4, y: 2, width: 1, height: 1, masterId: m.id });
     * ```
     */
    async addShape(props: NewShapeProps, parentId?: string): Promise<Shape> {
        const newId = await this.modifier.addShape(this.id, props, parentId);

        // Return a fresh Shape object representing the new shape
        // We construct a minimal internal shape to satisfy the wrapper
        // In a real scenario, we might want to re-read the shape from disk to get full defaults
        const internalStub = createVisioShapeStub({
            ID: newId,
            Type: props.type,
            Text: props.text,
            Cells: {
                'Width': props.width,
                'Height': props.height,
                'PinX': props.x,
                'PinY': props.y,
                'LocPinX': props.width / 2,
                'LocPinY': props.height / 2
            }
        });

        return new Shape(internalStub, this.id, this.pkg, this.modifier);
    }
    /**
     * Return all connector shapes on the page.
     * Each `Connector` exposes the from/to shape IDs, port targets, line style, arrows,
     * and a `delete()` method to remove the connector.
     */
    getConnectors(): Connector[] {
        const reader = new ShapeReader(this.pkg);
        try {
            const data = reader.readConnectors(this.pagePath, this.modifier.getPageXml(this.id));
            return data.map(d => new Connector(d, this.id, this.modifier));
        } catch (e) {
            console.warn(`Could not read connectors for page ${this.id}:`, e);
            return [];
        }
    }

    /**
     * Draw a connector (line/arrow) between two shapes on this page.
     *
     * @param fromShape   Source shape.
     * @param toShape     Target shape.
     * @param beginArrow  Arrow head at the source end (use {@link ArrowHeads} constants).
     * @param endArrow    Arrow head at the target end (use {@link ArrowHeads} constants).
     * @param style       Line color, weight, pattern, and routing style.
     * @param fromPort    Named connection point on the source shape (e.g. `'Right'`).
     * @param toPort      Named connection point on the target shape (e.g. `'Left'`).
     *
     * @example
     * ```typescript
     * import { ArrowHeads } from 'ts-visio';
     * const a = await page.addShape({ text: 'A', x: 1, y: 1, width: 1, height: 1 });
     * const b = await page.addShape({ text: 'B', x: 4, y: 1, width: 1, height: 1 });
     * await page.connectShapes(a, b, ArrowHeads.None, ArrowHeads.OpenArrow,
     *     { lineColor: '#333333', routing: 'orthogonal' });
     * ```
     */
    async connectShapes(
        fromShape: Shape,
        toShape: Shape,
        beginArrow?: string,
        endArrow?: string,
        style?: ConnectorStyle,
        fromPort?: ConnectionTarget,
        toPort?: ConnectionTarget,
    ): Promise<void> {
        await this.modifier.addConnector(this.id, fromShape.id, toShape.id, beginArrow, endArrow, style, fromPort, toPort);
    }

    /**
     * Embed an image on the page and return the resulting Foreign shape.
     *
     * @param data   Raw image bytes (PNG, JPEG, GIF, BMP, or TIFF).
     * @param name   Filename used to store the image inside the archive (e.g. `'logo.png'`).
     * @param x      Horizontal pin position in inches.
     * @param y      Vertical pin position in inches.
     * @param width  Display width in inches.
     * @param height Display height in inches.
     *
     * @example
     * ```typescript
     * import fs from 'fs';
     * const imgBuffer = fs.readFileSync('./logo.png');
     * await page.addImage(imgBuffer, 'logo.png', 1, 1, 2, 1);
     * ```
     */
    async addImage(data: Buffer, name: string, x: number, y: number, width: number, height: number): Promise<Shape> {
        const mediaPath = this.media.addMedia(name, data);
        const rId = await this.rels.addImageRelationship(this.pagePath, mediaPath);
        const newId = await this.modifier.addShape(this.id, {
            text: '',
            x, y, width, height,
            type: 'Foreign',
            imgRelId: rId
        });

        const internalStub = createVisioShapeStub({
            ID: newId,
            Text: '',
            Cells: {
                'Width': width,
                'Height': height,
                'PinX': x,
                'PinY': y
            }
        });

        return new Shape(internalStub, this.id, this.pkg, this.modifier);
    }

    async addContainer(props: NewShapeProps): Promise<Shape> {
        const newId = await this.modifier.addContainer(this.id, props);

        const internalStub = createVisioShapeStub({
            ID: newId,
            Text: props.text,
            Cells: {
                'Width': props.width,
                'Height': props.height,
                'PinX': props.x,
                'PinY': props.y
            }
        });

        return new Shape(internalStub, this.id, this.pkg, this.modifier);
    }

    async addList(props: NewShapeProps, direction: 'vertical' | 'horizontal' = 'vertical'): Promise<Shape> {
        const newId = await this.modifier.addList(this.id, props, direction);

        const internalStub = createVisioShapeStub({
            ID: newId,
            Text: props.text,
            Cells: {
                'Width': props.width,
                'Height': props.height,
                'PinX': props.x,
                'PinY': props.y
            }
        });

        return new Shape(internalStub, this.id, this.pkg, this.modifier);
    }

    /** Creates a Swimlane Pool (a vertical List of Containers). */
    async addSwimlanePool(props: NewShapeProps): Promise<Shape> {
        return SwimlanePattern.addPool(this, props);
    }

    /** Creates a Swimlane Lane (a Container inside a pool) and attaches it to the pool. */
    async addSwimlaneLane(pool: Shape, props: NewShapeProps): Promise<Shape> {
        return SwimlanePattern.addLane(this, pool, props);
    }

    async addTable(x: number, y: number, title: string, columns: string[]): Promise<Shape> {
        return TablePattern.add(this, x, y, title, columns);
    }

    async addLayer(name: string, options?: { visible?: boolean, lock?: boolean, print?: boolean }): Promise<Layer> {
        const info = await this.modifier.addLayer(this.id, name, options);
        return new Layer(
            info.name, info.index, this.id, this.pkg, this.modifier,
            options?.visible ?? true,
            options?.lock ?? false,
            options?.print ?? true,
        );
    }

    /**
     * Return all layers defined on this page, ordered by index.
     * Works for both newly created documents and loaded `.vsdx` files.
     *
     * @example
     * const layers = page.getLayers();
     * // [{ name: 'Background', index: 0, visible: true, locked: false }, ...]
     */
    getLayers(): Layer[] {
        const infos = this.modifier.getPageLayers(this.id);
        return infos.map(
            l => new Layer(l.name, l.index, this.id, this.pkg, this.modifier, l.visible, l.locked, l.print)
        );
    }

    /**
     * Return the current drawing scale, or `null` if no custom scale is set (1:1).
     *
     * @example
     * const scale = page.getDrawingScale();
     * // { pageScale: 1, pageUnit: 'in', drawingScale: 10, drawingUnit: 'ft' }
     */
    getDrawingScale(): DrawingScaleInfo | null {
        return this.modifier.getDrawingScale(this.id);
    }

    /**
     * Set a custom drawing scale for the page.
     * One `pageScale` `pageUnit` on paper equals `drawingScale` `drawingUnit` in the real world.
     *
     * @example
     * // 1 inch on paper = 10 feet in the real world
     * page.setDrawingScale(1, 'in', 10, 'ft');
     *
     * // 1:100 metric
     * page.setDrawingScale(1, 'cm', 100, 'cm');
     */
    setDrawingScale(
        pageScale: number, pageUnit: LengthUnit,
        drawingScale: number, drawingUnit: LengthUnit
    ): this {
        this.modifier.setDrawingScale(this.id, pageScale, pageUnit, drawingScale, drawingUnit);
        return this;
    }

    /**
     * Remove any custom drawing scale, reverting the page to a 1:1 ratio.
     */
    clearDrawingScale(): this {
        this.modifier.clearDrawingScale(this.id);
        return this;
    }

    /** @internal Used by VisioDocument.renamePage() to keep in-memory state in sync. */
    _updateName(newName: string): void {
        this.internalPage.Name = newName;
    }
}
