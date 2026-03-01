import { VisioPage, ConnectorStyle, PageOrientation, PageSizes, PageSizeName, ConnectionTarget } from './types/VisioTypes';
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

    getShapes(): Shape[] {
        const reader = new ShapeReader(this.pkg);
        try {
            const internalShapes = reader.readShapes(this.pagePath);
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
        const internal = reader.readShapeById(this.pagePath, id);
        if (!internal) return undefined;
        return new Shape(internal, this.id, this.pkg, this.modifier);
    }

    /**
     * Return all shapes on the page (including nested group children) that satisfy
     * the predicate. Equivalent to getAllShapes().filter(predicate).
     */
    findShapes(predicate: (shape: Shape) => boolean): Shape[] {
        const reader = new ShapeReader(this.pkg);
        const all = reader.readAllShapes(this.pagePath);
        return all
            .map(s => new Shape(s, this.id, this.pkg, this.modifier))
            .filter(predicate);
    }

    async addShape(props: NewShapeProps, parentId?: string): Promise<Shape> {
        const newId = await this.modifier.addShape(this.id, props, parentId);

        // Return a fresh Shape object representing the new shape
        // We construct a minimal internal shape to satisfy the wrapper
        // In a real scenario, we might want to re-read the shape from disk to get full defaults
        const internalStub = createVisioShapeStub({
            ID: newId,
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
            const data = reader.readConnectors(this.pagePath);
            return data.map(d => new Connector(d, this.id, this.modifier));
        } catch (e) {
            console.warn(`Could not read connectors for page ${this.id}:`, e);
            return [];
        }
    }

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

    async addImage(data: Buffer, name: string, x: number, y: number, width: number, height: number): Promise<Shape> {
        // 1. Upload Media
        const mediaPath = this.media.addMedia(name, data);

        // 2. Link Page to Media (use resolved path so loaded files work correctly)
        const rId = await this.rels.addImageRelationship(this.pagePath, mediaPath);

        // 3. Create Shape
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

    /**
     * Creates a Swimlane Pool (which is technically a Vertical List of Containers).
     * @param props Visual properties
     */
    async addSwimlanePool(props: NewShapeProps): Promise<Shape> {
        // A Pool is just a vertical list
        return this.addList(props, 'vertical');
    }

    /**
     * Creates a Swimlane Lane (which is technically a Container).
     * @param props Visual properties
     */
    async addSwimlaneLane(props: NewShapeProps): Promise<Shape> {
        // A Lane is just a container
        return this.addContainer(props);
    }

    async addTable(x: number, y: number, title: string, columns: string[]): Promise<Shape> {
        // ... (previous implementation)
        // Dimensions
        const width = 3;
        const headerHeight = 0.5;
        const lineItemHeight = 0.25;
        const bodyHeight = Math.max(0.5, columns.length * lineItemHeight + 0.1); // Min height
        const totalHeight = headerHeight + bodyHeight;

        // 1. Create Main Group Shape (Transparent container)
        // Group Logic:
        // Positioned at (x, y) on the Page.
        // Size encapsulates both header and body.
        const groupShape = await this.addShape({
            text: '', // No text on container
            x: x,
            y: y,
            width: width,
            height: totalHeight,
            type: 'Group'
        });

        // 2. Header Shape (Inside Group)
        // Coords relative to Group (Bottom-Left is 0,0)
        // Header is at top. Center X is Width/2.
        // Center Y = BodyHeight + (HeaderHeight/2)
        const headerCenterY = bodyHeight + (headerHeight / 2);

        await this.addShape({
            text: title,
            x: width / 2, // Relative PinX
            y: headerCenterY, // Relative PinY
            width: width,
            height: headerHeight,
            fillColor: '#DDDDDD',
            bold: true
        }, groupShape.id);

        // 3. Body Shape (Inside Group)
        // Bottom part. Center X is Width/2.
        // Center Y = BodyHeight/2
        const bodyCenterY = bodyHeight / 2;
        const bodyText = columns.join('\n');

        await this.addShape({
            text: bodyText,
            x: width / 2, // Relative PinX
            y: bodyCenterY, // Relative PinY
            width: width,
            height: bodyHeight,
            fillColor: '#FFFFFF',
            fontColor: '#000000'
        }, groupShape.id);

        // Return the Group Shape
        return groupShape;
    }

    async addLayer(name: string, options?: { visible?: boolean, lock?: boolean, print?: boolean }): Promise<Layer> {
        const info = await this.modifier.addLayer(this.id, name, options);
        return new Layer(info.name, info.index, this.id, this.pkg, this.modifier);
    }
}
