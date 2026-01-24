import { VisioPage } from './types/VisioTypes';
import { VisioPackage } from './VisioPackage';
import { ShapeReader } from './ShapeReader';
import { ShapeModifier } from './ShapeModifier';
import { NewShapeProps } from './types/VisioTypes';
import { Shape } from './Shape';
import { MediaManager } from './core/MediaManager';
import { RelsManager } from './core/RelsManager';
import { createVisioShapeStub } from './utils/StubHelpers';

export class Page {
    private media: MediaManager;
    private rels: RelsManager;
    private modifier: ShapeModifier;

    constructor(
        private internalPage: VisioPage,
        private pkg: VisioPackage,
        media?: MediaManager,
        rels?: RelsManager,
        modifier?: ShapeModifier
    ) {
        this.media = media || new MediaManager(pkg);
        this.rels = rels || new RelsManager(pkg);
        this.modifier = modifier || new ShapeModifier(pkg);
    }

    get id(): string {
        return this.internalPage.ID;
    }

    get name(): string {
        return this.internalPage.Name;
    }

    getShapes(): Shape[] {
        const reader = new ShapeReader(this.pkg);
        // Assuming standard path mapping for now
        const pagePath = `visio/pages/page${this.id}.xml`;

        try {
            const internalShapes = reader.readShapes(pagePath);
            return internalShapes.map(s => new Shape(s, this.id, this.pkg));
        } catch (e) {
            // If page file doesn't exist or is empty, return empty array
            console.warn(`Could not read shapes for page ${this.id}:`, e);
            return [];
        }
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

        return new Shape(internalStub, this.id, this.pkg);
    }
    async connectShapes(fromShape: Shape, toShape: Shape, beginArrow?: string, endArrow?: string): Promise<void> {
        await this.modifier.addConnector(this.id, fromShape.id, toShape.id, beginArrow, endArrow);
    }

    async addImage(data: Buffer, name: string, x: number, y: number, width: number, height: number): Promise<Shape> {
        // 1. Upload Media
        const mediaPath = this.media.addMedia(name, data);

        // 2. Link Page to Media
        const rId = await this.rels.addPageImageRel(this.id, mediaPath);

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

        return new Shape(internalStub, this.id, this.pkg);
    }

    async addTable(x: number, y: number, title: string, columns: string[]): Promise<Shape> {
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

        // Make it a Container so it expands with content
        groupShape.convertToContainer();

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
}
