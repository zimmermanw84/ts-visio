import { VisioPage } from './types/VisioTypes';
import { VisioPackage } from './VisioPackage';
import { ShapeReader } from './ShapeReader';
import { ShapeModifier, NewShapeProps } from './ShapeModifier';
import { Shape } from './Shape';

export class Page {
    constructor(
        private internalPage: VisioPage,
        private pkg: VisioPackage
    ) { }

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

    async addShape(props: NewShapeProps): Promise<Shape> {
        const modifier = new ShapeModifier(this.pkg);
        const newId = await modifier.addShape(this.id, props);

        // Return a fresh Shape object representing the new shape
        // We construct a minimal internal shape to satisfy the wrapper
        // In a real scenario, we might want to re-read the shape from disk to get full defaults
        const internalStub: any = {
            ID: newId,
            Name: `Sheet.${newId}`,
            Text: props.text,
            Cells: {
                'Width': { V: props.width.toString() },
                'Height': { V: props.height.toString() }
            }
        };

        return new Shape(internalStub, this.id, this.pkg);
    }
    async connectShapes(fromShape: Shape, toShape: Shape): Promise<void> {
        const modifier = new ShapeModifier(this.pkg);
        await modifier.addConnector(this.id, fromShape.id, toShape.id);
    }
}
