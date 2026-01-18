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
    async connectShapes(fromShape: Shape, toShape: Shape, beginArrow?: string, endArrow?: string): Promise<void> {
        const modifier = new ShapeModifier(this.pkg);
        await modifier.addConnector(this.id, fromShape.id, toShape.id, beginArrow, endArrow);
    }

    async addTable(x: number, y: number, title: string, columns: string[]): Promise<string> {
        // Dimensions
        const width = 3;
        const headerHeight = 0.5;
        const lineItemHeight = 0.25;
        const bodyHeight = Math.max(0.5, columns.length * lineItemHeight + 0.1); // Min height

        // Math for stacking:
        // We want Header Bottom == Body Top
        // Header Center Y = hy
        // Body Center Y = by
        // Header Bottom = hy - (hh/2)
        // Body Top = by + (bh/2)
        // hy - hh/2 = by + bh/2  =>  by = hy - (hh + bh)/2

        const headerY = y;
        const bodyY = headerY - (headerHeight + bodyHeight) / 2;

        // 1. Header Shape
        // Grey background, bold text
        const headerShape = await this.addShape({
            text: title,
            x: x,
            y: headerY,
            width: width,
            height: headerHeight,
            fillColor: '#DDDDDD', // Light Grey
            bold: true
        });

        // 2. Body Shape
        // White background, list of columns
        const bodyText = columns.join('\n');
        await this.addShape({
            text: bodyText,
            x: x,
            y: bodyY,
            width: width,
            height: bodyHeight,
            fillColor: '#FFFFFF',
            fontColor: '#000000' // Explicit black
        });

        // Return ID of the "main" shape (Header)
        return headerShape.id;
    }
}
