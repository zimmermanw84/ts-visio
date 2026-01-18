import { VisioShape } from './types/VisioTypes';
import { VisioPackage } from './VisioPackage';
import { ShapeModifier, ShapeStyle } from './ShapeModifier';

export class Shape {
    constructor(
        private internalShape: VisioShape,
        private pageId: string,
        private pkg: VisioPackage
    ) { }

    get id(): string {
        return this.internalShape.ID;
    }

    get name(): string {
        return this.internalShape.Name;
    }

    get text(): string {
        return this.internalShape.Text || '';
    }

    async setText(newText: string): Promise<void> {
        const modifier = new ShapeModifier(this.pkg);
        await modifier.updateShapeText(this.pageId, this.id, newText);
        // Update local state to reflect change
        this.internalShape.Text = newText;
    }

    get width(): number {
        return this.internalShape.Cells['Width'] ? Number(this.internalShape.Cells['Width'].V) : 0;
    }

    get height(): number {
        return this.internalShape.Cells['Height'] ? Number(this.internalShape.Cells['Height'].V) : 0;
    }

    get x(): number {
        return this.internalShape.Cells['PinX'] ? Number(this.internalShape.Cells['PinX'].V) : 0;
    }

    get y(): number {
        return this.internalShape.Cells['PinY'] ? Number(this.internalShape.Cells['PinY'].V) : 0;
    }

    async connectTo(targetShape: Shape, beginArrow?: string, endArrow?: string): Promise<this> {
        const modifier = new ShapeModifier(this.pkg);
        await modifier.addConnector(this.pageId, this.id, targetShape.id, beginArrow, endArrow);
        return this;
    }

    async setStyle(style: ShapeStyle): Promise<this> {
        const modifier = new ShapeModifier(this.pkg);
        await modifier.updateShapeStyle(this.pageId, this.id, style);
        // Minimal local state update to reflect changes if necessary
        // For now, valid XML is the priority.
        return this;
    }
}
