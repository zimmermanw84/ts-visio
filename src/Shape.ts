import { VisioShape } from './types/VisioTypes';
import { VisioPackage } from './VisioPackage';
import { ShapeModifier, ShapeStyle } from './ShapeModifier';
import { VisioPropType } from './types/VisioTypes';

export interface ShapeData {
    value: string | number | boolean | Date;
    label?: string;
    hidden?: boolean;
    type?: VisioPropType;
}

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

    async placeRightOf(targetShape: Shape, options: { gap: number } = { gap: 1 }): Promise<this> {
        const newX = targetShape.x + targetShape.width + options.gap;
        const newY = targetShape.y; // Keep same Y (per instructions, aligns center-to-center if PinY is center)

        const modifier = new ShapeModifier(this.pkg);
        await modifier.updateShapePosition(this.pageId, this.id, newX, newY);

        // Update local state is crucial for chaining successive placements
        if (this.internalShape.Cells['PinX']) this.internalShape.Cells['PinX'].V = newX.toString();
        else this.internalShape.Cells['PinX'] = { V: newX.toString(), N: 'PinX' };

        if (this.internalShape.Cells['PinY']) this.internalShape.Cells['PinY'].V = newY.toString();
        else this.internalShape.Cells['PinY'] = { V: newY.toString(), N: 'PinY' };

        return this;
    }

    async placeBelow(targetShape: Shape, options: { gap: number } = { gap: 1 }): Promise<this> {
        const newX = targetShape.x; // Align Centers
        // Target Bottom = target.y - target.height / 2
        // My Top = Target Bottom - gap
        // My Center = My Top - my.height / 2
        // My Center = target.y - target.height/2 - gap - my.height/2
        const newY = targetShape.y - (targetShape.height + this.height) / 2 - options.gap;

        const modifier = new ShapeModifier(this.pkg);
        await modifier.updateShapePosition(this.pageId, this.id, newX, newY);

        if (this.internalShape.Cells['PinX']) this.internalShape.Cells['PinX'].V = newX.toString();
        else this.internalShape.Cells['PinX'] = { V: newX.toString(), N: 'PinX' };

        if (this.internalShape.Cells['PinY']) this.internalShape.Cells['PinY'].V = newY.toString();
        else this.internalShape.Cells['PinY'] = { V: newY.toString(), N: 'PinY' };

        return this;
    }

    async addPropertyDefinition(name: string, type: number, options: { label?: string, invisible?: boolean } = {}): Promise<this> {
        const modifier = new ShapeModifier(this.pkg);
        await modifier.addPropertyDefinition(this.pageId, this.id, name, type, options);
        return this;
    }

    async setPropertyValue(name: string, value: string | number | boolean | Date): Promise<this> {
        const modifier = new ShapeModifier(this.pkg);
        await modifier.setPropertyValue(this.pageId, this.id, name, value);
        return this;
    }

    async addData(key: string, data: ShapeData): Promise<this> {
        // Auto-detect type if not provided
        let type = data.type;
        if (type === undefined) {
            if (data.value instanceof Date) {
                type = VisioPropType.Date;
            } else if (typeof data.value === 'number') {
                type = VisioPropType.Number;
            } else if (typeof data.value === 'boolean') {
                type = VisioPropType.Boolean;
            } else {
                type = VisioPropType.String;
            }
        }

        // 1. Define Property
        await this.addPropertyDefinition(key, type, {
            label: data.label,
            invisible: data.hidden
        });

        // 2. Set Value
        await this.setPropertyValue(key, data.value);

        return this;
    }
}
