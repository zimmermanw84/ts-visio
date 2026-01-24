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

    addPropertyDefinition(name: string, type: number, options: { label?: string, invisible?: boolean } = {}): this {
        const modifier = new ShapeModifier(this.pkg);
        modifier.addPropertyDefinition(this.pageId, this.id, name, type, options);
        return this;
    }

    setPropertyValue(name: string, value: string | number | boolean | Date): this {
        const modifier = new ShapeModifier(this.pkg);
        modifier.setPropertyValue(this.pageId, this.id, name, value);
        return this;
    }

    addData(key: string, data: ShapeData): this {
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
        this.addPropertyDefinition(key, type, {
            label: data.label,
            invisible: data.hidden
        });

        // 2. Set Value
        this.setPropertyValue(key, data.value);

        return this;
    }

    async addMember(memberShape: Shape): Promise<this> {
        const modifier = new ShapeModifier(this.pkg);
        // Type="Container" is the standard for Container relationships
        await modifier.addRelationship(this.pageId, this.id, memberShape.id, 'Container');
        return this;
    }

    async resizeToFit(padding: number = 0.25): Promise<this> {
        const modifier = new ShapeModifier(this.pkg);
        const memberIds = modifier.getContainerMembers(this.pageId, this.id);

        if (memberIds.length === 0) return this;

        // Calculate Bounding Box
        let minX = Infinity, minY = Infinity;
        let maxX = -Infinity, maxY = -Infinity;

        for (const mid of memberIds) {
            const geo = modifier.getShapeGeometry(this.pageId, mid);
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
        await modifier.updateShapePosition(this.pageId, this.id, newPinX, newPinY);
        // Need to update Width/Height as well (helper currently only does PinX/PinY)
        // We'll use setPropertyValue or just access internal structure via updateShapePosition-like logic
        // Actually, updateShapePosition in modifier is specific. Let's add updateShapeSize to modifier or extend logic?
        // For Speed, let's just use what we have or add a quick updateSize helper.

        // Wait, ShapeModifier.updateShapePosition only does PinX/PinY.
        // We need to update Width/Height. I'll add a helper method to Shape.ts for now or use the generic property setter if applicable (it's not, Size is Cells, not Props).

        // Let's implement updateShapeSize in modifier for completeness or just do it here cleanly.
        // I will use a custom implementation here or add it to Modifier. Adding to modifier is cleaner.
        // I'll pause and add `updateShapeSize` to modifier first? No, I can't interrupt this tool call.
        // I will implement the calls assuming updateShapeSize exists, or access the lower level if possible.
        // Actually, I can add `updateShapeSize` to Modifier in the next step or right now if I could.
        // Since I'm in `Shape.ts`, I'll assume I'll add `updateShapeSize` to modifier next.
        // BUT, I can't run code that doesn't exist.
        // I'll call a new method `updateShapeDimensions` on modifier and implement it immediately after.

        await modifier.updateShapeDimensions(this.pageId, this.id, newWidth, newHeight);

        // Update Z-Order (Send to Back)
        await modifier.reorderShape(this.pageId, this.id, 'back');

        // Update local state to reflect changes (crucial for tests/chaining)
        const updateLocalCell = (n: string, v: string) => {
            if (this.internalShape.Cells[n]) this.internalShape.Cells[n].V = v;
            else this.internalShape.Cells[n] = { V: v, N: n };
        };

        updateLocalCell('PinX', newPinX.toString());
        updateLocalCell('PinY', newPinY.toString());
        updateLocalCell('Width', newWidth.toString());
        updateLocalCell('Height', newHeight.toString());

        return this;
    }

}
