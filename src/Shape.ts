import { VisioShape } from './types/VisioTypes';
import { VisioPackage } from './VisioPackage';
import { ShapeModifier, ShapeStyle } from './ShapeModifier';
import { VisioPropType } from './types/VisioTypes';
import { Layer } from './Layer';

export interface ShapeData {
    value: string | number | boolean | Date;
    label?: string;
    hidden?: boolean;
    type?: VisioPropType;
}

/** Round a coordinate to 10 decimal places to prevent float-to-string-to-float precision drift. */
function fmtCoord(n: number): string {
    return parseFloat(n.toFixed(10)).toString();
}

export class Shape {
    private modifier: ShapeModifier;

    constructor(
        private internalShape: VisioShape,
        private pageId: string,
        private pkg: VisioPackage,
        modifier?: ShapeModifier
    ) {
        this.modifier = modifier ?? new ShapeModifier(pkg);
    }

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
        await this.modifier.updateShapeText(this.pageId, this.id, newText);
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
        await this.modifier.addConnector(this.pageId, this.id, targetShape.id, beginArrow, endArrow);
        return this;
    }

    async setStyle(style: ShapeStyle): Promise<this> {
        await this.modifier.updateShapeStyle(this.pageId, this.id, style);
        return this;
    }

    async placeRightOf(targetShape: Shape, options: { gap: number } = { gap: 1 }): Promise<this> {
        // PinX is the shape centre, so right edge of target = target.x + target.width/2;
        // left edge of this = newX - this.width/2. Set left edge = right edge of target + gap.
        const newX = targetShape.x + (targetShape.width / 2) + options.gap + (this.width / 2);
        const newY = targetShape.y; // Align centres vertically

        await this.modifier.updateShapePosition(this.pageId, this.id, newX, newY);

        // Update local state — rounded to avoid float precision drift in chained placements
        this.setLocalCoord('PinX', newX);
        this.setLocalCoord('PinY', newY);

        return this;
    }

    async placeBelow(targetShape: Shape, options: { gap: number } = { gap: 1 }): Promise<this> {
        const newX = targetShape.x; // Align centres horizontally
        // Target bottom edge = target.y - target.height/2
        // This centre = target bottom - gap - this.height/2
        const newY = targetShape.y - (targetShape.height + this.height) / 2 - options.gap;

        await this.modifier.updateShapePosition(this.pageId, this.id, newX, newY);

        this.setLocalCoord('PinX', newX);
        this.setLocalCoord('PinY', newY);

        return this;
    }

    addPropertyDefinition(name: string, type: number, options: { label?: string, invisible?: boolean } = {}): this {
        this.modifier.addPropertyDefinition(this.pageId, this.id, name, type, options);
        return this;
    }

    setPropertyValue(name: string, value: string | number | boolean | Date): this {
        this.modifier.setPropertyValue(this.pageId, this.id, name, value);
        return this;
    }

    addData(key: string, data: ShapeData): this {
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
        this.addPropertyDefinition(key, type, { label: data.label, invisible: data.hidden });
        this.setPropertyValue(key, data.value);
        return this;
    }

    async addMember(memberShape: Shape): Promise<this> {
        await this.modifier.addRelationship(this.pageId, this.id, memberShape.id, 'Container');
        return this;
    }

    async addListItem(item: Shape): Promise<this> {
        await this.modifier.addListItem(this.pageId, this.id, item.id);
        await this.refreshLocalState();
        return this;
    }

    async resizeToFit(padding: number = 0.25): Promise<this> {
        await this.modifier.resizeContainerToFit(this.pageId, this.id, padding);
        await this.refreshLocalState();
        return this;
    }

    private async refreshLocalState() {
        const geo = this.modifier.getShapeGeometry(this.pageId, this.id);
        this.setLocalCoord('PinX', geo.x);
        this.setLocalCoord('PinY', geo.y);
        this.setLocalCoord('Width', geo.width);
        this.setLocalCoord('Height', geo.height);
    }

    async addHyperlink(address: string, description?: string): Promise<this> {
        await this.modifier.addHyperlink(this.pageId, this.id, { address, description });
        return this;
    }

    async linkToPage(targetPage: { name: string }, description?: string): Promise<this> {
        await this.modifier.addHyperlink(this.pageId, this.id, {
            address: '',
            subAddress: targetPage.name,
            description
        });
        return this;
    }

    /**
     * Adds an external hyperlink to the shape.
     * Shows up in the right-click menu in Visio.
     * @param url External URL (e.g. https://google.com)
     * @param description Text to show in the menu
     */
    async toUrl(url: string, description?: string): Promise<this> {
        return this.addHyperlink(url, description);
    }

    /**
     * Adds an internal link to another page.
     * Shows up in the right-click menu in Visio.
     * @param targetPage The Page object to link to
     * @param description Text to show in the menu
     */
    async toPage(targetPage: { name: string }, description?: string): Promise<this> {
        return this.linkToPage(targetPage, description);
    }

    async assignLayer(layer: Layer | number): Promise<this> {
        const index = typeof layer === 'number' ? layer : layer.index;
        await this.modifier.assignLayer(this.pageId, this.id, index);
        return this;
    }

    /**
     * Alias for assignLayer. Adds this shape to a layer.
     */
    async addToLayer(layer: Layer | number): Promise<this> {
        return this.assignLayer(layer);
    }

    private setLocalCoord(name: string, value: number): void {
        const v = fmtCoord(value);
        if (this.internalShape.Cells[name]) this.internalShape.Cells[name].V = v;
        else this.internalShape.Cells[name] = { V: v, N: name };
    }
}
