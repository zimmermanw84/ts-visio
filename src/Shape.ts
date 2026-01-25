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

        const modifier = new ShapeModifier(this.pkg);
        modifier.autoSave = false;

        // 1. Define Property
        modifier.addPropertyDefinition(this.pageId, this.id, key, type, {
            label: data.label,
            invisible: data.hidden
        });

        // 2. Set Value
        modifier.setPropertyValue(this.pageId, this.id, key, data.value);

        modifier.flush();

        return this;
    }

    async addMember(memberShape: Shape): Promise<this> {
        const modifier = new ShapeModifier(this.pkg);
        // Type="Container" is the standard for Container relationships
        await modifier.addRelationship(this.pageId, this.id, memberShape.id, 'Container');
        return this;
    }

    async addListItem(item: Shape): Promise<this> {
        const modifier = new ShapeModifier(this.pkg);
        await modifier.addListItem(this.pageId, this.id, item.id);

        // Refresh local state after modifer updates (resizeToFit called internally)
        await this.refreshLocalState();
        return this;
    }

    async resizeToFit(padding: number = 0.25): Promise<this> {
        const modifier = new ShapeModifier(this.pkg);
        await modifier.resizeContainerToFit(this.pageId, this.id, padding);

        await this.refreshLocalState();
        return this;
    }

    private async refreshLocalState() {
        // Reloads internal Cells from modifier's fresh XML
        // This is a bit expensive but ensures consistency
        const modifier = new ShapeModifier(this.pkg);
        const geo = modifier.getShapeGeometry(this.pageId, this.id);

        const update = (n: string, v: string) => {
            if (this.internalShape.Cells[n]) this.internalShape.Cells[n].V = v;
            else this.internalShape.Cells[n] = { V: v, N: n };
        };

        update('PinX', geo.x.toString());
        update('PinY', geo.y.toString());
        update('Width', geo.width.toString());
        update('Height', geo.height.toString());
    }

    async addHyperlink(address: string, description?: string): Promise<this> {
        const modifier = new ShapeModifier(this.pkg);
        await modifier.addHyperlink(this.pageId, this.id, { address, description });
        return this;
    }

    async linkToPage(targetPage: { name: string }, description?: string): Promise<this> {
        const modifier = new ShapeModifier(this.pkg);
        // Internal links use SubAddress='PageName' and empty Address
        await modifier.addHyperlink(this.pageId, this.id, {
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
        const modifier = new ShapeModifier(this.pkg);
        await modifier.assignLayer(this.pageId, this.id, index);
        return this;
    }

    /**
     * Alias for assignLayer. Adds this shape to a layer.
     */
    async addToLayer(layer: Layer | number): Promise<this> {
        return this.assignLayer(layer);
    }
}
