import { VisioShape, ConnectorStyle, ConnectionTarget, ConnectionPointDef } from './types/VisioTypes';
import { VisioPackage } from './VisioPackage';
import { ShapeModifier, ShapeStyle } from './ShapeModifier';
import { VisioPropType } from './types/VisioTypes';
import { Layer } from './Layer';
import { SHAPE_TYPES, STRUCT_RELATIONSHIP_TYPES } from './core/VisioConstants';

export interface ShapeData {
    value: string | number | boolean | Date;
    label?: string;
    hidden?: boolean;
    type?: VisioPropType;
}

export interface ShapeHyperlink {
    address?: string;
    subAddress?: string;
    description?: string;
    newWindow: boolean;
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

    /**
     * The shape's Type attribute — `'Group'` for group shapes, `'Shape'` (or `undefined`
     * normalised to `'Shape'`) for regular shapes.
     */
    get type(): string {
        return this.internalShape.Type ?? 'Shape';
    }

    /**
     * `true` if this shape is a Group (i.e. it can contain nested child shapes).
     * Use `shape.getChildren()` to retrieve those children.
     */
    get isGroup(): boolean {
        return this.internalShape.Type === SHAPE_TYPES.Group;
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

    async delete(): Promise<void> {
        await this.modifier.deleteShape(this.pageId, this.id);
    }

    async connectTo(
        targetShape: Shape,
        beginArrow?: string,
        endArrow?: string,
        style?: ConnectorStyle,
        fromPort?: ConnectionTarget,
        toPort?: ConnectionTarget,
    ): Promise<this> {
        await this.modifier.addConnector(this.pageId, this.id, targetShape.id, beginArrow, endArrow, style, fromPort, toPort);
        return this;
    }

    /**
     * Add a connection point to this shape.
     * Returns the zero-based index (IX) of the newly added point.
     */
    addConnectionPoint(point: ConnectionPointDef): number {
        return this.modifier.addConnectionPoint(this.pageId, this.id, point);
    }

    /**
     * Apply a document-level stylesheet to this shape.
     * Create styles via `doc.createStyle()` and pass the returned `id`.
     *
     * @param styleId  The stylesheet ID to apply.
     * @param which    `'all'` (default) applies to line, fill, and text;
     *                 `'line'`, `'fill'`, or `'text'` applies to only that category.
     */
    applyStyle(styleId: number, which: 'all' | 'line' | 'fill' | 'text' = 'all'): this {
        this.modifier.applyStyle(this.pageId, this.id, styleId, which);
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

    /**
     * Read back all custom property (shape data) entries written to this shape.
     * Returns a map of property key → ShapeData. Values are coerced to the
     * declared Visio type (Number, Boolean, Date, or String).
     */
    getProperties(): Record<string, ShapeData> {
        return this.modifier.getShapeProperties(this.pageId, this.id);
    }

    /**
     * Read back all hyperlinks attached to this shape.
     */
    getHyperlinks(): ShapeHyperlink[] {
        return this.modifier.getShapeHyperlinks(this.pageId, this.id);
    }

    /**
     * Read back the layer indices this shape is assigned to.
     * Returns an empty array if the shape has no layer assignment.
     */
    getLayerIndices(): number[] {
        return this.modifier.getShapeLayerIndices(this.pageId, this.id);
    }

    /**
     * Return the direct child shapes of this group.
     * Returns an empty array for non-group shapes or groups with no children.
     *
     * Only direct children are returned — grandchildren are accessible by calling
     * `getChildren()` on the child shape.
     *
     * @example
     * const group = await page.addShape({ text: 'G', x: 5, y: 5, width: 4, height: 4, type: 'Group' });
     * await page.addShape({ text: 'Child A', x: 1, y: 1, width: 1, height: 1 }, group.id);
     * await page.addShape({ text: 'Child B', x: 2, y: 1, width: 1, height: 1 }, group.id);
     * group.getChildren(); // → [Shape('Child A'), Shape('Child B')]
     */
    getChildren(): Shape[] {
        const children = this.modifier.getShapeChildren(this.pageId, this.id);
        return children.map(c => new Shape(c, this.pageId, this.pkg, this.modifier));
    }

    /** Current rotation angle in degrees (0 if no Angle cell is set). */
    get angle(): number {
        const cell = this.internalShape.Cells['Angle'];
        return cell ? (parseFloat(cell.V) * 180) / Math.PI : 0;
    }

    /**
     * Rotate the shape to an absolute angle (degrees, clockwise).
     * Replaces any existing rotation.
     */
    async rotate(degrees: number): Promise<this> {
        await this.modifier.rotateShape(this.pageId, this.id, degrees);
        const radians = (degrees * Math.PI) / 180;
        this.setLocalCoord('Angle', radians);
        return this;
    }

    /**
     * Resize the shape to the given width and height (in inches).
     * Updates LocPinX/LocPinY to keep the centre-pin at width/2, height/2.
     */
    async resize(width: number, height: number): Promise<this> {
        if (width <= 0 || height <= 0) throw new Error('Shape dimensions must be positive');
        await this.modifier.resizeShape(this.pageId, this.id, width, height);
        this.setLocalCoord('Width', width);
        this.setLocalCoord('Height', height);
        this.setLocalCoord('LocPinX', width / 2);
        this.setLocalCoord('LocPinY', height / 2);
        return this;
    }

    /**
     * Flip the shape horizontally. Pass `false` to un-flip.
     */
    async flipX(enabled: boolean = true): Promise<this> {
        this.modifier.setShapeFlip(this.pageId, this.id, 'x', enabled);
        this.setLocalRawCell('FlipX', enabled ? '1' : '0');
        return this;
    }

    /**
     * Flip the shape vertically. Pass `false` to un-flip.
     */
    async flipY(enabled: boolean = true): Promise<this> {
        this.modifier.setShapeFlip(this.pageId, this.id, 'y', enabled);
        this.setLocalRawCell('FlipY', enabled ? '1' : '0');
        return this;
    }

    async addMember(memberShape: Shape): Promise<this> {
        await this.modifier.addRelationship(this.pageId, this.id, memberShape.id, STRUCT_RELATIONSHIP_TYPES.Container);
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

    private setLocalRawCell(name: string, value: string): void {
        if (this.internalShape.Cells[name]) this.internalShape.Cells[name].V = value;
        else this.internalShape.Cells[name] = { V: value, N: name };
    }
}
