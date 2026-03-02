import { VisioPackage } from './VisioPackage';
import { ShapeModifier } from './ShapeModifier';

/**
 * A named display layer on a Visio page.
 *
 * Layers control the visibility and printability of groups of shapes.
 * Obtain instances via `page.getLayers()` or `page.addLayer()`.
 *
 * @example
 * ```typescript
 * const annotations = await page.addLayer('Annotations');
 * ann.setVisible(false);  // hide all shapes on this layer
 * ```
 *
 * @category Layers
 */
export class Layer {
    private modifier: ShapeModifier | null;

    constructor(
        public name: string,
        public index: number,
        private pageId?: string,
        private pkg?: VisioPackage,
        modifier?: ShapeModifier,
        private _visible: boolean = true,
        private _locked: boolean = false,
    ) {
        this.modifier = modifier ?? (pkg ? new ShapeModifier(pkg) : null);
    }

    /** Whether the layer is currently visible. */
    get visible(): boolean {
        return this._visible;
    }

    /** Whether the layer is currently locked. */
    get locked(): boolean {
        return this._locked;
    }

    async setVisible(visible: boolean): Promise<this> {
        if (!this.pageId || !this.modifier) {
            throw new Error('Layer was not created with page context. Cannot update properties.');
        }
        await this.modifier.updateLayerProperty(this.pageId, this.index, 'Visible', visible ? '1' : '0');
        this._visible = visible;
        return this;
    }

    async setLocked(locked: boolean): Promise<this> {
        if (!this.pageId || !this.modifier) {
            throw new Error('Layer was not created with page context. Cannot update properties.');
        }
        await this.modifier.updateLayerProperty(this.pageId, this.index, 'Lock', locked ? '1' : '0');
        this._locked = locked;
        return this;
    }

    async hide(): Promise<this> {
        return this.setVisible(false);
    }

    async show(): Promise<this> {
        return this.setVisible(true);
    }

    /**
     * Rename this layer.
     */
    async rename(newName: string): Promise<this> {
        if (!this.pageId || !this.modifier) {
            throw new Error('Layer was not created with page context. Cannot update properties.');
        }
        await this.modifier.updateLayerProperty(this.pageId, this.index, 'Name', newName);
        this.name = newName;
        return this;
    }

    /**
     * Delete this layer from the page.
     * Removes the layer definition and strips it from all shape LayerMember cells.
     */
    async delete(): Promise<void> {
        if (!this.pageId || !this.modifier) {
            throw new Error('Layer was not created with page context. Cannot delete.');
        }
        this.modifier.deleteLayer(this.pageId, this.index);
    }
}
