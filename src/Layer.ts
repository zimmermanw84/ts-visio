import { VisioPackage } from './VisioPackage';
import { ShapeModifier } from './ShapeModifier';

export class Layer {
    constructor(
        public name: string,
        public index: number,
        private pageId?: string,
        private pkg?: VisioPackage
    ) { }

    async setVisible(visible: boolean): Promise<this> {
        if (!this.pageId || !this.pkg) {
            throw new Error('Layer was not created with page context. Cannot update properties.');
        }
        const modifier = new ShapeModifier(this.pkg);
        await modifier.updateLayerProperty(this.pageId, this.index, 'Visible', visible ? '1' : '0');
        return this;
    }

    async setLocked(locked: boolean): Promise<this> {
        if (!this.pageId || !this.pkg) {
            throw new Error('Layer was not created with page context. Cannot update properties.');
        }
        const modifier = new ShapeModifier(this.pkg);
        await modifier.updateLayerProperty(this.pageId, this.index, 'Lock', locked ? '1' : '0');
        return this;
    }

    async hide(): Promise<this> {
        return this.setVisible(false);
    }

    async show(): Promise<this> {
        return this.setVisible(true);
    }
}
