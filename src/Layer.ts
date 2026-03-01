import { VisioPackage } from './VisioPackage';
import { ShapeModifier } from './ShapeModifier';

export class Layer {
    private modifier: ShapeModifier | null;

    constructor(
        public name: string,
        public index: number,
        private pageId?: string,
        private pkg?: VisioPackage,
        modifier?: ShapeModifier
    ) {
        this.modifier = modifier ?? (pkg ? new ShapeModifier(pkg) : null);
    }

    async setVisible(visible: boolean): Promise<this> {
        if (!this.pageId || !this.modifier) {
            throw new Error('Layer was not created with page context. Cannot update properties.');
        }
        await this.modifier.updateLayerProperty(this.pageId, this.index, 'Visible', visible ? '1' : '0');
        return this;
    }

    async setLocked(locked: boolean): Promise<this> {
        if (!this.pageId || !this.modifier) {
            throw new Error('Layer was not created with page context. Cannot update properties.');
        }
        await this.modifier.updateLayerProperty(this.pageId, this.index, 'Lock', locked ? '1' : '0');
        return this;
    }

    async hide(): Promise<this> {
        return this.setVisible(false);
    }

    async show(): Promise<this> {
        return this.setVisible(true);
    }
}
