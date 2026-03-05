import { describe, it, expect } from 'vitest';
import { VisioDocument } from '../src/VisioDocument';

async function freshDoc() {
    const doc  = await VisioDocument.create();
    const page = doc.pages[0];
    return { doc, page };
}

// ── page.getLayers() — basic ───────────────────────────────────────────────────

describe('page.getLayers() — basic', () => {
    it('returns empty array when no layers have been added', async () => {
        const { page } = await freshDoc();
        expect(page.getLayers()).toEqual([]);
    });

    it('returns one layer after addLayer()', async () => {
        const { page } = await freshDoc();
        await page.addLayer('Background');
        const layers = page.getLayers();
        expect(layers).toHaveLength(1);
        expect(layers[0].name).toBe('Background');
        expect(layers[0].index).toBe(0);
    });

    it('returns all layers in index order', async () => {
        const { page } = await freshDoc();
        await page.addLayer('A');
        await page.addLayer('B');
        await page.addLayer('C');
        const layers = page.getLayers();
        expect(layers).toHaveLength(3);
        expect(layers.map(l => l.name)).toEqual(['A', 'B', 'C']);
        expect(layers.map(l => l.index)).toEqual([0, 1, 2]);
    });

    it('layer.visible reflects the visible option', async () => {
        const { page } = await freshDoc();
        await page.addLayer('Hidden', { visible: false });
        await page.addLayer('Shown',  { visible: true  });
        const [hidden, shown] = page.getLayers();
        expect(hidden.visible).toBe(false);
        expect(shown.visible).toBe(true);
    });

    it('layer.locked reflects the lock option', async () => {
        const { page } = await freshDoc();
        await page.addLayer('Locked',   { lock: true  });
        await page.addLayer('Unlocked', { lock: false });
        const [locked, unlocked] = page.getLayers();
        expect(locked.locked).toBe(true);
        expect(unlocked.locked).toBe(false);
    });
});

// ── page.getLayers() — round-trip ─────────────────────────────────────────────

describe('page.getLayers() — round-trip save/reload', () => {
    it('layers survive save and reload', async () => {
        const { doc, page } = await freshDoc();
        await page.addLayer('Fore');
        await page.addLayer('Back', { visible: false, lock: true });

        const buf   = await doc.save();
        const doc2  = await VisioDocument.load(buf);
        const page2 = doc2.pages[0];
        const layers = page2.getLayers();

        expect(layers).toHaveLength(2);
        expect(layers[0].name).toBe('Fore');
        expect(layers[0].visible).toBe(true);
        expect(layers[0].locked).toBe(false);
        expect(layers[1].name).toBe('Back');
        expect(layers[1].visible).toBe(false);
        expect(layers[1].locked).toBe(true);
    });
});

// ── getLayers() returns functional Layer instances ────────────────────────────

describe('layers returned by getLayers() are fully functional', () => {
    it('setVisible() persists through getLayers()', async () => {
        const { page } = await freshDoc();
        await page.addLayer('Test');
        const [layer] = page.getLayers();
        await layer.setVisible(false);
        expect(page.getLayers()[0].visible).toBe(false);
    });

    it('setLocked() persists through getLayers()', async () => {
        const { page } = await freshDoc();
        await page.addLayer('Test');
        const [layer] = page.getLayers();
        await layer.setLocked(true);
        expect(page.getLayers()[0].locked).toBe(true);
    });

    it('hide() and show() toggle visible correctly', async () => {
        const { page } = await freshDoc();
        await page.addLayer('Toggle');
        const [layer] = page.getLayers();
        await layer.hide();
        expect(layer.visible).toBe(false);
        await layer.show();
        expect(layer.visible).toBe(true);
    });
});

// ── layer.rename() ────────────────────────────────────────────────────────────

describe('layer.rename()', () => {
    it('updates the layer name in-memory', async () => {
        const { page } = await freshDoc();
        await page.addLayer('Old');
        const [layer] = page.getLayers();
        await layer.rename('New');
        expect(layer.name).toBe('New');
    });

    it('rename persists after re-reading via getLayers()', async () => {
        const { page } = await freshDoc();
        await page.addLayer('Before');
        const [layer] = page.getLayers();
        await layer.rename('After');
        expect(page.getLayers()[0].name).toBe('After');
    });

    it('rename survives save/reload', async () => {
        const { doc, page } = await freshDoc();
        await page.addLayer('Original');
        const [layer] = page.getLayers();
        await layer.rename('Renamed');

        const buf  = await doc.save();
        const doc2 = await VisioDocument.load(buf);
        expect(doc2.pages[0].getLayers()[0].name).toBe('Renamed');
    });
});

// ── layer.delete() ────────────────────────────────────────────────────────────

describe('layer.delete()', () => {
    it('removes the layer from getLayers()', async () => {
        const { page } = await freshDoc();
        await page.addLayer('Keep');
        await page.addLayer('Remove');
        const [, remove] = page.getLayers();
        await remove.delete();
        const remaining = page.getLayers();
        expect(remaining).toHaveLength(1);
        expect(remaining[0].name).toBe('Keep');
    });

    it('deleting the only layer leaves an empty list', async () => {
        const { page } = await freshDoc();
        await page.addLayer('Solo');
        const [layer] = page.getLayers();
        await layer.delete();
        expect(page.getLayers()).toHaveLength(0);
    });

    it('removes the deleted layer index from shape LayerMember cells', async () => {
        const { page } = await freshDoc();
        const layerA = await page.addLayer('A');
        const layerB = await page.addLayer('B');
        const shape  = await page.addShape({ text: 'S', x: 1, y: 1, width: 1, height: 1 });
        await shape.addToLayer(layerA);
        await shape.addToLayer(layerB);

        // Confirm both indices are present
        expect(shape.getLayerIndices()).toContain(0);
        expect(shape.getLayerIndices()).toContain(1);

        // Delete layer A (index 0). Re-indexing shifts B from 1 → 0.
        await layerA.delete();

        // After re-indexing: B is now at index 0; old index 1 is gone
        const indices = shape.getLayerIndices();
        expect(indices).toContain(0);    // B shifted from 1 → 0
        expect(indices).not.toContain(1); // old index 1 no longer exists
    });

    it('deletion survives save/reload', async () => {
        const { doc, page } = await freshDoc();
        await page.addLayer('Stay');
        await page.addLayer('Go');
        const [, go] = page.getLayers();
        await go.delete();

        const buf  = await doc.save();
        const doc2 = await VisioDocument.load(buf);
        const layers = doc2.pages[0].getLayers();
        expect(layers).toHaveLength(1);
        expect(layers[0].name).toBe('Stay');
    });
});
