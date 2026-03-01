import { describe, it, expect } from 'vitest';
import { VisioDocument } from '../src/VisioDocument';
import { VisioPropType } from '../src/types/VisioTypes';

// ---------------------------------------------------------------------------
// getProperties()
// ---------------------------------------------------------------------------

describe('Shape.getProperties()', () => {
    it('returns an empty object when no properties have been set', async () => {
        const doc = await VisioDocument.create();
        const page = doc.pages[0];
        const shape = await page.addShape({ text: 'Plain', x: 1, y: 1, width: 2, height: 1 });

        expect(shape.getProperties()).toEqual({});
    });

    it('reads back a string property written via addData', async () => {
        const doc = await VisioDocument.create();
        const page = doc.pages[0];
        const shape = await page.addShape({ text: 'Node', x: 1, y: 1, width: 2, height: 1 });
        shape.addData('env', { value: 'production', label: 'Environment' });

        const props = shape.getProperties();
        expect(props['env']).toBeDefined();
        expect(props['env'].value).toBe('production');
        expect(props['env'].label).toBe('Environment');
        expect(props['env'].type).toBe(VisioPropType.String);
    });

    it('reads back a number property with correct coercion', async () => {
        const doc = await VisioDocument.create();
        const page = doc.pages[0];
        const shape = await page.addShape({ text: 'N', x: 1, y: 1, width: 1, height: 1 });
        shape.addData('priority', { value: 42, label: 'Priority' });

        const props = shape.getProperties();
        expect(typeof props['priority'].value).toBe('number');
        expect(props['priority'].value).toBe(42);
    });

    it('reads back a boolean property with correct coercion', async () => {
        const doc = await VisioDocument.create();
        const page = doc.pages[0];
        const shape = await page.addShape({ text: 'B', x: 1, y: 1, width: 1, height: 1 });
        shape.addData('active', { value: true });

        const props = shape.getProperties();
        expect(typeof props['active'].value).toBe('boolean');
        expect(props['active'].value).toBe(true);
    });

    it('reads back a Date property as a Date instance with correct year', async () => {
        const doc = await VisioDocument.create();
        const page = doc.pages[0];
        const shape = await page.addShape({ text: 'D', x: 1, y: 1, width: 1, height: 1 });
        // Use an explicit UTC date so the stored ISO string is unambiguous
        const original = new Date(Date.UTC(2025, 0, 15, 12, 0, 0));
        shape.addData('created', { value: original });

        const props = shape.getProperties();
        const retrieved = props['created'].value as Date;
        expect(retrieved).toBeInstanceOf(Date);
        expect(retrieved.getUTCFullYear()).toBe(2025);
        expect(retrieved.getUTCMonth()).toBe(0); // January
    });

    it('reads back multiple properties', async () => {
        const doc = await VisioDocument.create();
        const page = doc.pages[0];
        const shape = await page.addShape({ text: 'Multi', x: 1, y: 1, width: 2, height: 1 });
        shape.addData('name', { value: 'Alice' });
        shape.addData('age', { value: 30 });
        shape.addData('admin', { value: false });

        const props = shape.getProperties();
        expect(Object.keys(props)).toHaveLength(3);
        expect(props['name'].value).toBe('Alice');
        expect(props['age'].value).toBe(30);
        expect(props['admin'].value).toBe(false);
    });

    it('respects the hidden flag', async () => {
        const doc = await VisioDocument.create();
        const page = doc.pages[0];
        const shape = await page.addShape({ text: 'H', x: 1, y: 1, width: 1, height: 1 });
        shape.addData('secret', { value: 'x', hidden: true });

        const props = shape.getProperties();
        expect(props['secret'].hidden).toBe(true);
    });

    it('reflects updates after setPropertyValue is called again', async () => {
        const doc = await VisioDocument.create();
        const page = doc.pages[0];
        const shape = await page.addShape({ text: 'U', x: 1, y: 1, width: 1, height: 1 });
        shape.addData('status', { value: 'draft' });

        // Update through the low-level API
        shape.setPropertyValue('status', 'published');

        const props = shape.getProperties();
        expect(props['status'].value).toBe('published');
    });
});

// ---------------------------------------------------------------------------
// getHyperlinks()
// ---------------------------------------------------------------------------

describe('Shape.getHyperlinks()', () => {
    it('returns an empty array when no hyperlinks have been added', async () => {
        const doc = await VisioDocument.create();
        const page = doc.pages[0];
        const shape = await page.addShape({ text: 'Plain', x: 1, y: 1, width: 2, height: 1 });

        expect(shape.getHyperlinks()).toEqual([]);
    });

    it('reads back an external URL', async () => {
        const doc = await VisioDocument.create();
        const page = doc.pages[0];
        const shape = await page.addShape({ text: 'Link', x: 1, y: 1, width: 2, height: 1 });
        await shape.toUrl('https://example.com', 'Example Site');

        const links = shape.getHyperlinks();
        expect(links).toHaveLength(1);
        expect(links[0].address).toBe('https://example.com');
        expect(links[0].description).toBe('Example Site');
        expect(links[0].newWindow).toBe(false);
    });

    it('reads back an internal page link', async () => {
        const doc = await VisioDocument.create();
        const p2 = await doc.addPage('Dashboard');
        const page = doc.pages[0];
        const shape = await page.addShape({ text: 'Go', x: 1, y: 1, width: 2, height: 1 });
        await shape.toPage(p2, 'Go to Dashboard');

        const links = shape.getHyperlinks();
        expect(links).toHaveLength(1);
        expect(links[0].subAddress).toBe('Dashboard');
        expect(links[0].description).toBe('Go to Dashboard');
    });

    it('reads back multiple hyperlinks', async () => {
        const doc = await VisioDocument.create();
        const page = doc.pages[0];
        const shape = await page.addShape({ text: 'Multi', x: 1, y: 1, width: 2, height: 1 });
        await shape.toUrl('https://alpha.com', 'Alpha');
        await shape.toUrl('https://beta.com', 'Beta');

        const links = shape.getHyperlinks();
        expect(links).toHaveLength(2);
        expect(links.map(l => l.address)).toContain('https://alpha.com');
        expect(links.map(l => l.address)).toContain('https://beta.com');
    });

    it('address is undefined for a page-only link', async () => {
        const doc = await VisioDocument.create();
        const p2 = await doc.addPage('P2');
        const page = doc.pages[0];
        const shape = await page.addShape({ text: 'X', x: 1, y: 1, width: 1, height: 1 });
        await shape.toPage(p2);

        const [link] = shape.getHyperlinks();
        // address is set to '' by addHyperlink for page links
        expect(link.address === '' || link.address === undefined).toBe(true);
    });
});

// ---------------------------------------------------------------------------
// getLayerIndices()
// ---------------------------------------------------------------------------

describe('Shape.getLayerIndices()', () => {
    it('returns an empty array when the shape has no layer assignment', async () => {
        const doc = await VisioDocument.create();
        const page = doc.pages[0];
        const shape = await page.addShape({ text: 'Bare', x: 1, y: 1, width: 1, height: 1 });

        expect(shape.getLayerIndices()).toEqual([]);
    });

    it('returns the correct layer index after assignLayer', async () => {
        const doc = await VisioDocument.create();
        const page = doc.pages[0];
        const layer = await page.addLayer('Network');
        const shape = await page.addShape({ text: 'Node', x: 1, y: 1, width: 1, height: 1 });
        await shape.assignLayer(layer);

        expect(shape.getLayerIndices()).toEqual([layer.index]);
    });

    it('returns multiple indices when a shape belongs to several layers', async () => {
        const doc = await VisioDocument.create();
        const page = doc.pages[0];
        const l0 = await page.addLayer('Layer A');
        const l1 = await page.addLayer('Layer B');
        const shape = await page.addShape({ text: 'Multi', x: 1, y: 1, width: 1, height: 1 });
        await shape.assignLayer(l0);
        await shape.assignLayer(l1);

        const indices = shape.getLayerIndices();
        expect(indices).toContain(l0.index);
        expect(indices).toContain(l1.index);
        expect(indices).toHaveLength(2);
    });

    it('does not duplicate an index if assignLayer is called twice with the same layer', async () => {
        const doc = await VisioDocument.create();
        const page = doc.pages[0];
        const layer = await page.addLayer('Solo');
        const shape = await page.addShape({ text: 'X', x: 1, y: 1, width: 1, height: 1 });
        await shape.assignLayer(layer);
        await shape.assignLayer(layer); // idempotent

        expect(shape.getLayerIndices()).toHaveLength(1);
    });
});
