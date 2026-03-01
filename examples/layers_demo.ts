/**
 * layers_demo.ts
 *
 * Demonstrates: layers (create, assign, show/hide), page size, document
 * metadata, rounded-rectangle geometry for UI wireframe components,
 * and setStyle({ lineColor, lineWeight }) for annotation box borders.
 */
import { VisioDocument } from '../src/index';
import fs from 'fs';

async function run() {
    const doc  = await VisioDocument.create();
    const page = doc.pages[0];

    console.log('Creating Layers Demo...');

    // ── Page setup ─────────────────────────────────────────────────────────────
    page.setNamedSize('Letter');

    doc.setMetadata({
        title:       'UI Wireframe — Layers Demo',
        author:      'ts-visio',
        description: 'Illustrates Visio layers: Wireframe shapes, annotation notes, and a background fill.',
    });

    // ── Layers ─────────────────────────────────────────────────────────────────
    const wireframeLayer   = await page.addLayer('Wireframe');
    const annotationsLayer = await page.addLayer('Annotations', { visible: true });
    const backgroundLayer  = await page.addLayer('Background',  { lock: true });

    console.log(`Created layers: Wireframe(${wireframeLayer.index}), Annotations(${annotationsLayer.index}), Background(${backgroundLayer.index})`);

    // ── Background ─────────────────────────────────────────────────────────────
    const bgBox = await page.addShape({
        text:     '',
        x: 4.25, y: 5.5,
        width:   8,    height: 9,
        fillColor: '#F7F7F7',
    });
    await bgBox.setStyle({ lineColor: '#CCCCCC', lineWeight: 0.5 });
    await bgBox.addToLayer(backgroundLayer);

    // ── Wireframe components (rounded-rectangle geometry) ──────────────────────
    const headerBox = await page.addShape({
        text:     'Header Component',
        x: 4.25, y: 9.5,
        width:   7, height: 1,
        fillColor: '#E0E0E0',
        bold:     true,
        geometry: 'rounded-rectangle',
        cornerRadius: 0.1,
    });
    await headerBox.setStyle({ lineColor: '#999999', lineWeight: 1 });
    await headerBox.addToLayer(wireframeLayer);

    const sidebarBox = await page.addShape({
        text:     'Sidebar\nNavigation',
        x: 1.25, y: 6,
        width:   1.75, height: 5,
        fillColor: '#E8E8E8',
        geometry: 'rounded-rectangle',
        cornerRadius: 0.1,
        fontSize: 9,
    });
    await sidebarBox.setStyle({ lineColor: '#999999', lineWeight: 1 });
    await sidebarBox.addToLayer(wireframeLayer);

    const contentBox = await page.addShape({
        text:     'Main Content Area',
        x: 4.75, y: 6,
        width:   5, height: 5,
        fillColor: '#FFFFFF',
        geometry: 'rounded-rectangle',
        cornerRadius: 0.1,
        fontColor: '#AAAAAA',
        italic:   true,
    });
    await contentBox.setStyle({ lineColor: '#CCCCCC', lineWeight: 0.75 });
    await contentBox.addToLayer(wireframeLayer);

    const footerBox = await page.addShape({
        text:     'Footer',
        x: 4.25, y: 1.5,
        width:   7, height: 0.6,
        fillColor: '#E0E0E0',
        geometry: 'rounded-rectangle',
        cornerRadius: 0.1,
        fontSize: 9,
    });
    await footerBox.setStyle({ lineColor: '#999999', lineWeight: 0.75 });
    await footerBox.addToLayer(wireframeLayer);

    // ── Annotations ────────────────────────────────────────────────────────────
    const noteBox = await page.addShape({
        text:     '📝 TODO: Add login form\nand OAuth integration',
        x: 7.5, y: 7,
        width:   2.5, height: 1,
        fillColor: '#FFFFE0',
        fontSize: 9,
        italic:   true,
    });
    // Amber dashed border to make annotations stand out
    await noteBox.setStyle({ lineColor: '#DAA520', lineWeight: 1, linePattern: 2 });
    await noteBox.addToLayer(annotationsLayer);

    const perfNote = await page.addShape({
        text:     '⚡ Perf note: lazy-load\nsidebar items',
        x: 7.5, y: 5.5,
        width:   2.5, height: 0.85,
        fillColor: '#FFF3CD',
        fontSize: 9,
        italic:   true,
    });
    await perfNote.setStyle({ lineColor: '#DAA520', lineWeight: 1, linePattern: 2 });
    await perfNote.addToLayer(annotationsLayer);

    // ── Save variants ──────────────────────────────────────────────────────────
    console.log('Hiding annotations layer...');
    await annotationsLayer.hide();

    const buffer = await doc.save();
    fs.writeFileSync('examples/layers_demo.vsdx', buffer);
    console.log('Generated examples/layers_demo.vsdx (annotations hidden)');

    await annotationsLayer.show();
    const bufferAnnotated = await doc.save();
    fs.writeFileSync('examples/layers_demo_annotated.vsdx', bufferAnnotated);
    console.log('Generated examples/layers_demo_annotated.vsdx (annotations visible)');
}

run().catch(console.error);
