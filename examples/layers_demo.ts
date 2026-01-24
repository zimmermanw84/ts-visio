import { VisioDocument } from '../src/index';
// @ts-ignore
import fs from 'fs';

async function run() {
    const doc = await VisioDocument.create();
    const page = doc.pages[0];

    console.log('Creating Layers Demo...');

    // 1. Define Layers
    const wireframeLayer = await page.addLayer('Wireframe');
    const annotationsLayer = await page.addLayer('Annotations', { visible: true });
    const backgroundLayer = await page.addLayer('Background', { lock: true });

    console.log(`Created layers: Wireframe(${wireframeLayer.index}), Annotations(${annotationsLayer.index}), Background(${backgroundLayer.index})`);

    // 2. Create shapes and assign to layers
    const bgBox = await page.addShape({
        text: 'Background Grid',
        x: 4, y: 5,
        width: 8, height: 6,
        fillColor: '#F0F0F0'
    });
    await bgBox.addToLayer(backgroundLayer);

    const headerBox = await page.addShape({
        text: 'Header Component',
        x: 4, y: 9,
        width: 6, height: 1
    });
    await headerBox.addToLayer(wireframeLayer);

    const sidebarBox = await page.addShape({
        text: 'Sidebar',
        x: 1.5, y: 5,
        width: 2, height: 4
    });
    await sidebarBox.addToLayer(wireframeLayer);

    const noteBox = await page.addShape({
        text: 'TODO: Add login form',
        x: 6, y: 6,
        width: 3, height: 1,
        fillColor: '#FFFFE0'
    });
    await noteBox.addToLayer(annotationsLayer);

    // 3. Toggle Visibility (e.g., hide annotations for presentation)
    console.log('Hiding annotations layer...');
    await annotationsLayer.hide();

    // 4. Save with annotations hidden
    const buffer = await doc.save();
    fs.writeFileSync('examples/layers_demo.vsdx', buffer);
    console.log('Generated examples/layers_demo.vsdx (annotations hidden)');

    // 5. Show annotations again and save variant
    await annotationsLayer.show();
    const bufferAnnotated = await doc.save();
    fs.writeFileSync('examples/layers_demo_annotated.vsdx', bufferAnnotated);
    console.log('Generated examples/layers_demo_annotated.vsdx (annotations visible)');
}

run().catch(console.error);
