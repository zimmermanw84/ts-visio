
import { VisioDocument } from '../src/index'; // Assuming index exports everything needed, or use relative paths if not built
// Since we are running with tsx in the repo, valid imports would be from '../src/VisioDocument' etc.
// But valid 'example' usage should try to import from the package entry point if possible, or simulate it.
// Let's use relative imports for now to work with tsx without build.
import { VisioDocument as VisioDoc } from '../src/VisioDocument';
import * as fs from 'fs';
import * as path from 'path';

async function run() {
    console.log('Generating Swimlane Diagram...');

    // 1. Create Document
    const doc = await VisioDoc.create();
    const page = doc.pages[0];

    // 2. Create the Pool (Process Overview)
    const pool = await page.addSwimlanePool({
        text: 'User Registration Process',
        x: 5.5, y: 6, // Roughly center of an 8.5x11 page
        width: 10, height: 6,
        fillColor: '#FFFFFF'
    });

    // 3. Create Lanes (Roles)
    const clientLane = await page.addSwimlaneLane({
        text: 'Client (Browser)',
        x: 0, y: 0, // Position relative to Pool handled by auto-layout
        width: 10, height: 2,
        fillColor: '#DDEBF7' // Light Blue
    });

    const serverLane = await page.addSwimlaneLane({
        text: 'Server (API)',
        x: 0, y: 0,
        width: 10, height: 2,
        fillColor: '#E2F0D9' // Light Green
    });

    const dbLane = await page.addSwimlaneLane({
        text: 'Database',
        x: 0, y: 0,
        width: 10, height: 2,
        fillColor: '#FFF2CC' // Light Orange
    });

    // 4. Add Lanes to Pool (Order matters!)
    await pool.addListItem(clientLane);
    await pool.addListItem(serverLane);
    await pool.addListItem(dbLane);

    // 5. Add Shapes (Process Steps)

    // -- Client Lane --
    const startNode = await page.addShape({
        text: 'Start',
        x: 1.5, y: 8, // Relative-ish visual placement (Visual coordinates are absolute in Page)
        width: 0.8, height: 0.8,
        type: 'Shape', masterId: 'Start/End'
    });
    // For now we must update visual position ourselves or use a layout engine.
    // Let's place them explicitly.
    // Client Lane Y range: Top third of Pool?
    // Wait, Visio lists usually stack Top-to-Bottom.
    // If Pool Center Y=6, Height=6. Top=9, Bottom=3.
    // Lane 1 (Client): Top at 9, Height 2. Center Y = 8.
    // Lane 2 (Server): Top at 7, Height 2. Center Y = 6.
    // Lane 3 (DB): Top at 5, Height 2. Center Y = 4.

    // Actually, `addListItem` triggers `resizeContainerToFit`, which might shift things.
    // But let's assume standard stacking for now.

    // Update: We should place shapes *roughly* where they belong.
    // We can use `shape.updatePosition` if needed, or just create them at correct Y.
    const formNode = await page.addShape({
        text: 'Submit Form',
        x: 3.5, y: 8,
        width: 1.5, height: 0.75,
        masterId: 'Process'
    });

    await clientLane.addMember(startNode);
    // await page.modifier.updateShapePosition(page.id, startNode.id, 1.5, 8); // Accessing private modifier - Fixed by initial placement

    await clientLane.addMember(formNode);

    // -- Server Lane --
    const validateNode = await page.addShape({
        text: 'Validate Data',
        x: 5.5, y: 6,
        width: 1.5, height: 0.75,
        masterId: 'Decision'
    });

    await serverLane.addMember(validateNode);

    // -- Database Lane --
    const saveNode = await page.addShape({
        text: 'Save User',
        x: 7.5, y: 4,
        width: 1.5, height: 0.75,
        masterId: 'Database'
    });

    await dbLane.addMember(saveNode);

    // 6. Connect the Flows
    await page.connectShapes(startNode, formNode);
    await page.connectShapes(formNode, validateNode);
    await page.connectShapes(validateNode, saveNode, 'Yes'); // Label on connector? Not yet supported in fluent API maybe

    // 7. Save
    const outFile = 'swimlane_example.vsdx';
    const buffer = await doc.save();
    fs.writeFileSync(outFile, buffer);
    console.log(`Saved ${outFile}`);
}

run().catch(console.error);
