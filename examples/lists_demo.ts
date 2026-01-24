import { VisioDocument } from '../src/index';
// @ts-ignore
import fs from 'fs';

async function run() {
    const doc = await VisioDocument.create();
    const page = doc.pages[0];

    // 1. Vertical List (Database Table Style)
    console.log('Creating Users Table (Vertical List)...');
    const usersTable = await page.addList({
        text: 'Users',
        x: 2, y: 10,
        width: 3, height: 1
    }, 'vertical'); // Strict vertical stacking

    // Add Columns (Items)
    const col1 = await page.addShape({ text: 'PK: ID (int)', width: 3, height: 0.5, x: 0, y: 0 });
    const col2 = await page.addShape({ text: 'Username (varchar)', width: 3, height: 0.5, x: 0, y: 0 });
    const col3 = await page.addShape({ text: 'Email (varchar)', width: 3, height: 0.5, x: 0, y: 0 });
    const col4 = await page.addShape({ text: 'Created_At (datetime)', width: 3, height: 0.5, x: 0, y: 0 });

    await usersTable.addListItem(col1);
    await usersTable.addListItem(col2);
    await usersTable.addListItem(col3);
    await usersTable.addListItem(col4);


    // 2. Horizontal List (Timeline Style)
    console.log('Creating Timeline (Horizontal List)...');
    const timeline = await page.addList({
        text: 'Project Phases',
        x: 2, y: 5,
        width: 2, height: 1.5
    }, 'horizontal'); // Strict horizontal stacking

    const phase1 = await page.addShape({ text: 'Phase 1: Planning', width: 2, height: 1, fillColor: '#e1f5fe', x: 0, y: 0 });
    const phase2 = await page.addShape({ text: 'Phase 2: Execution', width: 2, height: 1, fillColor: '#fff9c4', x: 0, y: 0 });
    const phase3 = await page.addShape({ text: 'Phase 3: Launch', width: 2, height: 1, fillColor: '#e0f2f1', x: 0, y: 0 });

    await timeline.addListItem(phase1);
    await timeline.addListItem(phase2);
    await timeline.addListItem(phase3);

    // Save
    const buffer = await doc.save();
    fs.writeFileSync('examples/lists_demo.vsdx', buffer);
    console.log('Generated examples/lists_demo.vsdx');
}

run().catch(console.error);
