
import { VisioDocument } from '../src/VisioDocument';
import * as fs from 'fs';
import * as path from 'path';

async function run() {
    const doc = await VisioDocument.create();
    const page = doc.pages[0];

    // 1. Create a "Network" Container
    const netContainer = await page.addContainer({
        text: 'Network Zone',
        x: 1, y: 1,
        width: 6, height: 4
    });

    // 2. Add shapes visually "inside" the container
    // Note: Visio containers rely on spatial containment + relationship logic.
    // Ideally we would add 'MemberOf' relationships, but for Phase 1 we just place them adding
    // them to the page.
    // (If we implemented true container logic they would move with it, but for now just visual).

    const server1 = await page.addShape({
        text: 'Server A',
        x: 2, y: 2,
        width: 1, height: 1
    });

    const server2 = await page.addShape({
        text: 'Server B',
        x: 4, y: 2,
        width: 1, height: 1
    });

    // 3. Create a "Database" Container
    const dbContainer = await page.addContainer({
        text: 'Database Cluster',
        x: 8, y: 1,
        width: 4, height: 4
    });

    const db1 = await page.addShape({
        text: 'Primary DB',
        x: 9, y: 2,
        width: 1.5, height: 1
    });

    // Connect Server A to DB
    await page.connectShapes(server1, db1);

    const buffer = await doc.save();
    const outPath = path.join(__dirname, 'containers_demo.vsdx');
    fs.writeFileSync(outPath, buffer);
    console.log(`Generated ${outPath}`);
}

run().catch(console.error);
