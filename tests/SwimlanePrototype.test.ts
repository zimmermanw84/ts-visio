
import { describe, it } from 'vitest';
import { VisioDocument } from '../src/VisioDocument';
import * as fs from 'fs';
import * as path from 'path';

describe('Swimlane Prototype', () => {
    it('should create a basic Swimlane Diagram (Pool + Lanes)', async () => {
        const doc = await VisioDocument.create();
        const page = doc.pages[0];

        // 1. Create the Pool (A Vertical List)
        const pool = await page.addSwimlanePool({
            text: 'System Process Pool',
            x: 5, y: 6, // Center of page roughly
            width: 8, height: 6,
            fillColor: '#ffffff'
        });

        // 2. Create Lanes (Containers)
        // Lane 1
        const lane1 = await page.addSwimlaneLane({
            text: 'Lane 1: User',
            x: 0, y: 0, // Pos inside list is auto-handled by addListItem
            width: 8, height: 2,
            fillColor: '#ddebf7'
        });

        // Lane 2
        const lane2 = await page.addSwimlaneLane({
            text: 'Lane 2: System',
            x: 0, y: 0,
            width: 8, height: 2,
            fillColor: '#e2f0d9'
        });

        // 3. Add Lanes to Pool (As List Items)
        // Use public Shape API
        await pool.addListItem(lane1);
        await pool.addListItem(lane2);


        // 4. Add Shapes into Lanes
        // Shape in Lane 1
        const startNode = await page.addShape({
            text: 'Start',
            x: 2, y: 1.5,
            width: 1, height: 0.5,
            fillColor: '#ffffff',
            type: 'Shape',
            masterId: 'Terminator'
        });

        // Shape in Lane 2
        const processNode = await page.addShape({
            text: 'Process Request',
            x: 4, y: 3.5,
            width: 1.5, height: 0.75,
            fillColor: '#ffffff',
            type: 'Shape',
            masterId: 'Process'
        });

        // Add to containers using public API
        await lane1.addMember(startNode);
        await lane2.addMember(processNode);

        // 5. Connect them
        await page.connectShapes(startNode, processNode);

        // 6. Save
        const outDir = path.join(__dirname, 'out');
        if (!fs.existsSync(outDir)) fs.mkdirSync(outDir);
        const buffer = await doc.save();
        fs.writeFileSync(path.join(outDir, 'swimlane_prototype.vsdx'), buffer);
        console.log('Swimlane prototype saved to tests/out/swimlane_prototype.vsdx');
    });
});
