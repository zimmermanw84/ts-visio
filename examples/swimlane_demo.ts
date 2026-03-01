/**
 * swimlane_demo.ts
 *
 * Demonstrates: swimlane pool + lane layout, page size (Letter landscape),
 * document metadata, non-rectangular geometry (ellipse for start/end,
 * diamond for decision nodes), connector styling with arrow heads.
 */
import { VisioDocument } from '../src/index';
import { ArrowHeads } from '../src/index';
import * as fs   from 'fs';
import * as path from 'path';

async function run() {
    console.log('Generating Swimlane Diagram...');

    const doc  = await VisioDocument.create();
    const page = doc.pages[0];

    // ── Page setup ─────────────────────────────────────────────────────────────
    page.setNamedSize('Letter', 'landscape');

    doc.setMetadata({
        title:       'User Registration Process',
        author:      'ts-visio',
        description: 'Swimlane BPMN-style flow for user registration across Client, Server, and DB.',
    });

    // ── Pool ───────────────────────────────────────────────────────────────────
    const pool = await page.addSwimlanePool({
        text:     'User Registration Process',
        x: 5.5, y: 4.5,
        width: 10, height: 7,
        fillColor: '#FFFFFF',
    });

    // ── Lanes ──────────────────────────────────────────────────────────────────
    const clientLane = await page.addSwimlaneLane({
        text:     'Client (Browser)',
        x: 0, y: 0,
        width: 10, height: 2.2,
        fillColor: '#DDEBF7',   // light blue
    });

    const serverLane = await page.addSwimlaneLane({
        text:     'Server (API)',
        x: 0, y: 0,
        width: 10, height: 2.2,
        fillColor: '#E2F0D9',   // light green
    });

    const dbLane = await page.addSwimlaneLane({
        text:     'Database',
        x: 0, y: 0,
        width: 10, height: 2.2,
        fillColor: '#FFF2CC',   // light amber
    });

    await pool.addListItem(clientLane);
    await pool.addListItem(serverLane);
    await pool.addListItem(dbLane);

    // ── Process nodes ──────────────────────────────────────────────────────────
    // Client lane: Start (ellipse), Submit Form (rectangle)
    const startNode = await page.addShape({
        text:     'Start',
        x: 1.2, y: 8,
        width:   0.8, height: 0.8,
        geometry: 'ellipse',    // start/end are circles in BPMN
        fillColor: '#70AD47',
        fontColor: '#FFFFFF',
        bold:     true,
        fontSize: 9,
    });
    await startNode.setStyle({ lineColor: '#507C34', lineWeight: 1.5 });

    const formNode = await page.addShape({
        text:     'Submit\nRegistration Form',
        x: 3.5, y: 8,
        width:   1.8, height: 1,
        fillColor: '#DDEBF7',
        fontSize: 9,
        horzAlign: 'center',
    });
    await formNode.setStyle({ lineColor: '#2E75B6', lineWeight: 1 });

    await clientLane.addMember(startNode);
    await clientLane.addMember(formNode);

    // Server lane: Validate Data (diamond = decision)
    const validateNode = await page.addShape({
        text:     'Validate\nData?',
        x: 5.5, y: 5.8,
        width:   1.6, height: 1.2,
        geometry: 'diamond',    // decision node
        fillColor: '#E2F0D9',
        fontSize: 9,
        horzAlign: 'center',
    });
    await validateNode.setStyle({ lineColor: '#375623', lineWeight: 1.5 });

    const rejectNode = await page.addShape({
        text:     'Return\nError',
        x: 7.8, y: 8,
        width:   1.5, height: 0.9,
        fillColor: '#FFCCCC',
        fontSize: 9,
    });
    await rejectNode.setStyle({ lineColor: '#CC0000', lineWeight: 1 });

    await serverLane.addMember(validateNode);
    await clientLane.addMember(rejectNode);

    // Database lane: Save User, Send Welcome Email
    const saveNode = await page.addShape({
        text:     'Save User\nRecord',
        x: 5.5, y: 3.5,
        width:   1.8, height: 1,
        fillColor: '#FFF2CC',
        fontSize: 9,
    });
    await saveNode.setStyle({ lineColor: '#7F6000', lineWeight: 1 });

    const emailNode = await page.addShape({
        text:     'Send Welcome\nEmail',
        x: 8, y: 3.5,
        width:   1.8, height: 1,
        fillColor: '#FFF2CC',
        fontSize: 9,
    });
    await emailNode.setStyle({ lineColor: '#7F6000', lineWeight: 1 });

    // End node (ellipse)
    const endNode = await page.addShape({
        text:     'End',
        x: 9.5, y: 8,
        width:   0.8, height: 0.8,
        geometry: 'ellipse',
        fillColor: '#FF0000',
        fontColor: '#FFFFFF',
        bold:     true,
        fontSize: 9,
    });
    await endNode.setStyle({ lineColor: '#CC0000', lineWeight: 2 });

    await dbLane.addMember(saveNode);
    await dbLane.addMember(emailNode);
    await clientLane.addMember(endNode);

    // ── Connect the flow ───────────────────────────────────────────────────────
    const flowStyle = { lineColor: '#333333', lineWeight: 1, routing: 'orthogonal' as const };

    await page.connectShapes(startNode,    formNode,     ArrowHeads.None, ArrowHeads.Standard, flowStyle);
    await page.connectShapes(formNode,     validateNode, ArrowHeads.None, ArrowHeads.Standard, flowStyle);

    // Valid path: validate → save
    await page.connectShapes(validateNode, saveNode,     ArrowHeads.None, ArrowHeads.Standard,
        { ...flowStyle, lineColor: '#375623' });

    // Invalid path: validate → reject (dashed red)
    await page.connectShapes(validateNode, rejectNode,   ArrowHeads.None, ArrowHeads.Standard,
        { lineColor: '#CC0000', lineWeight: 1, linePattern: 2, routing: 'orthogonal' });

    await page.connectShapes(saveNode,     emailNode,    ArrowHeads.None, ArrowHeads.Standard, flowStyle);
    await page.connectShapes(emailNode,    endNode,      ArrowHeads.None, ArrowHeads.Standard, flowStyle);
    await page.connectShapes(rejectNode,   endNode,      ArrowHeads.None, ArrowHeads.Standard,
        { lineColor: '#CC0000', lineWeight: 1, linePattern: 2, routing: 'straight' });

    // ── Save ───────────────────────────────────────────────────────────────────
    const outFile = path.resolve(__dirname, 'swimlane_demo.vsdx');
    const buffer  = await doc.save();
    fs.writeFileSync(outFile, buffer);
    console.log(`Saved ${outFile}`);
}

run().catch(console.error);
