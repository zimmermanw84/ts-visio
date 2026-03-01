/**
 * containers_demo.ts
 *
 * Demonstrates: containers, page size, document metadata,
 * setStyle({ lineColor, lineWeight }) for container borders,
 * ellipse geometry for server icons, and connector styling.
 */
import { VisioDocument } from '../src/VisioDocument';
import { ArrowHeads } from '../src/index';
import * as fs from 'fs';
import * as path from 'path';

async function run() {
    const doc  = await VisioDocument.create();
    const page = doc.pages[0];

    // ── Page setup ─────────────────────────────────────────────────────────────
    page.setNamedSize('Letter');

    doc.setMetadata({
        title:       'Container Layout Demo',
        author:      'ts-visio',
        description: 'Demonstrates containers grouping related shapes.',
    });

    // ── Network zone container ─────────────────────────────────────────────────
    const netContainer = await page.addContainer({
        text: 'Network Zone',
        x: 4, y: 6,
        width: 6, height: 4,
        fillColor: '#EBF3FB',
    });
    // Blue border via setStyle line props
    await netContainer.setStyle({ lineColor: '#2E75B6', lineWeight: 1.5 });

    // Servers inside the network zone (ellipse geometry)
    const server1 = await page.addShape({
        text:     'Server A',
        x: 2,    y: 6,
        width:   1.2, height: 1.2,
        geometry: 'ellipse',
        fillColor: '#D6E4F0',
        lineColor: '#2E75B6',
    });

    const server2 = await page.addShape({
        text:     'Server B',
        x: 4,    y: 6,
        width:   1.2, height: 1.2,
        geometry: 'ellipse',
        fillColor: '#D6E4F0',
        lineColor: '#2E75B6',
    });

    await netContainer.addMember(server1);
    await netContainer.addMember(server2);

    // ── Database cluster container ─────────────────────────────────────────────
    const dbContainer = await page.addContainer({
        text: 'Database Cluster',
        x: 4, y: 2,
        width: 6, height: 3,
        fillColor: '#FFF2CC',
    });
    // Orange border
    await dbContainer.setStyle({ lineColor: '#C55A11', lineWeight: 1.5 });

    const primaryDb = await page.addShape({
        text:     'Primary DB',
        x: 2.5, y: 2,
        width:   1.5, height: 1,
        fillColor: '#FCE4D6',
        lineColor: '#C55A11',
    });

    const replicaDb = await page.addShape({
        text:     'Replica DB',
        x: 5.5, y: 2,
        width:   1.5, height: 1,
        fillColor: '#FCE4D6',
        lineColor: '#C55A11',
        linePattern: 2,   // dashed border = replica
    });

    await dbContainer.addMember(primaryDb);
    await dbContainer.addMember(replicaDb);

    // ── Connections ────────────────────────────────────────────────────────────
    // Server A → Primary DB
    await page.connectShapes(
        server1, primaryDb,
        ArrowHeads.None, ArrowHeads.Standard,
        { lineColor: '#2E75B6', lineWeight: 1.5, routing: 'orthogonal' },
    );

    // Primary → Replica (replication, dashed)
    await page.connectShapes(
        primaryDb, replicaDb,
        ArrowHeads.None, ArrowHeads.Standard,
        { lineColor: '#C55A11', lineWeight: 1, linePattern: 2, routing: 'straight' },
    );

    // ── Save ───────────────────────────────────────────────────────────────────
    const buffer  = await doc.save();
    const outPath = path.join(__dirname, 'containers_demo.vsdx');
    fs.writeFileSync(outPath, buffer);
    console.log(`Generated ${outPath}`);
}

run().catch(console.error);
