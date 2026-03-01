/**
 * network-diagram.ts
 *
 * Demonstrates: page size, document metadata, document-level stylesheets,
 * non-rectangular geometry (ellipse, diamond), named connection points,
 * connector styling, shape.setStyle({ lineColor, lineWeight }),
 * and page.getConnectors() to read connections back after creation.
 */
import { VisioDocument } from '../src/VisioDocument';
import { StandardConnectionPoints, ArrowHeads } from '../src/index';
import * as path from 'path';
import * as fs from 'fs';

async function run() {
    console.log('Generating network-topology.vsdx...');

    // ── 1. Document setup ──────────────────────────────────────────────────────
    const doc  = await VisioDocument.create();
    const page = doc.pages[0];

    // Page size: Letter landscape
    page.setNamedSize('Letter', 'landscape');

    // Document-level metadata
    doc.setMetadata({
        title:       'Network Topology Diagram',
        author:      'ts-visio',
        description: 'Example network diagram demonstrating stylesheets, geometry, and connection points.',
        company:     'Acme Corp',
    });

    // ── 2. Document-level stylesheets ─────────────────────────────────────────
    // Standard device style: grey fill, medium border, bold Calibri
    const deviceStyle = doc.createStyle('NetworkDevice', {
        fillColor:  '#E8E8E8',
        lineColor:  '#555555',
        lineWeight: 1,
        bold:       true,
        fontFamily: 'Calibri',
        fontSize:   10,
        horzAlign:  'center',
        verticalAlign: 'middle',
    });

    // Critical / firewall style: red tint, thick red border
    const criticalStyle = doc.createStyle('CriticalDevice', {
        fillColor:  '#FFCCCC',
        lineColor:  '#CC0000',
        lineWeight: 2,
        bold:       true,
        fontFamily: 'Calibri',
        fontSize:   10,
    });

    // Uplink connector style: dark grey, 1.5 pt, orthogonal routing
    const uplinkStyle = doc.createStyle('UplinkLink', {
        lineColor:  '#333333',
        lineWeight: 1.5,
    });

    console.log(`Created styles: #${deviceStyle.id} NetworkDevice, #${criticalStyle.id} CriticalDevice, #${uplinkStyle.id} UplinkLink`);

    // ── 3. Core switch (centre) ────────────────────────────────────────────────
    const coreSwitch = await page.addShape({
        text:     'Core Switch\n(L3)',
        x: 5.5, y: 5.5,
        width: 2.5, height: 1.5,
        styleId: deviceStyle.id,    // apply all-in-one stylesheet
    });

    // Named connection points on the core switch for precise port-to-port connections
    coreSwitch.addConnectionPoint({ name: 'Top',    xFraction: 0.5, yFraction: 1.0 });
    coreSwitch.addConnectionPoint({ name: 'Right',  xFraction: 1.0, yFraction: 0.5 });
    coreSwitch.addConnectionPoint({ name: 'Bottom', xFraction: 0.5, yFraction: 0.0 });
    coreSwitch.addConnectionPoint({ name: 'Left',   xFraction: 0.0, yFraction: 0.5 });

    // Shape data
    coreSwitch
        .addData('ip',    { value: '10.0.0.1',       label: 'Management IP' })
        .addData('model', { value: 'Cisco Catalyst',  label: 'Hardware Model' })
        .addData('vlan',  { value: 100,               label: 'Mgmt VLAN' })
        .addData('asset', { value: 'SW-CORE-99',      hidden: true });

    // ── 4. Edge firewall (top) — ellipse shape, critical style ─────────────────
    const firewall = await page.addShape({
        text:     'Edge Firewall',
        x: 5.5, y: 9,
        width: 2.5, height: 1.2,
        geometry: 'ellipse',        // non-rectangular geometry
        styleId:  criticalStyle.id,
    });
    firewall.addData('policy_ver', { value: '2.5.1' });

    // Also demonstrate post-creation line style change
    await firewall.setStyle({ lineWeight: 2.5 });  // bump the border even thicker

    // ── 5. Web server (left) — ellipse ─────────────────────────────────────────
    const webServer = await page.addShape({
        text:     'Web01\n(Nginx)',
        x: 2, y: 5.5,
        width: 1.8, height: 1.2,
        geometry: 'ellipse',
        styleId:  deviceStyle.id,
    });
    webServer
        .addData('ip',        { value: '10.0.10.5' })
        .addData('role',      { value: 'Frontend' })
        .addData('last_patch',{ value: new Date() })
        .addData('auto_scale',{ value: true });

    // ── 6. DB server (right) — ellipse ─────────────────────────────────────────
    const dbServer = await page.addShape({
        text:     'DB01\n(Postgres)',
        x: 9, y: 5.5,
        width: 1.8, height: 1.2,
        geometry: 'ellipse',
        styleId:  deviceStyle.id,
    });
    dbServer
        .addData('ip',        { value: '10.0.20.5' })
        .addData('role',      { value: 'Database' })
        .addData('is_primary',{ value: true, label: 'Primary Node' });

    // ── 7. Load balancer (bottom) — diamond ────────────────────────────────────
    const loadBalancer = await page.addShape({
        text:     'Load\nBalancer',
        x: 5.5, y: 2,
        width: 2, height: 1.5,
        geometry: 'diamond',        // decision / router diamond shape
        styleId:  deviceStyle.id,
    });
    loadBalancer.addData('algo', { value: 'Round Robin', label: 'LB Algorithm' });

    // ── 8. Connect with styled connectors using named ports ────────────────────
    // Firewall ↔ Core Switch  (port-to-port: firewall Bottom → switch Top)
    await page.connectShapes(
        firewall, coreSwitch,
        ArrowHeads.None, ArrowHeads.None,
        { lineColor: '#CC0000', lineWeight: 2, routing: 'orthogonal' },
        'center',
        { name: 'Top' },    // toPort: named connection point on the switch
    );

    // Web Server → Core Switch (Left port)
    await page.connectShapes(
        webServer, coreSwitch,
        ArrowHeads.None, ArrowHeads.Standard,
        { lineColor: '#333333', lineWeight: 1.5, routing: 'orthogonal' },
        'center',
        { name: 'Left' },
    );

    // DB Server → Core Switch (Right port)
    await page.connectShapes(
        dbServer, coreSwitch,
        ArrowHeads.None, ArrowHeads.Standard,
        { lineColor: '#333333', lineWeight: 1.5, routing: 'orthogonal' },
        'center',
        { name: 'Right' },
    );

    // Core Switch → Load Balancer (Bottom port)
    await page.connectShapes(
        coreSwitch, loadBalancer,
        ArrowHeads.None, ArrowHeads.Standard,
        { lineColor: '#555555', lineWeight: 1, routing: 'orthogonal' },
        { name: 'Bottom' },
        'center',
    );

    // Web ↔ DB application link (dashed)
    await page.connectShapes(
        webServer, dbServer,
        ArrowHeads.None, ArrowHeads.Standard,
        { lineColor: '#999999', lineWeight: 1, linePattern: 2, routing: 'straight' },
    );

    // ── 9. Read connectors back ────────────────────────────────────────────────
    const connectors = page.getConnectors();
    console.log(`\nPage has ${connectors.length} connectors:`);
    for (const conn of connectors) {
        console.log(`  ${conn.fromShapeId} → ${conn.toShapeId}  color=${conn.style.lineColor}  routing=${conn.style.routing}`);
    }

    // ── 10. Save ───────────────────────────────────────────────────────────────
    const outPath = path.resolve(__dirname, 'network-topology.vsdx');
    await doc.save(outPath);
    console.log(`\nSaved to ${outPath}`);
}

run().catch(console.error);
