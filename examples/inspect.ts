/**
 * inspect.ts
 *
 * Demonstrates the ts-visio reading APIs:
 *   - VisioDocument.load() — load an existing .vsdx file
 *   - doc.getMetadata()    — read document properties
 *   - doc.getPage(name)    — look up a page by name
 *   - page.getShapes()     — enumerate shapes on a page
 *   - page.findShapes()    — filter shapes by predicate
 *   - page.getConnectors() — enumerate connector shapes
 *   - shape.getProperties()  — read custom shape data
 *   - shape.getHyperlinks()  — read hyperlinks
 *
 * Run network-diagram.ts first to produce network-topology.vsdx,
 * or the script will create a fresh document to inspect.
 */
import { VisioDocument } from '../src/VisioDocument';
import * as fs   from 'fs';
import * as path from 'path';

function hr(label: string) {
    console.log(`\n${'─'.repeat(60)}`);
    console.log(`  ${label}`);
    console.log('─'.repeat(60));
}

async function run() {
    // ── Load document ──────────────────────────────────────────────────────────
    const vsdxPath = path.resolve(__dirname, 'network-topology.vsdx');

    let doc: VisioDocument;

    if (fs.existsSync(vsdxPath)) {
        console.log(`Loading: ${vsdxPath}`);
        const buf = fs.readFileSync(vsdxPath);
        doc = await VisioDocument.load(buf);
    } else {
        // Create a minimal example document inline if the file doesn't exist yet
        console.log('network-topology.vsdx not found — creating a sample document instead.');
        doc = await VisioDocument.create();
        doc.setMetadata({ title: 'Sample', author: 'ts-visio', description: 'Inline sample' });
        const page = doc.pages[0];
        const a = await page.addShape({ text: 'Node A', x: 2, y: 5, width: 1.5, height: 1 });
        const b = await page.addShape({ text: 'Node B', x: 6, y: 5, width: 1.5, height: 1 });
        a.addData('ip', { value: '10.0.0.1', label: 'IP Address' });
        await page.connectShapes(a, b, undefined, undefined, { lineColor: '#333333' });
    }

    // ── 1. Document metadata ───────────────────────────────────────────────────
    hr('Document Metadata');
    const meta = doc.getMetadata();
    console.log(`  Title:       ${meta.title       ?? '(none)'}`);
    console.log(`  Author:      ${meta.author      ?? '(none)'}`);
    console.log(`  Description: ${meta.description ?? '(none)'}`);
    console.log(`  Company:     ${meta.company     ?? '(none)'}`);
    if (meta.modified) console.log(`  Modified:    ${meta.modified.toISOString()}`);

    // ── 2. Pages ───────────────────────────────────────────────────────────────
    hr('Pages');
    for (const p of doc.pages) {
        console.log(`  Page ${p.id}: "${p.name}"  ${p.pageWidth.toFixed(2)}" × ${p.pageHeight.toFixed(2)}"`);
    }

    // ── 3. Shapes ──────────────────────────────────────────────────────────────
    hr('Shapes on first page');
    const page   = doc.pages[0];
    const shapes = page.getShapes();

    // Filter out connector shapes (connectors also appear as shapes)
    const connectorIds = new Set(page.getConnectors().map(c => c.id));
    const regularShapes = shapes.filter(s => !connectorIds.has(s.id));

    console.log(`  ${regularShapes.length} non-connector shape(s):`);
    for (const shape of regularShapes) {
        const props = shape.getProperties();
        const propSummary = Object.keys(props).length > 0
            ? `  [data: ${Object.keys(props).join(', ')}]`
            : '';
        console.log(`    ID ${shape.id.padStart(3)}: "${shape.text}"  (${shape.x.toFixed(2)}, ${shape.y.toFixed(2)})  ${shape.width.toFixed(2)}"×${shape.height.toFixed(2)}"${propSummary}`);
    }

    // ── 4. findShapes — shapes that have custom data ───────────────────────────
    hr('Shapes with custom properties');
    const withData = page.findShapes(s => Object.keys(s.getProperties()).length > 0);
    if (withData.length === 0) {
        console.log('  (none)');
    }
    for (const shape of withData) {
        console.log(`  "${shape.text}" (ID ${shape.id}):`);
        for (const [key, data] of Object.entries(shape.getProperties())) {
            console.log(`    ${key}: ${JSON.stringify(data.value)}${data.label ? ` (${data.label})` : ''}`);
        }
    }

    // ── 5. Hyperlinks ──────────────────────────────────────────────────────────
    hr('Shapes with hyperlinks');
    const withLinks = page.findShapes(s => s.getHyperlinks().length > 0);
    if (withLinks.length === 0) {
        console.log('  (none)');
    }
    for (const shape of withLinks) {
        for (const link of shape.getHyperlinks()) {
            const target = link.address || link.subAddress || '(internal)';
            console.log(`  "${shape.text}" → ${target}${link.description ? ` [${link.description}]` : ''}`);
        }
    }

    // ── 6. Connectors ─────────────────────────────────────────────────────────
    hr('Connectors');
    const connectors = page.getConnectors();
    console.log(`  ${connectors.length} connector(s):`);
    for (const conn of connectors) {
        const fromShape = regularShapes.find(s => s.id === conn.fromShapeId);
        const toShape   = regularShapes.find(s => s.id === conn.toShapeId);
        const fromLabel = fromShape ? `"${fromShape.text}"` : `ID ${conn.fromShapeId}`;
        const toLabel   = toShape   ? `"${toShape.text}"`   : `ID ${conn.toShapeId}`;
        const styleParts: string[] = [];
        if (conn.style.lineColor)   styleParts.push(`color=${conn.style.lineColor}`);
        if (conn.style.routing)     styleParts.push(`routing=${conn.style.routing}`);
        if (conn.style.linePattern) styleParts.push(`pattern=${conn.style.linePattern}`);
        console.log(`    ${fromLabel} → ${toLabel}${styleParts.length ? `  [${styleParts.join(', ')}]` : ''}`);
    }

    // ── 7. Styles ──────────────────────────────────────────────────────────────
    hr('Document-level StyleSheets');
    const styles = doc.getStyles();
    for (const style of styles) {
        console.log(`  ID ${style.id}: "${style.name}"`);
    }

    console.log('\nDone.');
}

run().catch(console.error);
