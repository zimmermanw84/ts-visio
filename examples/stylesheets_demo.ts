/**
 * stylesheets_demo.ts
 *
 * Dedicated demonstration of the ts-visio StyleSheet API:
 *   - doc.createStyle(name, props)  — define a reusable named style
 *   - doc.getStyles()               — enumerate all styles in the document
 *   - addShape({ styleId })         — apply all three style categories at creation
 *   - addShape({ fillStyleId, lineStyleId, textStyleId }) — granular at creation
 *   - shape.applyStyle(id)          — apply all categories post-creation
 *   - shape.applyStyle(id, 'fill')  — apply only fill post-creation
 *   - shape.setStyle({ lineColor, lineWeight, linePattern }) — direct line edits
 *
 * Covers: fill color, line color/weight/pattern, font color/size/family,
 *         bold/italic, h/v alignment, text margins, paragraph spacing.
 */
import { VisioDocument } from '../src/VisioDocument';
import * as fs   from 'fs';
import * as path from 'path';

async function run() {
    console.log('Generating stylesheets_demo.vsdx...');

    const doc  = await VisioDocument.create();
    const page = doc.pages[0];

    page.setNamedSize('Letter');
    doc.setMetadata({
        title:       'Stylesheet Demo',
        author:      'ts-visio',
        description: 'Shows createStyle(), applyStyle(), and setStyle() line props.',
    });

    // ── 1. Define styles ───────────────────────────────────────────────────────
    const headerStyle = doc.createStyle('Header', {
        fillColor:     '#4472C4',
        lineColor:     '#2F5597',
        lineWeight:    1.5,
        fontColor:     '#FFFFFF',
        bold:          true,
        fontSize:      14,
        fontFamily:    'Calibri',
        horzAlign:     'center',
        verticalAlign: 'middle',
    });

    const bodyStyle = doc.createStyle('Body', {
        fillColor:     '#DEEAF1',
        lineColor:     '#2E75B6',
        lineWeight:    1,
        fontSize:      10,
        fontFamily:    'Calibri',
        horzAlign:     'left',
        verticalAlign: 'middle',
        textMarginLeft: 0.1,
    });

    const warningStyle = doc.createStyle('Warning', {
        fillColor:     '#FFF2CC',
        lineColor:     '#BF8F00',
        lineWeight:    1.5,
        linePattern:   2,           // dashed border
        fontColor:     '#7F6000',
        bold:          true,
        italic:        true,
        fontSize:      10,
        horzAlign:     'center',
    });

    const errorStyle = doc.createStyle('Error', {
        fillColor:     '#FFE0E0',
        lineColor:     '#CC0000',
        lineWeight:    2,
        fontColor:     '#CC0000',
        bold:          true,
        fontSize:      11,
        horzAlign:     'center',
        verticalAlign: 'middle',
    });

    const codeStyle = doc.createStyle('Code', {
        fillColor:     '#1E1E1E',
        lineColor:     '#555555',
        lineWeight:    0.75,
        fontColor:     '#D4D4D4',
        fontFamily:    'Courier New',
        fontSize:      9,
        horzAlign:     'left',
        textMarginLeft:  0.1,
        textMarginTop:   0.05,
        textMarginBottom: 0.05,
    });

    // List all registered styles
    const allStyles = doc.getStyles();
    console.log('Registered styles:');
    for (const s of allStyles) {
        console.log(`  [${s.id}] ${s.name}`);
    }

    // ── 2. Apply styleId at creation ───────────────────────────────────────────
    const titleBox = await page.addShape({
        text:    'Document Title',
        x: 4.25, y: 10,
        width:   7, height: 0.8,
        styleId: headerStyle.id,   // sets LineStyle + FillStyle + TextStyle
    });

    const descBox = await page.addShape({
        text:    'This diagram demonstrates the ts-visio stylesheet system.\nStyles can be defined once and reused across many shapes.',
        x: 4.25, y: 9,
        width:   7, height: 0.75,
        styleId: bodyStyle.id,
        lineSpacing: 1.3,
    });

    // ── 3. Granular style IDs at creation ──────────────────────────────────────
    // Fill from headerStyle, text from bodyStyle — intentional mix
    const mixedBox = await page.addShape({
        text:        'Mixed: blue fill, body text style',
        x: 4.25, y: 7.9,
        width:       7, height: 0.7,
        fillStyleId: headerStyle.id,
        textStyleId: bodyStyle.id,
    });
    // Override line directly via setStyle after creation
    await mixedBox.setStyle({ lineColor: '#2F5597', lineWeight: 1, linePattern: 1 });

    // ── 4. Apply styles post-creation ─────────────────────────────────────────
    const warningBox = await page.addShape({
        text:  '⚠ Warning: this is a post-creation applyStyle() call',
        x: 4.25, y: 7,
        width: 7, height: 0.7,
    });
    warningBox.applyStyle(warningStyle.id);   // all three categories

    const errorBox = await page.addShape({
        text:  '✕ Error: applyStyle with fill only, then setStyle line',
        x: 4.25, y: 6.1,
        width: 7, height: 0.7,
    });
    errorBox.applyStyle(errorStyle.id, 'fill');  // fill only
    errorBox.applyStyle(errorStyle.id, 'text');  // text only
    await errorBox.setStyle({ lineColor: '#CC0000', lineWeight: 2 }); // line via setStyle

    const codeBox = await page.addShape({
        text:  'const doc = await VisioDocument.create();\ndoc.createStyle("Header", { fillColor: "#4472C4" });',
        x: 4.25, y: 5.1,
        width: 7, height: 0.9,
        styleId: codeStyle.id,
    });

    // ── 5. setStyle — post-creation line props ─────────────────────────────────
    // Start with body style, then change border interactively
    const lineDemo1 = await page.addShape({
        text:  'setStyle: solid red border, 2 pt',
        x: 2, y: 4,
        width: 3, height: 0.65,
        styleId: bodyStyle.id,
    });
    await lineDemo1.setStyle({ lineColor: '#FF0000', lineWeight: 2, linePattern: 1 });

    const lineDemo2 = await page.addShape({
        text:  'setStyle: dashed orange border, 1.5 pt',
        x: 5.5, y: 4,
        width: 3, height: 0.65,
        styleId: bodyStyle.id,
    });
    await lineDemo2.setStyle({ lineColor: '#FF9900', lineWeight: 1.5, linePattern: 2 });

    const lineDemo3 = await page.addShape({
        text:  'setStyle: dotted purple border, 1 pt',
        x: 4.25, y: 3.1,
        width: 3, height: 0.65,
        styleId: bodyStyle.id,
    });
    await lineDemo3.setStyle({ lineColor: '#7030A0', lineWeight: 1, linePattern: 3 });

    // ── 6. Style palette legend ────────────────────────────────────────────────
    const legendLabels = [
        { text: 'Header style',  style: headerStyle.id,  x: 1.2 },
        { text: 'Body style',    style: bodyStyle.id,    x: 2.8 },
        { text: 'Warning style', style: warningStyle.id, x: 4.4 },
        { text: 'Error style',   style: errorStyle.id,   x: 6.0 },
        { text: 'Code style',    style: codeStyle.id,    x: 7.6 },
    ];
    for (const leg of legendLabels) {
        await page.addShape({
            text:    leg.text,
            x:       leg.x, y: 1.8,
            width:   1.4, height: 0.5,
            styleId: leg.style,
            fontSize: 8,
        });
    }

    // ── Save ───────────────────────────────────────────────────────────────────
    const outPath = path.resolve(__dirname, 'stylesheets_demo.vsdx');
    const buffer  = await doc.save();
    fs.writeFileSync(outPath, buffer);
    console.log(`Saved to ${outPath}`);
}

run().catch(console.error);
