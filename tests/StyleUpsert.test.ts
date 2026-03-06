/**
 * Regression tests for Bug 20:
 * updateShapeStyle must upsert individual cells within existing sections,
 * not replace the entire section (which would destroy extra cells from
 * real Visio files such as FillForegndTrans, ShdwPattern, etc.).
 */
import { describe, it, expect } from 'vitest';
import JSZip from 'jszip';
import { VisioDocument } from '../src/VisioDocument';
import { VisioPackage } from '../src/VisioPackage';
import { ShapeModifier } from '../src/ShapeModifier';
import { XMLParser } from 'fast-xml-parser';

const PARSER = new XMLParser({ ignoreAttributes: false, attributeNamePrefix: '@_' });

function normArray(val: any): any[] {
    if (!val) return [];
    return Array.isArray(val) ? val : [val];
}

async function getRawShape(doc: VisioDocument, shapeId: string): Promise<any> {
    const buf = await doc.save();
    const pkg = new VisioPackage();
    await pkg.load(buf);
    const xml = pkg.getFileText('visio/pages/page1.xml');
    const parsed = PARSER.parse(xml);
    const shapes = normArray(parsed.PageContents?.Shapes?.Shape);
    return shapes.find((s: any) => s['@_ID'] === shapeId);
}

/**
 * Build a minimal VisioPackage whose page1.xml already contains a Fill section
 * with extra cells (FillForegndTrans) that a real Visio file might include.
 */
async function buildPkgWithExtraFillCell(): Promise<{ pkg: VisioPackage; modifier: ShapeModifier }> {
    const zip = new JSZip();
    zip.file('visio/pages/page1.xml', `
        <PageContents>
            <Shapes>
                <Shape ID="1" Type="Shape">
                    <Section N="Fill">
                        <Cell N="FillForegnd" V="#aabbcc" F="RGB(170,187,204)"/>
                        <Cell N="FillBkgnd" V="#ffffff"/>
                        <Cell N="FillPattern" V="1"/>
                        <Cell N="FillForegndTrans" V="0.5"/>
                    </Section>
                </Shape>
            </Shapes>
        </PageContents>
    `);
    const buf = await zip.generateAsync({ type: 'nodebuffer' });
    const pkg = new VisioPackage();
    await pkg.load(buf);
    const modifier = new ShapeModifier(pkg);
    modifier.registerPage('1', 'visio/pages/page1.xml');
    return { pkg, modifier };
}

describe('Bug 20: updateShapeStyle preserves extra cells in existing sections', () => {
    it('second fillColor call preserves extra Fill cells (FillBkgnd, FillPattern)', async () => {
        const doc   = await VisioDocument.create();
        const page  = doc.pages[0];
        const shape = await page.addShape({ text: 'Test', x: 1, y: 1, width: 2, height: 1 });

        // First call creates the Fill section with FillBkgnd and FillPattern defaults.
        await shape.setStyle({ fillColor: '#ff0000' });
        // Second call must upsert FillForegnd only — not wipe FillBkgnd / FillPattern.
        await shape.setStyle({ fillColor: '#00ff00' });

        const raw = await getRawShape(doc, shape.id);
        const sections = normArray(raw?.Section);
        const fillSections = sections.filter((s: any) => s['@_N'] === 'Fill');

        // Exactly one Fill section.
        expect(fillSections).toHaveLength(1);

        const cells = normArray(fillSections[0].Cell);
        // FillForegnd updated to new color.
        expect(cells.find((c: any) => c['@_N'] === 'FillForegnd')?.['@_V']).toBe('#00ff00');
        // FillBkgnd and FillPattern retained from initial section.
        expect(cells.find((c: any) => c['@_N'] === 'FillBkgnd')).toBeDefined();
        expect(cells.find((c: any) => c['@_N'] === 'FillPattern')).toBeDefined();
    });

    it('extra FillForegndTrans cell from real Visio XML is preserved after fillColor update', async () => {
        const { pkg, modifier } = await buildPkgWithExtraFillCell();

        // Update only fillColor — FillForegndTrans must survive.
        await modifier.updateShapeStyle('1', '1', { fillColor: '#112233' });

        const xml    = pkg.getFileText('visio/pages/page1.xml');
        const parsed = PARSER.parse(xml);
        const raw    = normArray(parsed.PageContents?.Shapes?.Shape).find((s: any) => s['@_ID'] === '1' || s['@_ID'] === 1);
        const cells  = normArray(normArray(raw?.Section).find((s: any) => s['@_N'] === 'Fill')?.Cell);

        expect(cells.find((c: any) => c['@_N'] === 'FillForegnd')?.['@_V']).toBe('#112233');
        // Extra cell must NOT be destroyed.
        expect(cells.find((c: any) => c['@_N'] === 'FillForegndTrans')?.['@_V']).toBe('0.5');
    });

    it('second lineColor call preserves LinePattern and LineWeight', async () => {
        const doc   = await VisioDocument.create();
        const page  = doc.pages[0];
        const shape = await page.addShape({ text: 'L', x: 1, y: 1, width: 2, height: 1 });

        await shape.setStyle({ lineColor: '#111111', lineWeight: 2, linePattern: 2 });
        // Update only color — weight and pattern must survive.
        await shape.setStyle({ lineColor: '#999999' });

        const raw = await getRawShape(doc, shape.id);
        const sections = normArray(raw?.Section);
        const lineSections = sections.filter((s: any) => s['@_N'] === 'Line');
        expect(lineSections).toHaveLength(1);

        const cells = normArray(lineSections[0].Cell);
        expect(cells.find((c: any) => c['@_N'] === 'LineColor')?.['@_V']).toBe('#999999');
        expect(parseFloat(cells.find((c: any) => c['@_N'] === 'LineWeight')?.['@_V'])).toBeCloseTo(2 / 72, 6);
        expect(cells.find((c: any) => c['@_N'] === 'LinePattern')?.['@_V']).toBe('2');
    });

    it('second bold call preserves italic bit in Style', async () => {
        const doc   = await VisioDocument.create();
        const page  = doc.pages[0];
        const shape = await page.addShape({ text: 'T', x: 1, y: 1, width: 2, height: 1 });

        // Set italic first.
        await shape.setStyle({ italic: true });
        // Then set bold — italic must be preserved.
        await shape.setStyle({ bold: true });

        const raw = await getRawShape(doc, shape.id);
        const charSec = normArray(raw?.Section).find((s: any) => s['@_N'] === 'Character');
        const rows = normArray(charSec?.Row);
        const cells = normArray(rows[0]?.Cell);
        const styleVal = parseInt(cells.find((c: any) => c['@_N'] === 'Style')?.['@_V'] || '0');
        expect(styleVal & 1).toBe(1); // bold bit set
        expect(styleVal & 2).toBe(2); // italic bit preserved
    });

    it('updating only lineWeight does not add a second Line section', async () => {
        const doc   = await VisioDocument.create();
        const page  = doc.pages[0];
        const shape = await page.addShape({ text: 'W', x: 1, y: 1, width: 2, height: 1 });

        await shape.setStyle({ lineColor: '#000000' });
        await shape.setStyle({ lineWeight: 3 });

        const raw = await getRawShape(doc, shape.id);
        const lineSections = normArray(raw?.Section).filter((s: any) => s['@_N'] === 'Line');
        expect(lineSections).toHaveLength(1);
    });
});
