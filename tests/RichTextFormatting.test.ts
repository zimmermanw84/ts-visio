import { describe, it, expect } from 'vitest';
import { VisioDocument } from '../src/VisioDocument';
import { createCharacterSection, createParagraphSection, createTextBlockSection } from '../src/utils/StyleHelpers';

// ---------------------------------------------------------------------------
// createCharacterSection — Style bitmask
// ---------------------------------------------------------------------------

describe('createCharacterSection Style bitmask', () => {
    it('no flags → Style=0', () => {
        const section = createCharacterSection({});
        const styleCell = section.Row![0].Cell.find((c: any) => c['@_N'] === 'Style');
        expect(styleCell['@_V']).toBe('0');
    });

    it('bold → Style=1', () => {
        const section = createCharacterSection({ bold: true });
        const styleCell = section.Row![0].Cell.find((c: any) => c['@_N'] === 'Style');
        expect(styleCell['@_V']).toBe('1');
    });

    it('italic → Style=2', () => {
        const section = createCharacterSection({ italic: true });
        const styleCell = section.Row![0].Cell.find((c: any) => c['@_N'] === 'Style');
        expect(styleCell['@_V']).toBe('2');
    });

    it('underline → Style=4', () => {
        const section = createCharacterSection({ underline: true });
        const styleCell = section.Row![0].Cell.find((c: any) => c['@_N'] === 'Style');
        expect(styleCell['@_V']).toBe('4');
    });

    it('strikethrough → Style=8', () => {
        const section = createCharacterSection({ strikethrough: true });
        const styleCell = section.Row![0].Cell.find((c: any) => c['@_N'] === 'Style');
        expect(styleCell['@_V']).toBe('8');
    });

    it('bold + italic → Style=3', () => {
        const section = createCharacterSection({ bold: true, italic: true });
        const styleCell = section.Row![0].Cell.find((c: any) => c['@_N'] === 'Style');
        expect(styleCell['@_V']).toBe('3');
    });

    it('bold + underline + strikethrough → Style=13', () => {
        const section = createCharacterSection({ bold: true, underline: true, strikethrough: true });
        const styleCell = section.Row![0].Cell.find((c: any) => c['@_N'] === 'Style');
        expect(styleCell['@_V']).toBe('13');
    });

    it('all four flags → Style=15', () => {
        const section = createCharacterSection({ bold: true, italic: true, underline: true, strikethrough: true });
        const styleCell = section.Row![0].Cell.find((c: any) => c['@_N'] === 'Style');
        expect(styleCell['@_V']).toBe('15');
    });
});

// ---------------------------------------------------------------------------
// createParagraphSection — spacing cells
// ---------------------------------------------------------------------------

describe('createParagraphSection paragraph spacing', () => {
    it('horzAlign still works', () => {
        const section = createParagraphSection({ horzAlign: 'center' });
        const cell = section.Row![0].Cell.find((c: any) => c['@_N'] === 'HorzAlign');
        expect(cell['@_V']).toBe('1');
    });

    it('spaceBefore converts points to inches', () => {
        const section = createParagraphSection({ spaceBefore: 72 }); // 72pt = 1 inch
        const cell = section.Row![0].Cell.find((c: any) => c['@_N'] === 'SpBefore');
        expect(parseFloat(cell['@_V'])).toBeCloseTo(1.0, 5);
        expect(cell['@_U']).toBe('PT');
    });

    it('spaceAfter converts points to inches', () => {
        const section = createParagraphSection({ spaceAfter: 36 }); // 36pt = 0.5 inch
        const cell = section.Row![0].Cell.find((c: any) => c['@_N'] === 'SpAfter');
        expect(parseFloat(cell['@_V'])).toBeCloseTo(0.5, 5);
    });

    it('lineSpacing 1.0 stores -1.0 (proportional single spacing)', () => {
        const section = createParagraphSection({ lineSpacing: 1.0 });
        const cell = section.Row![0].Cell.find((c: any) => c['@_N'] === 'SpLine');
        expect(parseFloat(cell['@_V'])).toBeCloseTo(-1.0, 5);
    });

    it('lineSpacing 1.5 stores -1.5', () => {
        const section = createParagraphSection({ lineSpacing: 1.5 });
        const cell = section.Row![0].Cell.find((c: any) => c['@_N'] === 'SpLine');
        expect(parseFloat(cell['@_V'])).toBeCloseTo(-1.5, 5);
    });

    it('lineSpacing 2.0 stores -2.0 (double spacing)', () => {
        const section = createParagraphSection({ lineSpacing: 2.0 });
        const cell = section.Row![0].Cell.find((c: any) => c['@_N'] === 'SpLine');
        expect(parseFloat(cell['@_V'])).toBeCloseTo(-2.0, 5);
    });

    it('empty props produces no cells', () => {
        const section = createParagraphSection({});
        expect(section.Row![0].Cell).toHaveLength(0);
    });
});

// ---------------------------------------------------------------------------
// createTextBlockSection — margin cells
// ---------------------------------------------------------------------------

describe('createTextBlockSection margins', () => {
    it('topMargin is stored in inches with U=IN', () => {
        const section = createTextBlockSection({ topMargin: 0.1 });
        const cell = section.Cell!.find((c: any) => c['@_N'] === 'TopMargin');
        expect(parseFloat(cell['@_V'])).toBeCloseTo(0.1, 5);
        expect(cell['@_U']).toBe('IN');
    });

    it('bottomMargin, leftMargin, rightMargin are stored correctly', () => {
        const section = createTextBlockSection({ bottomMargin: 0.2, leftMargin: 0.05, rightMargin: 0.05 });
        const cells = section.Cell!;
        expect(cells.find((c: any) => c['@_N'] === 'BottomMargin')['@_V']).toBe('0.2');
        expect(cells.find((c: any) => c['@_N'] === 'LeftMargin')['@_V']).toBe('0.05');
        expect(cells.find((c: any) => c['@_N'] === 'RightMargin')['@_V']).toBe('0.05');
    });

    it('section name is TextBlock', () => {
        const section = createTextBlockSection({ topMargin: 0 });
        expect(section['@_N']).toBe('TextBlock');
    });

    it('empty props produces no cells', () => {
        const section = createTextBlockSection({});
        expect(section.Cell).toHaveLength(0);
    });
});

// ---------------------------------------------------------------------------
// Integration: NewShapeProps at creation time
// ---------------------------------------------------------------------------

describe('Shape creation with rich text props', () => {
    it('italic flag is written to the Character section', async () => {
        const doc = await VisioDocument.create();
        const page = doc.pages[0];
        await page.addShape({ text: 'Hi', x: 1, y: 1, width: 2, height: 1, italic: true });

        const buf = await doc.save();
        const reloaded = await VisioDocument.load(buf);
        const shapes = reloaded.pages[0].getShapes();
        const charSection = shapes[0].internalShape.Sections?.['Character'];
        const styleVal = parseInt(charSection?.Rows?.[0]?.Cells?.['Style']?.V ?? '0');
        expect(styleVal & 2).toBe(2); // italic bit
    });

    it('underline + strikethrough flags combine correctly', async () => {
        const doc = await VisioDocument.create();
        const page = doc.pages[0];
        await page.addShape({ text: 'X', x: 1, y: 1, width: 2, height: 1, underline: true, strikethrough: true });

        const buf = await doc.save();
        const reloaded = await VisioDocument.load(buf);
        const shapes = reloaded.pages[0].getShapes();
        const charSection = shapes[0].internalShape.Sections?.['Character'];
        const styleVal = parseInt(charSection?.Rows?.[0]?.Cells?.['Style']?.V ?? '0');
        expect(styleVal & 4).toBe(4);  // underline
        expect(styleVal & 8).toBe(8);  // strikethrough
    });

    it('lineSpacing is written to the Paragraph section', async () => {
        const doc = await VisioDocument.create();
        const page = doc.pages[0];
        await page.addShape({ text: 'Y', x: 1, y: 1, width: 2, height: 1, lineSpacing: 1.5 });

        const buf = await doc.save();
        const reloaded = await VisioDocument.load(buf);
        const shapes = reloaded.pages[0].getShapes();
        const paraSection = shapes[0].internalShape.Sections?.['Paragraph'];
        const spLine = parseFloat(paraSection?.Rows?.[0]?.Cells?.['SpLine']?.V ?? '0');
        expect(spLine).toBeCloseTo(-1.5, 5);
    });

    it('textMarginTop/Left are written to the TextBlock section', async () => {
        const doc = await VisioDocument.create();
        const page = doc.pages[0];
        await page.addShape({ text: 'Z', x: 1, y: 1, width: 2, height: 1, textMarginTop: 0.1, textMarginLeft: 0.05 });

        const buf = await doc.save();
        const reloaded = await VisioDocument.load(buf);
        const shapes = reloaded.pages[0].getShapes();
        const tbSection = shapes[0].internalShape.Sections?.['TextBlock'];
        expect(parseFloat(tbSection?.Cells?.['TopMargin']?.V ?? '0')).toBeCloseTo(0.1, 5);
        expect(parseFloat(tbSection?.Cells?.['LeftMargin']?.V ?? '0')).toBeCloseTo(0.05, 5);
    });
});

// ---------------------------------------------------------------------------
// Integration: shape.setStyle() post-creation
// ---------------------------------------------------------------------------

describe('shape.setStyle() rich text fields', () => {
    it('italic can be applied post-creation', async () => {
        const doc = await VisioDocument.create();
        const page = doc.pages[0];
        const shape = await page.addShape({ text: 'A', x: 1, y: 1, width: 2, height: 1 });
        await shape.setStyle({ italic: true });

        const buf = await doc.save();
        const reloaded = await VisioDocument.load(buf);
        const charSection = reloaded.pages[0].getShapes()[0].internalShape.Sections?.['Character'];
        const styleVal = parseInt(charSection?.Rows?.[0]?.Cells?.['Style']?.V ?? '0');
        expect(styleVal & 2).toBe(2);
    });

    it('lineSpacing can be applied post-creation', async () => {
        const doc = await VisioDocument.create();
        const page = doc.pages[0];
        const shape = await page.addShape({ text: 'B', x: 1, y: 1, width: 2, height: 1 });
        await shape.setStyle({ lineSpacing: 2.0 });

        const buf = await doc.save();
        const reloaded = await VisioDocument.load(buf);
        const paraSection = reloaded.pages[0].getShapes()[0].internalShape.Sections?.['Paragraph'];
        const spLine = parseFloat(paraSection?.Rows?.[0]?.Cells?.['SpLine']?.V ?? '0');
        expect(spLine).toBeCloseTo(-2.0, 5);
    });

    it('text margins can be applied post-creation', async () => {
        const doc = await VisioDocument.create();
        const page = doc.pages[0];
        const shape = await page.addShape({ text: 'C', x: 1, y: 1, width: 2, height: 1 });
        await shape.setStyle({ textMarginTop: 0.15, textMarginBottom: 0.15, textMarginLeft: 0.1, textMarginRight: 0.1 });

        const buf = await doc.save();
        const reloaded = await VisioDocument.load(buf);
        const tbSection = reloaded.pages[0].getShapes()[0].internalShape.Sections?.['TextBlock'];
        expect(parseFloat(tbSection?.Cells?.['TopMargin']?.V ?? '0')).toBeCloseTo(0.15, 5);
        expect(parseFloat(tbSection?.Cells?.['BottomMargin']?.V ?? '0')).toBeCloseTo(0.15, 5);
    });

    it('setStyle rich text + existing fill color coexist', async () => {
        const doc = await VisioDocument.create();
        const page = doc.pages[0];
        const shape = await page.addShape({ text: 'D', x: 1, y: 1, width: 2, height: 1, fillColor: '#ff0000' });
        await shape.setStyle({ italic: true, lineSpacing: 1.5, textMarginTop: 0.1 });

        const buf = await doc.save();
        const reloaded = await VisioDocument.load(buf);
        const s = reloaded.pages[0].getShapes()[0];
        // Fill section still present
        expect(s.internalShape.Sections?.['Fill']).toBeDefined();
        // Character italic bit set
        const styleVal = parseInt(s.internalShape.Sections?.['Character']?.Rows?.[0]?.Cells?.['Style']?.V ?? '0');
        expect(styleVal & 2).toBe(2);
    });
});
