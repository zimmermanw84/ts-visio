import { describe, it, expect } from 'vitest';
import { VisioDocument } from '../src/VisioDocument';
import { XMLParser } from 'fast-xml-parser';

const parser = new XMLParser({ ignoreAttributes: false, attributeNamePrefix: '@_' });

async function getPageXml(doc: VisioDocument, pageIndex = 0): Promise<any> {
    const buf = await doc.save();
    const JSZip = (await import('jszip')).default;
    const zip = await JSZip.loadAsync(buf);
    const pageName = `visio/pages/page${pageIndex + 1}.xml`;
    const file = zip.file(pageName);
    if (!file) return null;
    const xml = await file.async('string');
    return parser.parse(xml);
}

function getAllShapeIds(parsed: any): string[] {
    const ids: string[] = [];
    const gather = (shapes: any) => {
        if (!shapes?.Shape) return;
        const arr = Array.isArray(shapes.Shape) ? shapes.Shape : [shapes.Shape];
        for (const s of arr) {
            ids.push(s['@_ID']);
            if (s.Shapes) gather(s.Shapes);
        }
    };
    gather(parsed?.PageContents?.Shapes);
    return ids;
}

function getConnectSheetIds(parsed: any): string[] {
    const connects = parsed?.PageContents?.Connects?.Connect;
    if (!connects) return [];
    const arr = Array.isArray(connects) ? connects : [connects];
    return arr.flatMap((c: any) => [c['@_FromSheet'], c['@_ToSheet']].filter(Boolean));
}

describe('deleteShape', () => {
    it('removes the shape from the page', async () => {
        const doc = await VisioDocument.create();
        const page = doc.pages[0];
        const s1 = await page.addShape({ text: 'A', x: 1, y: 1, width: 1, height: 1 });
        await page.addShape({ text: 'B', x: 3, y: 1, width: 1, height: 1 });

        await s1.delete();

        const parsed = await getPageXml(doc);
        const ids = getAllShapeIds(parsed);
        expect(ids).not.toContain(s1.id);
        expect(ids).toHaveLength(1);
    });

    it('throws when shape does not exist', async () => {
        const doc = await VisioDocument.create();
        const page = doc.pages[0];
        const s = await page.addShape({ text: 'X', x: 1, y: 1, width: 1, height: 1 });
        await s.delete();

        // Deleting again should throw
        await expect(s.delete()).rejects.toThrow();
    });

    it('removes associated Connect elements when deleting a connector endpoint', async () => {
        const doc = await VisioDocument.create();
        const page = doc.pages[0];
        const s1 = await page.addShape({ text: 'A', x: 1, y: 1, width: 1, height: 1 });
        const s2 = await page.addShape({ text: 'B', x: 4, y: 1, width: 1, height: 1 });
        await page.connectShapes(s1, s2);

        await s1.delete();

        const parsed = await getPageXml(doc);
        const sheetIds = getConnectSheetIds(parsed);
        expect(sheetIds).not.toContain(s1.id);
    });

    it('removes associated Connect elements when deleting the other endpoint', async () => {
        const doc = await VisioDocument.create();
        const page = doc.pages[0];
        const s1 = await page.addShape({ text: 'A', x: 1, y: 1, width: 1, height: 1 });
        const s2 = await page.addShape({ text: 'B', x: 4, y: 1, width: 1, height: 1 });
        await page.connectShapes(s1, s2);

        await s2.delete();

        const parsed = await getPageXml(doc);
        const sheetIds = getConnectSheetIds(parsed);
        expect(sheetIds).not.toContain(s2.id);
    });

    it('removes container Relationship entries referencing deleted shape', async () => {
        const doc = await VisioDocument.create();
        const page = doc.pages[0];
        const container = await page.addContainer({ text: 'Box', x: 3, y: 3, width: 4, height: 4 });
        const member = await page.addShape({ text: 'Member', x: 3, y: 3, width: 1, height: 1 });
        await container.addMember(member);

        await member.delete();

        const parsed = await getPageXml(doc);
        const rels = parsed?.PageContents?.Relationships?.Relationship;
        if (rels) {
            const arr = Array.isArray(rels) ? rels : [rels];
            for (const r of arr) {
                expect(r['@_RelatedShapeID']).not.toBe(member.id);
            }
        }
    });

    it('deletes a shape nested inside a group', async () => {
        const doc = await VisioDocument.create();
        const page = doc.pages[0];
        const group = await page.addShape({ text: '', x: 2, y: 2, width: 4, height: 4, type: 'Group' });
        const child = await page.addShape({ text: 'Child', x: 2, y: 2, width: 1, height: 1 }, group.id);

        await child.delete();

        const parsed = await getPageXml(doc);
        const ids = getAllShapeIds(parsed);
        expect(ids).not.toContain(child.id);
        expect(ids).toContain(group.id); // parent remains
    });

    it('leaves other shapes untouched after deletion', async () => {
        const doc = await VisioDocument.create();
        const page = doc.pages[0];
        const s1 = await page.addShape({ text: 'Keep1', x: 1, y: 1, width: 1, height: 1 });
        const s2 = await page.addShape({ text: 'Delete', x: 3, y: 1, width: 1, height: 1 });
        const s3 = await page.addShape({ text: 'Keep2', x: 5, y: 1, width: 1, height: 1 });

        await s2.delete();

        const parsed = await getPageXml(doc);
        const ids = getAllShapeIds(parsed);
        expect(ids).toContain(s1.id);
        expect(ids).toContain(s3.id);
        expect(ids).not.toContain(s2.id);
    });

    it('removes Connect entries for child shapes when a group is deleted', async () => {
        const doc = await VisioDocument.create();
        const page = doc.pages[0];
        const outside = await page.addShape({ text: 'Outside', x: 8, y: 1, width: 1, height: 1 });
        const group   = await page.addShape({ text: 'Group',   x: 2, y: 2, width: 4, height: 4, type: 'Group' });
        const child   = await page.addShape({ text: 'Child',   x: 2, y: 2, width: 1, height: 1 }, group.id);

        // Connect an outside shape to a child inside the group
        await page.connectShapes(outside, child);

        // Delete the entire group (which includes the child)
        await group.delete();

        const parsed = await getPageXml(doc);
        const sheetIds = getConnectSheetIds(parsed);
        // Both the group and the child IDs must be purged from Connects
        expect(sheetIds).not.toContain(group.id);
        expect(sheetIds).not.toContain(child.id);
    });

    it('can delete all shapes from a page', async () => {
        const doc = await VisioDocument.create();
        const page = doc.pages[0];
        const s1 = await page.addShape({ text: 'A', x: 1, y: 1, width: 1, height: 1 });
        const s2 = await page.addShape({ text: 'B', x: 3, y: 1, width: 1, height: 1 });

        await s1.delete();
        await s2.delete();

        const parsed = await getPageXml(doc);
        const ids = getAllShapeIds(parsed);
        expect(ids).toHaveLength(0);
    });
});
