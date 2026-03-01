import { describe, it, expect, beforeEach } from 'vitest';
import JSZip from 'jszip';
import { VisioPackage } from '../src/VisioPackage';

describe('VisioPackage', () => {
    let pkg: VisioPackage;
    let mockVsdxBuffer: Buffer;

    beforeEach(async () => {
        pkg = new VisioPackage();

        // Create a mock VSDX file (which is just a zip)
        const zip = new JSZip();
        zip.file('[Content_Types].xml', '<Types></Types>');
        zip.file('docProps/app.xml', '<Properties><Application>Microsoft Visio</Application></Properties>');
        zip.file('visio/pages/page1.xml', '<Page></Page>');
        zip.folder('custom');

        mockVsdxBuffer = await zip.generateAsync({ type: 'nodebuffer' });
    });

    it('should load a vsdx file buffer and preload files', async () => {
        await pkg.load(mockVsdxBuffer);

        // Check if getFileText works synchronously
        const content = pkg.getFileText('[Content_Types].xml');
        expect(content).toBe('<Types></Types>');
    });

    it('should retrieve specific xml files', async () => {
        await pkg.load(mockVsdxBuffer);
        const appXml = pkg.getFileText('docProps/app.xml');
        expect(appXml).toContain('Microsoft Visio');
    });

    it('should throw error for non-existent files', async () => {
        await pkg.load(mockVsdxBuffer);
        expect(() => pkg.getFileText('non/existent/file.xml')).toThrow('File not found');
    });

    it('should load multiple files correctly', async () => {
        await pkg.load(mockVsdxBuffer);
        expect(pkg.getFileText('visio/pages/page1.xml')).toBe('<Page></Page>');
        expect(pkg.getFileText('docProps/app.xml')).toBe('<Properties><Application>Microsoft Visio</Application></Properties>');
    });

    it('should load EMF files as Buffer, not corrupted string (BUG 11)', async () => {
        // Binary bytes that would be corrupted if decoded as UTF-8
        const emfBytes = Buffer.from([0x01, 0x00, 0x00, 0x00, 0x80, 0xFF, 0xFE, 0xD8]);

        const zip = new JSZip();
        zip.file('visio/media/chart.emf', emfBytes);
        const buffer = await zip.generateAsync({ type: 'nodebuffer' });

        const p = new VisioPackage();
        await p.load(buffer);

        const loaded = p.filesMap.get('visio/media/chart.emf');
        expect(Buffer.isBuffer(loaded)).toBe(true);
        expect((loaded as Buffer).equals(emfBytes)).toBe(true);
    });

    it('should load WMF and other binary assets as Buffer (BUG 11)', async () => {
        const binaryFormats: Record<string, Buffer> = {
            'visio/media/logo.wmf': Buffer.from([0xD7, 0xCD, 0xC6, 0x9A]),
            'visio/media/object.ole': Buffer.from([0xD0, 0xCF, 0x11, 0xE0]),
            'visio/media/data.bin': Buffer.from([0x00, 0x01, 0x02, 0x03]),
            'visio/media/icon.ico': Buffer.from([0x00, 0x00, 0x01, 0x00]),
        };

        const zip = new JSZip();
        for (const [path, bytes] of Object.entries(binaryFormats)) {
            zip.file(path, bytes);
        }
        const buffer = await zip.generateAsync({ type: 'nodebuffer' });

        const p = new VisioPackage();
        await p.load(buffer);

        for (const [path, bytes] of Object.entries(binaryFormats)) {
            const loaded = p.filesMap.get(path);
            expect(Buffer.isBuffer(loaded), `${path} should be a Buffer`).toBe(true);
            expect((loaded as Buffer).equals(bytes), `${path} content should be intact`).toBe(true);
        }
    });

    it('should still load XML files as strings (BUG 11 regression)', async () => {
        const zip = new JSZip();
        zip.file('visio/pages/page1.xml', '<Page></Page>');
        zip.file('visio/_rels/document.xml.rels', '<Relationships></Relationships>');
        const buffer = await zip.generateAsync({ type: 'nodebuffer' });

        const p = new VisioPackage();
        await p.load(buffer);

        expect(typeof p.filesMap.get('visio/pages/page1.xml')).toBe('string');
        expect(typeof p.filesMap.get('visio/_rels/document.xml.rels')).toBe('string');
    });

    it('should handle reload correctly', async () => {
        await pkg.load(mockVsdxBuffer);
        expect(pkg.getFileText('[Content_Types].xml')).toBe('<Types></Types>');

        const zip2 = new JSZip();
        zip2.file('new.xml', '<New></New>');
        const buffer2 = await zip2.generateAsync({ type: 'nodebuffer' });

        await pkg.load(buffer2);
        expect(pkg.getFileText('new.xml')).toBe('<New></New>');
        expect(() => pkg.getFileText('[Content_Types].xml')).toThrow(); // Should be cleared
    });
});
