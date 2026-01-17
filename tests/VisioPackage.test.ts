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
