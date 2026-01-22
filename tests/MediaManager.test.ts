
import { describe, it, expect } from 'vitest';
import { VisioPackage } from '../src/VisioPackage';
import { MediaManager } from '../src/core/MediaManager';
import { XMLParser } from 'fast-xml-parser';

describe('MediaManager', () => {
    it('should add media file and update content types', async () => {
        const pkg = await VisioPackage.create();
        const media = new MediaManager(pkg);
        const dummyBuffer = Buffer.from([0x89, 0x50, 0x4E, 0x47]); // PNG header signature

        // Act
        const relPath = media.addMedia('test_image.png', dummyBuffer);

        // Assert Path
        expect(relPath).toBe('../media/test_image.png');

        // Assert Content Types updated
        const ctXml = pkg.getFileText('[Content_Types].xml');
        const parser = new XMLParser({ ignoreAttributes: false, attributeNamePrefix: "@_" });
        const parsed = parser.parse(ctXml);

        const defaults = Array.isArray(parsed.Types.Default) ? parsed.Types.Default : [parsed.Types.Default];
        const pngEntry = defaults.find((d: any) => d['@_Extension'] === 'png');
        expect(pngEntry).toBeDefined();
        expect(pngEntry['@_ContentType']).toBe('image/png');

        // Assert Binary persistence via Round-Trip
        const savedBuffer = await pkg.save();

        // Reload into a fresh package to verify zip integrity
        const loadedPkg = new VisioPackage();
        await loadedPkg.load(savedBuffer); // Buffer is Uint8Array compatible

        // Access internal files map directly for verification?
        // Or simply checking if we can get it?
        // VisioPackage doesn't expose a "getFile" for binary, but internal files map has it.
        // We can cast to any to access private members for testing.
        const files = (loadedPkg as any).files as Map<string, string | Buffer>;
        const loadedFile = files.get('visio/media/test_image.png');

        expect(loadedFile).toBeDefined();
        expect(Buffer.isBuffer(loadedFile)).toBe(true);
        expect((loadedFile as Buffer).equals(dummyBuffer)).toBe(true);
    });

    it('should handle filenames without extension gracefully', async () => {
        const pkg = await VisioPackage.create();
        const media = new MediaManager(pkg);
        const dummyBuffer = Buffer.from([0xFF]);

        // Act
        const relPath = media.addMedia('logfile', dummyBuffer);

        // Assert Path
        expect(relPath).toBe('../media/logfile');

        // Assert Content Types NOT updated (or at least no error)
        const ctXml = pkg.getFileText('[Content_Types].xml');
        const parser = new XMLParser({ ignoreAttributes: false, attributeNamePrefix: "@_" });
        const parsed = parser.parse(ctXml);

        const defaults = Array.isArray(parsed.Types.Default) ? parsed.Types.Default : [parsed.Types.Default];
        // Should not have an entry for 'logfile'
        const logEntry = defaults.find((d: any) => d['@_Extension'] === 'logfile');
        expect(logEntry).toBeUndefined();
    });
});
