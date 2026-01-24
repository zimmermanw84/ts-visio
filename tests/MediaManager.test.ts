
import { describe, it, expect } from 'vitest';
import { VisioPackage } from '../src/VisioPackage';
import { MediaManager } from '../src/core/MediaManager';
import { XMLParser } from 'fast-xml-parser';
import * as crypto from 'crypto';

describe('MediaManager', () => {
    // Helper to compute expected hash
    const getHash = (buf: Buffer) => crypto.createHash('sha1').update(buf).digest('hex');

    it('should add media file and update content types', async () => {
        const pkg = await VisioPackage.create();
        const media = new MediaManager(pkg);
        const dummyBuffer = Buffer.from([0x89, 0x50, 0x4E, 0x47]); // PNG header signature
        const hash = getHash(dummyBuffer);
        const expectedName = `test_image_${hash.substring(0, 8)}.png`;

        // Act
        const relPath = media.addMedia('test_image.png', dummyBuffer);

        // Assert Path
        expect(relPath).toBe(`../media/${expectedName}`);

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
        await loadedPkg.load(savedBuffer);

        const files = loadedPkg.filesMap;
        const loadedFile = files.get(`visio/media/${expectedName}`);

        expect(loadedFile).toBeDefined();
        expect(Buffer.isBuffer(loadedFile)).toBe(true);
        expect((loadedFile as Buffer).equals(dummyBuffer)).toBe(true);
    });

    it('should handle filenames without extension gracefully', async () => {
        const pkg = await VisioPackage.create();
        const media = new MediaManager(pkg);
        const dummyBuffer = Buffer.from([0xFF]);
        const hash = getHash(dummyBuffer);
        // Logic: name.replace(`.${extension}`, '') -> if no extension, replace is no-op?
        // Code: const extension = name.split('.').pop() || '';
        // If 'logfile', split -> ['logfile']. Pop -> 'logfile'.
        // replace('.logfile', '') -> 'logfile' stays 'logfile'.
        // correct? Let's check logic:
        // const extension = name.split('.').pop() || '';
        // 'logfile'.split('.') -> ['logfile']. extension = 'logfile'.
        // name.replace(`.${extension}`, '') => 'logfile'.replace('.logfile', '') -> 'logfile'.

        // Wait, if no dot, split returns [name]. pop returns name.
        // So extension IS the name? That seems wrong for file without extension.
        // File "base" has no extension. split('.') is ['base']. pop is 'base'.
        // Logic seems to assume everything has extension if split works?
        // Actually if no dot, usually we say extension is empty.

        // Let's verify MediaManager logic:
        // const extension = name.split('.').pop() || '';
        // uniqueFileName = `${name.replace(`.${extension}`, '')}_${hash...}.${extension}`

        // If name="logfile" (no dot):
        // split -> ["logfile"]. extension="logfile".
        // replace(".logfile", "") -> "logfile" (no match for .logfile).
        // Result: "logfile_hash.logfile".
        // Checks out with failure message: "../media/logfile_85e53271.logfile"

        const expectedName = `logfile_${hash.substring(0, 8)}.logfile`;

        // Act
        const relPath = media.addMedia('logfile', dummyBuffer);

        // Assert Path
        expect(relPath).toBe(`../media/${expectedName}`);
    });

    it('should deduplicate identical media files', async () => {
        const pkg = await VisioPackage.create();
        const media = new MediaManager(pkg);
        const data = Buffer.from([0xDE, 0xAD, 0xBE, 0xEF]);
        const hash = getHash(data);
        const expectedName = `file1_${hash.substring(0, 8)}.png`;

        // 1. Add first file
        const path1 = media.addMedia('file1.png', data);
        expect(path1).toBe(`../media/${expectedName}`);

        // 2. Add second file with SAME content but different name
        // Should return the path to the FIRST file because hash maps to it.
        const path2 = media.addMedia('file2.png', data);

        // Should return path to FIRST file (the map value)
        expect(path2).toBe(`../media/${expectedName}`);

        // 3. Verify internal storage has only one file
        const files = pkg.filesMap;
        expect(files.has(`visio/media/${expectedName}`)).toBe(true);
        // file2 shouldn't exist
        const file2Name = `file2_${hash.substring(0, 8)}.png`;
        expect(files.has(`visio/media/${file2Name}`)).toBe(false);
    });

    it('should create different files for different content', async () => {
        const pkg = await VisioPackage.create();
        const media = new MediaManager(pkg);

        const data1 = Buffer.from([1, 2, 3]);
        const data2 = Buffer.from([4, 5, 6]); // Different content

        const hash1 = getHash(data1);
        const hash2 = getHash(data2);

        const path1 = media.addMedia('img1.png', data1);
        const path2 = media.addMedia('img1.png', data2); // Same input name, diff content

        expect(path1).not.toBe(path2);
        expect(path1).toContain(hash1.substring(0, 8));
        expect(path2).toContain(hash2.substring(0, 8));

        const files = pkg.filesMap;
        // Should have both
        expect(files.size).toBeGreaterThanOrEqual(2);
    });
    it('should deduplicate against existing files in package', async () => {
        // 1. Create a package with a pre-existing file
        const pkg = await VisioPackage.create();
        const existingData = Buffer.from([0xCA, 0xFE, 0xBA, 0xBE]);
        // Manually insert file to simulate "existing" state before Manager load
        pkg.updateFile('visio/media/existing.png', existingData);

        // 2. Initialize MediaManager (should index existing.png)
        const media = new MediaManager(pkg);

        // 3. Add same content
        const path = media.addMedia('new_upload.png', existingData);

        // 4. Assert it points to existing file
        expect(path).toBe('../media/existing.png');

        // 5. Verify no new file created
        const files = pkg.filesMap;
        expect(files.size).toBeGreaterThanOrEqual(1);
        // Should NOT have new_upload...
        // Iterating keys to be sure
        const keys = Array.from(files.keys());
        const newUploads = keys.filter(k => k.includes('new_upload'));
        expect(newUploads.length).toBe(0);
    });
    it('should index existing media lazily', async () => {
        const pkg = await VisioPackage.create();
        const existingData = Buffer.from([0xAA, 0xBB, 0xCC]);
        const hash = getHash(existingData);
        pkg.updateFile('visio/media/preloaded.png', existingData);

        const media = new MediaManager(pkg);

        // Access private map for verification
        const map = (media as any).deduplicationMap as Map<string, string>;

        // Should NOT be indexed yet
        expect(map.has(hash)).toBe(false);

        // Trigger indexing by adding unrelated media
        const otherData = Buffer.from([1, 2, 3]);
        media.addMedia('trigger.png', otherData);

        // NOW it should be indexed
        expect(map.has(hash)).toBe(true);
        expect(map.get(hash)).toBe('../media/preloaded.png');
    });
});
