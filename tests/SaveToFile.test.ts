import { describe, it, expect, afterEach } from 'vitest';
import { VisioPackage } from '../src/VisioPackage';
import fs from 'fs';
import path from 'path';

describe('VisioPackage Save', () => {
    const testFile = path.resolve(__dirname, 'test_output.vsdx');

    afterEach(() => {
        if (fs.existsSync(testFile)) {
            fs.unlinkSync(testFile);
        }
    });

    it('should save to buffer', async () => {
        const pkg = await VisioPackage.create();
        const buffer = await pkg.save();
        expect(Buffer.isBuffer(buffer)).toBe(true);
    });

    it('should save to file if filename provided', async () => {
        const pkg = await VisioPackage.create();
        const buffer = await pkg.save(testFile);

        expect(Buffer.isBuffer(buffer)).toBe(true);
        expect(fs.existsSync(testFile)).toBe(true);

        // Read file back to verify content
        const fileContent = fs.readFileSync(testFile);
        expect(fileContent.equals(buffer)).toBe(true);
    });
});
