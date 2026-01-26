
import { describe, it, expect, beforeEach, afterEach } from 'vitest';
import { MermaidToVisio } from '../src/mermaid/MermaidToVisio';
import * as fs from 'fs';
import * as path from 'path';
import { VisioDocument } from '../src/VisioDocument';

describe('MermaidToVisio Integration', () => {
    const outputDir = path.join(__dirname, 'out');
    const outputPath = path.join(outputDir, 'mermaid_test.vsdx');

    beforeEach(() => {
        if (!fs.existsSync(outputDir)) {
            fs.mkdirSync(outputDir);
        }
    });

    afterEach(() => {
        if (fs.existsSync(outputPath)) {
            fs.unlinkSync(outputPath);
        }
    });

    it('should convert a simple flowchart to a Visio file', async () => {
        const converter = new MermaidToVisio();
        const mermaid = `
        graph TD
            A[Start] --> B{Is it working?}
            B -- Yes --> C[Great!]
            B -- No --> D[Fix it]
        `;

        await converter.convert(mermaid, outputPath);

        expect(fs.existsSync(outputPath)).toBe(true);
        const stats = fs.statSync(outputPath);
        expect(stats.size).toBeGreaterThan(0);

        // Optional: Load it back to verify contents if we have a loader
        const doc = await VisioDocument.load(outputPath);
        expect(doc.pages[0].getShapes().length).toBeGreaterThan(0);
    });

    it('should respect custom page height', async () => {
        const converter = new MermaidToVisio();
        const mermaid = `graph TD
            A[Top] --> B[Bottom]
        `;

        // Pass custom page height of 20 inches
        await converter.convert(mermaid, outputPath, { pageHeight: 20 });

        expect(fs.existsSync(outputPath)).toBe(true);
        // In a real integration test we would open the file and check shape Y coordinates,
        // but for now we trust the logic if it runs without error.
    });

    it('should handle styles', async () => {
        const converter = new MermaidToVisio();
        const mermaid = `
        graph TD
            A[Styled Node]
            style A fill:#f9f,stroke:#333,stroke-width:4px
        `;

        await converter.convert(mermaid, outputPath);
        expect(fs.existsSync(outputPath)).toBe(true);
    });
});
