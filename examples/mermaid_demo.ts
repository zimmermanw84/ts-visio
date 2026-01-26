import * as fs from 'fs';
import * as path from 'path';
import { MermaidConverter } from '../src/mermaid/MermaidConverter';

async function main() {
    const inputPath = path.join(__dirname, '../tests/assets/test-diagram.mmd');
    const outputPath = path.join(__dirname, 'mermaid_output.vsdx');

    console.log(`Reading Mermaid file: ${inputPath}`);
    const mermaidText = fs.readFileSync(inputPath, 'utf-8');

    console.log('Converting to Visio...');
    const converter = new MermaidConverter();
    await converter.convert(mermaidText, outputPath);

    console.log(`Successfully generated: ${outputPath}`);
}

main().catch(console.error);
