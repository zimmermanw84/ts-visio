
// @ts-ignore
import { parse } from '@mermaid-js/parser';

async function main() {
    console.log('Testing @mermaid-js/parser...');
    try {
        const text = `
        graph TD
          A --> B
        `;
        const result = await parse('flowchart', text);
        console.log('Parsed successfully:', JSON.stringify(result, null, 2));
    } catch (e) {
        console.error('Error parsing:', e);
    }
}

main();
