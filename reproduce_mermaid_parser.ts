
import { parse } from '@mermaid-js/parser';

async function main() {
    console.log("Attempting to use @mermaid-js/parser...");
    try {
        const text = `graph TD; A-->B;`;
        const result = await parse(text);
        console.log("Success:", result);
    } catch (e) {
        console.error("Failed:", e);
    }
}

main();
