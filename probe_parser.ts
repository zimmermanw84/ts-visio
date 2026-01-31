
import { parse } from '@a24z/mermaid-parser';

async function main() {
    console.log("Probing @a24z/mermaid-parser...");
    try {
        const text = `
        graph TD
            A[Start] --> B(End)
            style A fill:#f9f
        `;
        const result = await parse(text);
        console.log("Result Keys:", Object.keys(result));
        console.log("Full Result:", JSON.stringify(result, null, 2));
    } catch (e) {
        console.error("Failed:", e);
    }
}

main();
