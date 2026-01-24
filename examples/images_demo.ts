
import { VisioDocument } from '../src/VisioDocument';
import * as fs from 'fs';
import * as path from 'path';

async function main() {
    console.log('--- Image Embedding Demo ---');

    // 1. Create a new document
    const doc = await VisioDocument.create();
    const page = doc.pages[0];

    // 2. Load the example image
    const imagePath = path.join(__dirname, 'schmea', 'simple-schema-test.jpg');
    if (!fs.existsSync(imagePath)) {
        console.error(`Error: Example image not found at ${imagePath}`);
        process.exit(1);
    }
    const imageData = fs.readFileSync(imagePath);

    console.log(`Loading image from: ${imagePath}`);

    // 3. Add the image to the page
    // Parameters: data, filename, x, y, width, height
    const imageShape = await page.addImage(imageData, 'simple-schema-test.jpg', 4, 6, 4, 3);

    console.log(`Added image shape with ID: ${imageShape.id}`);

    // 4. Save the document
    const outputPath = path.join(__dirname, 'images_demo.vsdx');
    const buffer = await doc.save();
    fs.writeFileSync(outputPath, buffer);

    console.log(`Saved demo to: ${outputPath}`);
    console.log('--- Done ---');
}

main().catch(err => {
    console.error('Error in demo:', err);
    process.exit(1);
});
