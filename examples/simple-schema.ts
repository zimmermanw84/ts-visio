import { VisioDocument } from '../src/VisioDocument';
import { ArrowHeads } from '../src/utils/StyleHelpers';
import * as path from 'path';

async function run() {
    console.log("Generating simple-schema.vsdx...");
    const doc = await VisioDocument.create();
    const page = doc.pages[0];

    // 1. Create Tables
    // Users Table
    const usersId = await page.addTable(2, 6, "Users", [
        "ID: int [PK]",
        "username: varchar",
        "email: varchar",
        "created_at: timestamp"
    ]);

    // Posts Table
    const postsId = await page.addTable(8, 6, "Posts", [
        "ID: int [PK]",
        "user_id: int [FK]",
        "title: varchar",
        "content: text",
        "published: boolean"
    ]);

    // Comments Table
    const commentsId = await page.addTable(8, 2, "Comments", [
        "ID: int [PK]",
        "post_id: int [FK]",
        "user_id: int [FK]",
        "body: text"
    ]);

    // 2. Retrieve Shapes to connect them
    // We need shape objects to pass to connectShapes.
    // Since addTable returns an ID, we can fetch the shape objects via the page (conceptually).
    // However, our current API `getShapes()` reads from XML. We just added them, so we might need to save/reload or relies on internal state if we were robust.
    // BUT: The current implementation of `addTable` returns the ID.
    // The `Page` class has `getShapes()`, which reads from the package.
    // Since `addShape` updates the package in memory, `getShapes` should theoretically be able to read it back IF `ShapeReader` parses the current XML.
    // Let's rely on `page.getShapes()` finding them.

    // Helper to find shape by ID
    const findShape = (id: string) => {
        const shapes = page.getShapes();
        return shapes.find(s => s.id === id);
    }

    const usersShape = findShape(usersId);
    const postsShape = findShape(postsId);
    const commentsShape = findShape(commentsId);

    if (!usersShape || !postsShape || !commentsShape) {
        throw new Error("Could not find created shapes to connect.");
    }

    // 3. Connect Tables
    // Users -> Posts (One to Many)
    await page.connectShapes(usersShape, postsShape, ArrowHeads.One, ArrowHeads.CrowsFoot);

    // Posts -> Comments (One to Many)
    await page.connectShapes(postsShape, commentsShape, ArrowHeads.One, ArrowHeads.CrowsFoot);

    // Users -> Comments (One to Many)
    // Connecting from Users to Comments as well (optional, but typical)
    // Note: This might overlap with Posts connection visually without routing logic, but Visio 'Dynamic connector' should handle basic routing.
    await page.connectShapes(usersShape, commentsShape, ArrowHeads.One, ArrowHeads.CrowsFoot);

    // 4. Save
    const outPath = path.resolve(__dirname, 'simple-schema.vsdx');
    await doc.save(outPath);
    console.log(`Saved to ${outPath}`);
}

run().catch(console.error);
