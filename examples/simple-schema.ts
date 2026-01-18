import { VisioDocument } from '../src/VisioDocument';
import { ArrowHeads } from '../src/utils/StyleHelpers';
import * as path from 'path';

async function run() {
    console.log("Generating simple-schema.vsdx...");
    const doc = await VisioDocument.create();
    const page = doc.pages[0];

    // 1. Create Tables
    // Users Table
    // 1. Create Tables
    // Users Table
    const usersTable = await page.addTable(2, 6, "Users", [
        "ID: int [PK]",
        "username: varchar",
        "email: varchar",
        "created_at: timestamp"
    ]);

    // Posts Table
    const postsTable = await page.addTable(8, 6, "Posts", [
        "ID: int [PK]",
        "user_id: int [FK]",
        "title: varchar",
        "content: text",
        "published: boolean"
    ]);

    // Comments Table
    const commentsTable = await page.addTable(8, 2, "Comments", [
        "ID: int [PK]",
        "post_id: int [FK]",
        "user_id: int [FK]",
        "body: text"
    ]);

    // 2. Retrieve Shapes to connect them
    // Refactored: addTable now returns the Shape object directly.

    // 3. Connect Tables (Fluent API)
    // Users -> Posts (One to Many)
    await usersTable.connectTo(postsTable, ArrowHeads.One, ArrowHeads.CrowsFoot);

    // Posts -> Comments (One to Many)
    await postsTable.connectTo(commentsTable, ArrowHeads.One, ArrowHeads.CrowsFoot);

    // Users -> Comments (One to Many)
    await usersTable.connectTo(commentsTable, ArrowHeads.One, ArrowHeads.CrowsFoot);

    // 4. Save
    const outPath = path.resolve(__dirname, 'simple-schema.vsdx');
    await doc.save(outPath);
    console.log(`Saved to ${outPath}`);
}

run().catch(console.error);
