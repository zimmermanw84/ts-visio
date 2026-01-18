import { VisioDocument } from '../src/VisioDocument';
import { SchemaDiagram } from '../src/SchemaDiagram';
import * as path from 'path';

async function run() {
    console.log("Generating simple-schema.vsdx...");
    const doc = await VisioDocument.create();
    const page = doc.pages[0];
    const schema = new SchemaDiagram(page);

    // 1. Create Tables using Facade
    // Initial Table at specific coords
    const usersTable = await schema.addTable("Users", [
        "ID: int [PK]",
        "username: varchar",
        "email: varchar",
        "created_at: timestamp"
    ], 2, 6);

    // 2. Create other tables at (0,0) and use Layout to position them
    const postsTable = await schema.addTable("Posts", [
        "ID: int [PK]",
        "user_id: int [FK]",
        "title: varchar",
        "content: text",
        "published: boolean"
    ], 0, 0);

    const commentsTable = await schema.addTable("Comments", [
        "ID: int [PK]",
        "post_id: int [FK]",
        "user_id: int [FK]",
        "body: text"
    ], 0, 0);

    // 3. Apply Layout
    // Place Posts to the right of Users
    await postsTable.placeRightOf(usersTable, { gap: 1.5 });

    // Place Comments below Posts
    await commentsTable.placeBelow(postsTable, { gap: 1.0 });

    // 4. Create Relations (Semantic)
    // Users -> Posts (One to Many)
    await schema.addRelation(usersTable, postsTable, '1:N');

    // Posts -> Comments (One to Many)
    await schema.addRelation(postsTable, commentsTable, '1:N');

    // Users -> Comments (One to Many)
    await schema.addRelation(usersTable, commentsTable, '1:N');

    // 5. Save
    const outPath = path.resolve(__dirname, 'simple-schema.vsdx');
    await doc.save(outPath);
    console.log(`Saved to ${outPath}`);
}

run().catch(console.error);
