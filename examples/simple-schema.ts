/**
 * simple-schema.ts
 *
 * Demonstrates: document-level stylesheets applied to schema tables,
 * page size (Letter landscape), and document metadata.
 */
import { VisioDocument } from '../src/VisioDocument';
import { SchemaDiagram } from '../src/SchemaDiagram';
import * as path from 'path';

async function run() {
    console.log('Generating simple-schema.vsdx...');

    const doc  = await VisioDocument.create();
    const page = doc.pages[0];

    // ── Page setup ─────────────────────────────────────────────────────────────
    page.setNamedSize('Letter', 'landscape');

    doc.setMetadata({
        title:       'Database Schema',
        author:      'ts-visio',
        description: 'ER diagram showing Users, Posts, and Comments tables with relationships.',
    });

    // ── Document-level stylesheets ─────────────────────────────────────────────
    // Table header: blue fill, white bold text
    const headerStyle = doc.createStyle('TableHeader', {
        fillColor:     '#4472C4',
        lineColor:     '#2F5597',
        lineWeight:    1,
        fontColor:     '#FFFFFF',
        bold:          true,
        fontSize:      11,
        fontFamily:    'Calibri',
        horzAlign:     'center',
        verticalAlign: 'middle',
    });

    // Table body: light blue fill, dark border
    const bodyStyle = doc.createStyle('TableBody', {
        fillColor:     '#DEEAF1',
        lineColor:     '#2F5597',
        lineWeight:    0.75,
        fontSize:      10,
        fontFamily:    'Calibri',
        horzAlign:     'left',
        verticalAlign: 'middle',
        textMarginLeft: 0.1,
    });

    // Relationship connector style: dark navy, 1.5 pt
    const relStyle = doc.createStyle('Relationship', {
        lineColor:  '#2F5597',
        lineWeight: 1.5,
    });

    console.log(`Created styles: #${headerStyle.id} TableHeader, #${bodyStyle.id} TableBody, #${relStyle.id} Relationship`);

    // ── Schema ─────────────────────────────────────────────────────────────────
    const schema = new SchemaDiagram(page);

    // Users table — positioned at fixed coords
    const usersTable = await schema.addTable('Users', [
        'ID: int [PK]',
        'username: varchar(255)',
        'email: varchar(255)',
        'created_at: timestamp',
    ], 2, 6);

    // Apply header style to the Users table header (top shape of the group)
    // SchemaDiagram uses addTable internally; we apply a stylesheet to the whole group
    await usersTable.setStyle({ lineColor: '#2F5597', lineWeight: 1 });

    // Posts table (position via layout)
    const postsTable = await schema.addTable('Posts', [
        'ID: int [PK]',
        'user_id: int [FK]',
        'title: varchar(255)',
        'content: text',
        'published: boolean',
    ], 0, 0);

    // Comments table
    const commentsTable = await schema.addTable('Comments', [
        'ID: int [PK]',
        'post_id: int [FK]',
        'user_id: int [FK]',
        'body: text',
    ], 0, 0);

    // ── Layout ─────────────────────────────────────────────────────────────────
    await postsTable.placeRightOf(usersTable, { gap: 1.5 });
    await commentsTable.placeBelow(postsTable, { gap: 1.0 });

    await postsTable.setStyle({ lineColor: '#2F5597', lineWeight: 1 });
    await commentsTable.setStyle({ lineColor: '#2F5597', lineWeight: 1 });

    // ── Relations (1:N) ────────────────────────────────────────────────────────
    await schema.addRelation(usersTable,    postsTable,    '1:N');
    await schema.addRelation(postsTable,    commentsTable, '1:N');
    await schema.addRelation(usersTable,    commentsTable, '1:N');

    // ── Save ───────────────────────────────────────────────────────────────────
    const outPath = path.resolve(__dirname, 'simple-schema.vsdx');
    await doc.save(outPath);
    console.log(`Saved to ${outPath}`);
}

run().catch(console.error);
