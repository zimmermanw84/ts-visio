import { describe, it, expect } from 'vitest';
import { VisioDocument } from '../src/VisioDocument';

// ---------------------------------------------------------------------------
// Defaults (freshly created document)
// ---------------------------------------------------------------------------

describe('getMetadata() defaults', () => {
    it('title defaults to the template value "Drawing1"', async () => {
        const doc = await VisioDocument.create();
        const meta = doc.getMetadata();
        expect(meta.title).toBe('Drawing1');
    });

    it('author defaults to "ts-visio"', async () => {
        const doc = await VisioDocument.create();
        const meta = doc.getMetadata();
        expect(meta.author).toBe('ts-visio');
    });

    it('optional fields are undefined when not set', async () => {
        const doc = await VisioDocument.create();
        const meta = doc.getMetadata();
        expect(meta.description).toBeUndefined();
        expect(meta.keywords).toBeUndefined();
        expect(meta.company).toBeUndefined();
        expect(meta.manager).toBeUndefined();
        expect(meta.created).toBeUndefined();
        expect(meta.modified).toBeUndefined();
    });
});

// ---------------------------------------------------------------------------
// setMetadata() — individual fields
// ---------------------------------------------------------------------------

describe('setMetadata() core.xml fields', () => {
    it('sets title and reads it back', async () => {
        const doc = await VisioDocument.create();
        doc.setMetadata({ title: 'My Diagram' });
        expect(doc.getMetadata().title).toBe('My Diagram');
    });

    it('sets author and reads it back', async () => {
        const doc = await VisioDocument.create();
        doc.setMetadata({ author: 'Alice' });
        expect(doc.getMetadata().author).toBe('Alice');
    });

    it('sets description and reads it back', async () => {
        const doc = await VisioDocument.create();
        doc.setMetadata({ description: 'A network diagram' });
        expect(doc.getMetadata().description).toBe('A network diagram');
    });

    it('sets keywords and reads it back', async () => {
        const doc = await VisioDocument.create();
        doc.setMetadata({ keywords: 'network infrastructure cloud' });
        expect(doc.getMetadata().keywords).toBe('network infrastructure cloud');
    });

    it('sets lastModifiedBy and reads it back', async () => {
        const doc = await VisioDocument.create();
        doc.setMetadata({ lastModifiedBy: 'Bob' });
        expect(doc.getMetadata().lastModifiedBy).toBe('Bob');
    });

    it('sets created date and reads it back', async () => {
        const doc = await VisioDocument.create();
        const created = new Date('2024-01-15T10:30:00.000Z');
        doc.setMetadata({ created });
        const meta = doc.getMetadata();
        expect(meta.created).toBeInstanceOf(Date);
        expect(meta.created!.toISOString()).toBe(created.toISOString());
    });

    it('sets modified date and reads it back', async () => {
        const doc = await VisioDocument.create();
        const modified = new Date('2025-06-20T14:00:00.000Z');
        doc.setMetadata({ modified });
        const meta = doc.getMetadata();
        expect(meta.modified).toBeInstanceOf(Date);
        expect(meta.modified!.toISOString()).toBe(modified.toISOString());
    });
});

describe('setMetadata() app.xml fields', () => {
    it('sets company and reads it back', async () => {
        const doc = await VisioDocument.create();
        doc.setMetadata({ company: 'ACME Corp' });
        expect(doc.getMetadata().company).toBe('ACME Corp');
    });

    it('sets manager and reads it back', async () => {
        const doc = await VisioDocument.create();
        doc.setMetadata({ manager: 'Carol' });
        expect(doc.getMetadata().manager).toBe('Carol');
    });
});

// ---------------------------------------------------------------------------
// Partial updates — only specified fields change
// ---------------------------------------------------------------------------

describe('setMetadata() partial updates', () => {
    it('setting title does not wipe author', async () => {
        const doc = await VisioDocument.create();
        doc.setMetadata({ author: 'Alice' });
        doc.setMetadata({ title: 'New Title' });

        const meta = doc.getMetadata();
        expect(meta.title).toBe('New Title');
        expect(meta.author).toBe('Alice');
    });

    it('setting company does not wipe title', async () => {
        const doc = await VisioDocument.create();
        doc.setMetadata({ title: 'My Doc' });
        doc.setMetadata({ company: 'ACME' });

        const meta = doc.getMetadata();
        expect(meta.title).toBe('My Doc');
        expect(meta.company).toBe('ACME');
    });

    it('multiple fields can be set in one call', async () => {
        const doc = await VisioDocument.create();
        doc.setMetadata({
            title: 'T',
            author: 'A',
            description: 'D',
            keywords: 'k1 k2',
            company: 'C',
            manager: 'M',
        });

        const meta = doc.getMetadata();
        expect(meta.title).toBe('T');
        expect(meta.author).toBe('A');
        expect(meta.description).toBe('D');
        expect(meta.keywords).toBe('k1 k2');
        expect(meta.company).toBe('C');
        expect(meta.manager).toBe('M');
    });

    it('subsequent calls overwrite previous values', async () => {
        const doc = await VisioDocument.create();
        doc.setMetadata({ title: 'First' });
        doc.setMetadata({ title: 'Second' });
        expect(doc.getMetadata().title).toBe('Second');
    });
});

// ---------------------------------------------------------------------------
// XML-special characters are escaped / round-trip correctly
// ---------------------------------------------------------------------------

describe('setMetadata() XML escaping', () => {
    it('handles ampersand in title', async () => {
        const doc = await VisioDocument.create();
        doc.setMetadata({ title: 'Sales & Marketing' });
        expect(doc.getMetadata().title).toBe('Sales & Marketing');
    });

    it('handles angle brackets in description', async () => {
        const doc = await VisioDocument.create();
        doc.setMetadata({ description: '<important> note' });
        expect(doc.getMetadata().description).toBe('<important> note');
    });
});

// ---------------------------------------------------------------------------
// Persistence: changes survive save/reload cycle
// ---------------------------------------------------------------------------

describe('DocumentMetadata persistence', () => {
    it('all fields survive save/reload cycle', async () => {
        const doc = await VisioDocument.create();
        const created  = new Date('2024-03-01T08:00:00.000Z');
        const modified = new Date('2024-03-15T12:00:00.000Z');

        doc.setMetadata({
            title: 'Persisted Doc',
            author: 'Dave',
            description: 'Survives reload',
            keywords: 'test persistence',
            lastModifiedBy: 'CI',
            company: 'Widget Co',
            manager: 'Eve',
            created,
            modified,
        });

        const buffer = await doc.save();
        const reloaded = await VisioDocument.load(buffer);
        const meta = reloaded.getMetadata();

        expect(meta.title).toBe('Persisted Doc');
        expect(meta.author).toBe('Dave');
        expect(meta.description).toBe('Survives reload');
        expect(meta.keywords).toBe('test persistence');
        expect(meta.lastModifiedBy).toBe('CI');
        expect(meta.company).toBe('Widget Co');
        expect(meta.manager).toBe('Eve');
        expect(meta.created!.toISOString()).toBe(created.toISOString());
        expect(meta.modified!.toISOString()).toBe(modified.toISOString());
    });
});

// ---------------------------------------------------------------------------
// Export regression
// ---------------------------------------------------------------------------

describe('Public exports', () => {
    it('DocumentMetadata type is importable from the package root', async () => {
        // Type-only check — if this compiles, DocumentMetadata is exported
        const root = await import('../src/index');
        // VisioDocument (which uses DocumentMetadata) must be exported
        expect(root.VisioDocument).toBeDefined();
    });
});
