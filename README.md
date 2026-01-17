# js-visio

> [!WARNING]
> **Under Construction**
> This library is currently being developed with heavy assistance from AI and is primarily an experimental project. Use with caution.


A Node.js library to strict-type interact with Visio (`.vsdx`) files.
Built using specific schema-level abstractions to handle the complex internal structure of Visio documents (ShapeSheets, Pages, Masters).

> **Status**: Work In Progress (TDD).

## Features

- **Read VSDX**: Open and parse `.vsdx` files (zippped XML).
- **Strict Typing**: Interact with `VisioPage`, `VisioShape`, and `VisioConnect` objects.
- **ShapeSheet Access**: Read `Cells`, `Rows`, and `Sections` directly.
- **Connections**: Analyze connectivity between shapes.
- **Modify & Save**: Update internal XML and save back to a buffer.

## Installation

```bash
npm install js-visio
```

## Usage

### Loading a File

```typescript
import fs from 'fs';
import { VsdxLoader } from 'js-visio';

const run = async () => {
    const loader = new VsdxLoader();
    const buffer = fs.readFileSync('diagram.vsdx');

    await loader.load(buffer);

    // ... interact with the file
};
```

### Inspecting Pages and Shapes

```typescript
const pages = await loader.getPages();
console.log(`Found ${pages.length} pages`);

for (const page of pages) {
    console.log(`Page: ${page.Name} (ID: ${page.ID})`);

    // Get Shapes for this page
    // Note: You need to know the path to the page XML.
    // In standard VSDX, "Page-1" is usually "visio/pages/page1.xml" but parsing logic is improving.
    const shapes = await loader.getPageShapes('visio/pages/page1.xml');

    for (const shape of shapes) {
        console.log(`  Shape: ${shape.Name} (Text: ${shape.Text})`);

        // Access ShapeSheet Cells
        if (shape.Cells['Width']) {
            console.log(`    Width: ${shape.Cells['Width'].V}`);
        }

        // Access Geometry Sections
        if (shape.Sections['Geometry']) {
            console.log(`    Geometry Rows: ${shape.Sections['Geometry'].Rows.length}`);
        }
    }
}
```

### Analyzing Connections

```typescript
const connects = await loader.getPageConnects('visio/pages/page1.xml');

for (const conn of connects) {
    console.log(`Shape ${conn.FromSheet} connects to Shape ${conn.ToSheet}`);
}
```

### Saving Changes

```typescript
// Low-level XML modification (High-level APIs coming soon)
await loader.setFileContent('visio/pages/page1.xml', updatedXmlString);

const newBuffer = await loader.save();
fs.writeFileSync('updated_diagram.vsdx', newBuffer);
```

## Development

This project uses **Vitest** for testing.

```bash
npm install
npm test
```
