# ts-visio

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
- **Modular Architecture**: Use specialized components for loading, page management, shape reading, and modification.
- **Modify Content**: Update text content of shapes.

## Installation

```bash
npm install ts-visio
```

## Usage

### Simple Loader (Legacy)

```typescript
import fs from 'fs';
import { VsdxLoader } from 'ts-visio';

const run = async () => {
    const loader = new VsdxLoader();
    const buffer = fs.readFileSync('diagram.vsdx');

    await loader.load(buffer);
    // ...
};
```

### Modular API (New)

For more control, use the specialized classes `VisioPackage`, `PageManager`, and `ShapeReader`.

#### 1. Load the Package
`VisioPackage` handles loading the zip file and providing access to internal files.

```typescript
import { VisioPackage } from 'ts-visio';
import fs from 'fs';

const pkg = new VisioPackage();
const buffer = fs.readFileSync('diagram.vsdx');
await pkg.load(buffer);
```

#### 2. Create a Blank Package
You can also create a new, empty Visio document.

```typescript
const pkg = await VisioPackage.create();
```

#### 3. Manage Pages
`PageManager` lists available pages in the document.

```typescript
import { PageManager } from 'ts-visio';

const pageManager = new PageManager(pkg);
const pages = pageManager.getPages();

pages.forEach(page => {
    console.log(`Page: ${page.Name} (ID: ${page.ID})`);
});
```

#### 4. Read Shapes
`ShapeReader` parses shape data from a specific page's XML file.

```typescript
import { ShapeReader } from 'ts-visio';

const shapeReader = new ShapeReader(pkg);

// Typically pages are at 'visio/pages/page{ID}.xml' or similar,
// strictly you should map Page ID to file name, but commonly:
const shapes = shapeReader.readShapes('visio/pages/page1.xml');

shapes.forEach(shape => {
    console.log(`Shape: ${shape.Name}`);
    console.log(`  Text: ${shape.Text}`);

    // Access detailed Cell data
    if (shape.Cells['Width']) {
        console.log(`  Width: ${shape.Cells['Width'].V}`);
    }
});
    }
});
```

#### 5. Update Shapes
`ShapeModifier` allows you to modify shape properties (currently Text) and save the changes back to the package.

```typescript
import { ShapeModifier } from 'ts-visio';

const modifier = new ShapeModifier(pkg);

// Update the text of Shape with ID "1" on Page "1"
await modifier.updateShapeText('1', '1', 'New Text Content');

// Save the changes to a new buffer
const newBuffer = await pkg.save();
fs.writeFileSync('updated_diagram.vsdx', newBuffer);
```

## Development

This project uses **Vitest** for testing.

```bash
npm install
npm test
```
