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
- **Create Shapes**: Add new rectangular shapes with text to pages.
- **Connect Shapes**: Create dynamic connectors between shapes.

## Installation

```bash
npm install ts-visio
```

## Usage



### 1. Create or Load a Document
The `VisioDocument` class is the main entry point.

```typescript
import { VisioDocument } from 'ts-visio';

// Create a new blank document
const doc = await VisioDocument.create();

// OR Load an existing file
// const doc = await VisioDocument.load('diagram.vsdx');
```

#### 2. Access Pages
Access pages through the `pages` property.

```typescript
const page = doc.pages[0];
console.log(`Editing Page: ${page.name}`);
```

#### 3. Add & Modify Shapes
Add new shapes or modify existing ones without dealing with XML.

```typescript
// Add a new rectangle shape
const shape = await page.addShape({
    text: "Hello World",
    x: 1,
    y: 1,
    width: 3,
    height: 1,
    fillColor: "#ff0000", // Option hexadecimal fill color
    fontColor: "#ffffff",
    bold: true
});

// Modify text
await shape.setText("Updated Text");

console.log(`Shape ID: ${shape.id}`);
```



#### 4. Connect Shapes
Link two shapes with a dynamic connector.

```typescript
const shape1 = await page.addShape({ text: "From", x: 2, y: 4, width: 2, height: 1 });
const shape2 = await page.addShape({ text: "To", x: 6, y: 4, width: 2, height: 1 });

await page.connectShapes(shape1, shape2);
```

#### 5. Save the Document
Save the modified document back to disk.

```typescript
await doc.save('updated_diagram.vsdx');
```

## Development

This project uses **Vitest** for testing.

```bash
npm install
npm test
```
