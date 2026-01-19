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



#### 4. Groups & Nesting
Create Group shapes and add children to them.

```typescript
// 1. Create a Group container
const group = await page.addShape({
    text: "Group",
    x: 5,
    y: 5,
    width: 4,
    height: 4,
    type: 'Group'
});

// 2. Add child shape directly to the group
// Coordinates are relative to the Group's bottom-left corner
await page.addShape({
    text: "Child",
    x: 1, // Relative X
    y: 1, // Relative Y
    width: 1,
    height: 1
}, group.id);
```


#### 6. Automatic Layout
Easily position shapes relative to each other.

```typescript
const shape1 = await page.addShape({ text: "Step 1", x: 2, y: 5, width: 2, height: 1 });
const shape2 = await page.addShape({ text: "Step 2", x: 0, y: 0, width: 2, height: 1 }); // X,Y ignored if we place it next

// Place Shape 2 to the right of Shape 1 with a 1-inch gap
await shape2.placeRightOf(shape1, { gap: 1 });

// Chain placement
const shape3 = await page.addShape({ text: "Step 3", x: 0, y: 0, width: 2, height: 1 });
await shape3.placeRightOf(shape2);

// Vertical stacking
const shape4 = await page.addShape({ text: "Below", x: 0, y: 0, width: 2, height: 1 });
await shape4.placeBelow(shape1, { gap: 0.5 });
```

#### 7. Fluent API & Chaining
Combine creation, styling, and connection in a clean syntax.

```typescript
const shape1 = await page.addShape({ text: "Start", x: 2, y: 4, width: 2, height: 1 });
const shape2 = await page.addShape({ text: "End", x: 6, y: 4, width: 2, height: 1 });

// Fluent Connection
await shape1.connectTo(shape2);

// Chaining Styles & Connections
await shape1.setStyle({ fillColor: '#00FF00' })
            .connectTo(shape2);
```

#### 5. Advanced Connections
Use specific arrowheads (Crow's Foot, etc.)

```typescript
import { ArrowHeads } from 'ts-visio/utils/StyleHelpers';

await page.connectShapes(shape1, shape2, ArrowHeads.One, ArrowHeads.CrowsFoot);
// OR
await shape1.connectTo(shape2, ArrowHeads.One, ArrowHeads.CrowsFoot);
```

#### 6. Database Tables
Create a compound stacked shape for database tables.

```typescript
const tableShape = await page.addTable(
    5,
    5,
    "Users",
    ["ID: int", "Name: varchar", "Email: varchar"]
);
console.log(tableShape.id); // Access ID
```

#### 8. Typed Schema Builder (Facade)
For ER diagrams, use the `SchemaDiagram` wrapper for simplified semantics.

```typescript
import { VisioDocument, SchemaDiagram } from 'ts-visio';

const doc = await VisioDocument.create();
const page = doc.pages[0];
const schema = new SchemaDiagram(page);

// Add Tables
const users = await schema.addTable('Users', ['id', 'email'], 0, 0);
const posts = await schema.addTable('Posts', ['id', 'user_id'], 5, 0);

// Add Relations (1:N maps to Crow's Foot arrow)
await schema.addRelation(users, posts, '1:N');
```

#### 9. Save the Document
Save the modified document back to disk.

```typescript
await doc.save('updated_diagram.vsdx');
```

#### 10. Using Stencils (Masters)
Load a template that already contains stencils (like "Network" or "Flowchart") and drop shapes by their Master ID.

```typescript
// 1. Load a template with stencils
const doc = await VisioDocument.load('template_with_masters.vsdx');
const page = doc.pages[0];

// 2. Drop a shape using a Master ID
// (You can find IDs in visio/masters/masters.xml of the template)
await page.addShape({
    text: "Router 1",
    x: 2,
    y: 2,
    width: 1,
    height: 1,
    masterId: "5" // ID of the 'Router' master in the template
});

// The library automatically handles the relationships so the file stays valid.
```

#### 11. Multi-Page Documents
Create multiple pages and organize your diagram across them.

```typescript
const doc = await VisioDocument.create();

// Page 1 (Default)
const page1 = doc.pages[0];
await page1.addShape({ text: "Home", x: 1, y: 1 });

// Page 2 (New)
const page2 = await doc.addPage("Architecture Diagram");
await page2.addShape({ text: "Server", x: 4, y: 4 });

await doc.save("multipage.vsdx");
```

## Examples

Check out the [examples](./examples) directory for complete scripts.

- **[Simple Schema](./examples/simple-schema.ts)**: Generates a database schema ERD with tables, styling, and Crow's Foot connectors.

## Development

This project uses **Vitest** for testing.

```bash
npm install
npm test
```
