# ts-visio

> [!WARNING]
> **Under Construction**
> This library is currently being developed with heavy assistance from AI and is primarily an experimental project. Use with caution.


A Node.js library to strict-type interact with Visio (`.vsdx`) files.
Built using specific schema-level abstractions to handle the complex internal structure of Visio documents (ShapeSheets, Pages, Masters).

## Features

- **Read VSDX**: Open and parse `.vsdx` files (zipped XML).
- **Strict Typing**: Interact with `VisioPage`, `VisioShape`, and `VisioConnect` objects.
- **ShapeSheet Access**: Read `Cells`, `Rows`, and `Sections` directly.
- **Connections**: Analyze connectivity between shapes.
- **Modular Architecture**: Use specialized components for loading, page management, shape reading, and modification.
- **Modify Content**: Update text content of shapes.
- **Create Shapes**: Rectangles, ellipses, diamonds, rounded rectangles, triangles, parallelograms.
- **Connect Shapes**: Dynamic connectors with arrow styles, line styling, and routing (straight / orthogonal / curved).
- **Text Styling**: Font size, font family, bold, italic, underline, strikethrough, color, alignment, paragraph spacing, and text margins.
- **Shape Transformations**: Rotate, flip (X/Y), and resize shapes via a fluent API.
- **Deletion**: Remove shapes and pages cleanly (including orphaned connectors and relationships).
- **Lookup API**: Find shapes by ID, predicate, or look up pages by name.
- **Read-Back API**: Read custom properties, hyperlinks, and layer assignments from existing shapes.
- **Page Size & Orientation**: Set canvas dimensions with named sizes (`Letter`, `A4`, …) or raw inches; rotate between portrait and landscape.
- **Document Metadata**: Read and write document properties (title, author, description, keywords, company, dates) via `doc.getMetadata()` / `doc.setMetadata()`.
- **Named Connection Points**: Define specific ports on shapes (`Top`, `Right`, etc.) and connect to them precisely using `fromPort`/`toPort` on any connector API.
- **StyleSheets**: Create document-level named styles with fill, line, and text properties via `doc.createStyle()` and apply them to shapes at creation time (`styleId`) or post-creation (`shape.applyStyle()`).
- **Color Palette**: Register named colors in the document's color table via `doc.addColor()` and look them up by index or hex value with `doc.getColors()` / `doc.getColorIndex()`.
- **Read Layers Back**: Enumerate existing layers from loaded files via `page.getLayers()`; delete a layer with `layer.delete()`, rename with `layer.rename()`, and read `layer.visible` / `layer.locked` state.
- **Group Traversal**: Access nested child shapes via `shape.getChildren()`, check `shape.isGroup`, and read `shape.type`.
- **Drawing Scale**: Set a real-world unit mapping on any page (`page.setDrawingScale()`), read it back (`page.getDrawingScale()`), or reset to 1:1 (`page.clearDrawingScale()`). Supports all common imperial and metric units.

Feature gaps are being tracked in [FEATURES.md](./FEATURES.md).

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
    fillColor: "#ff0000",   // Hex fill color
    fontColor: "#ffffff",
    bold: true,
    fontSize: 14,           // Points
    fontFamily: "Segoe UI",
    horzAlign: "center",    // "left" | "center" | "right" | "justify"
    verticalAlign: "middle" // "top" | "middle" | "bottom"
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
import { ArrowHeads } from 'ts-visio';

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

#### 10. Masters & Stencils

**Create a master from scratch** — define a reusable shape with a built-in geometry, then stamp instances onto any page:

```typescript
const doc = await VisioDocument.create();
const page = doc.pages[0];

// Define masters (geometries: 'rectangle' | 'ellipse' | 'diamond' |
//   'rounded-rectangle' | 'triangle' | 'parallelogram')
const boxMaster     = doc.createMaster('Box');
const processMaster = doc.createMaster('Process', 'ellipse');
const decisionMaster = doc.createMaster('Decision', 'diamond');

// Stamp instances — geometry + styling come from the master definition
await page.addShape({ text: 'Start',    x: 1, y: 1, width: 2, height: 1, masterId: processMaster.id });
await page.addShape({ text: 'Step 1',   x: 4, y: 1, width: 2, height: 1, masterId: boxMaster.id });
await page.addShape({ text: 'Branch?',  x: 7, y: 1, width: 2, height: 1, masterId: decisionMaster.id });
```

**Import masters from a `.vssx` stencil file** — bring in an entire stencil and use any of its masters:

```typescript
const doc = await VisioDocument.create();

// Path to a .vssx file, or pass a Buffer / ArrayBuffer directly
const stencilMasters = await doc.importMastersFromStencil('./Network_Shapes.vssx');
const router = stencilMasters.find(m => m.name === 'Router')!;

await doc.pages[0].addShape({
    text: 'Router 1', x: 2, y: 2, width: 1, height: 1,
    masterId: router.id,
});
```

**List masters already in the document** — useful after loading a `.vsdx` that was created in Visio:

```typescript
const doc = await VisioDocument.load('existing_diagram.vsdx');
const masters = doc.getMasters();
// [{ id: '1', name: 'Box', nameU: 'Box', xmlPath: 'visio/masters/master1.xml' }, ...]

// Find and reuse a master by name
const routerMaster = masters.find(m => m.name === 'Router');
if (routerMaster) {
    await doc.pages[0].addShape({ text: 'Router 2', x: 5, y: 5, width: 1, height: 1, masterId: routerMaster.id });
}
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

#### 12. Shape Data (Custom Properties)
Add metadata to shapes, which is crucial for "Smart" diagrams like visual databases or inventories.

```typescript
const shape = await page.addShape({ text: "Server DB-01", x: 2, y: 2 });

// Synchronous Fluent API
shape.addData("IP", { value: "192.168.1.10", label: "IP Address" })
     .addData("Status", { value: "Active" })
     .addData("LastRefreshed", { value: new Date() }) // Auto-serialized as Visio Date
     .addData("ConfigID", { value: 1024, hidden: true }); // Invisible to user
```

#### 13. Image Embedding
Embed PNG or JPEG images directly into the diagram.

```typescript
import * as fs from 'fs';

const buffer = fs.readFileSync('logo.png');
const page = doc.pages[0];

// Add image at (x=2, y=5) with width=3, height=2
await page.addImage(buffer, 'logo.png', 2, 5, 3, 2);
```

#### 14. Containers
Create visual grouping containers (Classic Visio style).

```typescript
// Create a Container
const container = await page.addContainer({
    text: "Network Zone",
    x: 1, y: 1,
    width: 6, height: 4
});

// Add shapes "inside" the container area
// (Visio treats them as members if they are spatially within)
const shape = await page.addShape({ text: "Server A", x: 2, y: 2, width: 1, height: 1 });

// Explicitly link the shape to the container (so they move together)
await container.addMember(shape);
```

#### 15. Lists (Stacking Containers)
Create ordered lists that automatically stack items either vertically (Tables) or horizontally (Timelines).

```typescript
// Vertical List (e.g. Database Table)
const list = await page.addList({
    text: "Users",
    x: 1, y: 10,
    width: 3, height: 1
}, 'vertical'); // 'vertical' or 'horizontal'

// Items stack automatically
await list.addListItem(item1);
await list.addListItem(item2);
```

#### 16. Interactivity & Navigation (Hyperlinks)
Add internal links (drill-downs) or external links (web references). These appear in the right-click menu in Visio.

```typescript
// 1. External URL
await shape.toUrl('https://jira.com/123', 'Open Ticket');

// 2. Internal Page Link
const detailPage = await doc.addPage('Details');
await shape.toPage(detailPage, 'Go to Details');

// 3. Chainable
await shape.toUrl('https://google.com')
           .toPage(detailPage);
```

#### 17. Layers
Organize complex diagrams with layers. Control visibility and locking programmatically, and read layers back from loaded files.

```typescript
// 1. Define Layers
const wireframe   = await page.addLayer('Wireframe');
const annotations = await page.addLayer('Annotations', { visible: false });

// 2. Assign Shapes to Layers
await shape.addToLayer(wireframe);
await noteBox.addToLayer(annotations);

// 3. Toggle Visibility
await annotations.hide();  // Hide for presentation
await annotations.show();  // Show again

// 4. Lock Layer
await wireframe.setLocked(true);

// 5. Read all layers back (works on loaded files too)
const layers = page.getLayers();
// [ { name: 'Wireframe', index: 0, visible: true, locked: true },
//   { name: 'Annotations', index: 1, visible: true, locked: false } ]

for (const layer of layers) {
    console.log(layer.name, layer.visible, layer.locked);
}

// 6. Rename a layer
await wireframe.rename('Structural');

// 7. Delete a layer (cleans up shape assignments automatically)
await annotations.delete();
```

#### 18. Cross-Functional Flowcharts (Swimlanes)
Create structured Swimlane diagrams involving Pools and Lanes.

```typescript
// 1. Create a Pool (Vertical List)
const pool = await page.addSwimlanePool({
    text: "User Registration",
    x: 5, y: 5, width: 10, height: 6
});

// 2. Create Lanes (Containers)
const lane1 = await page.addSwimlaneLane({ text: "Client", width: 10, height: 2 });
const lane2 = await page.addSwimlaneLane({ text: "Server", width: 10, height: 2 });

// 3. Add Lanes to Pool (Order matters)
await pool.addListItem(lane1);
await pool.addListItem(lane2);

// 4. Group Shapes into Lanes
// This binds their movement so they stay inside the lane
await lane1.addMember(startShape);
await lane2.addMember(serverShape);
```

#### 19. Shape Transformations
Rotate, flip, and resize shapes using a fluent API.

```typescript
const shape = await page.addShape({ text: "Widget", x: 3, y: 3, width: 2, height: 1 });

// Rotate 45 degrees (clockwise)
await shape.rotate(45);
console.log(shape.angle); // 45

// Mirror horizontally or vertically
await shape.flipX();
await shape.flipY(false); // un-flip

// Resize (keeps the pin point centred)
await shape.resize(4, 2);
console.log(shape.width, shape.height); // 4, 2

// Chainable
await shape.rotate(90).then(s => s.resize(3, 1));
```

#### 20. Deleting Shapes and Pages
Remove shapes or entire pages. Orphaned connectors and relationships are cleaned up automatically.

```typescript
// Delete a shape
await shape.delete();

// Delete a page (removes page file, rels, and all back-page references)
await doc.deletePage(page2);
```

#### 21. Lookup API
Find shapes and pages without iterating manually.

```typescript
// Find a shape by its numeric ID (searches nested groups too)
const target = await page.getShapeById("42");

// Find all shapes matching a predicate
const servers = await page.findShapes(s => s.text.startsWith("Server"));

// Look up a page by name (exact, case-sensitive)
const detailPage = doc.getPage("Architecture Diagram");
```

#### 22. Reading Shape Data Back
Retrieve custom properties, hyperlinks, and layer assignments that were previously written.

```typescript
// Custom properties (shape data)
const props = shape.getProperties();
console.log(props["IP"].value);    // "192.168.1.10"
console.log(props["Port"].type);   // VisioPropType.Number

// Hyperlinks
const links = shape.getHyperlinks();
// [ { address: "https://example.com", description: "Docs", newWindow: false } ]

// Layer indices
const indices = shape.getLayerIndices(); // e.g. [0, 2]
```

#### 23. Non-Rectangular Geometry
Use the `geometry` prop on `addShape()` to create common flowchart primitives without touching XML.

```typescript
// Ellipse / circle
await page.addShape({ text: "Start", x: 1, y: 5, width: 2, height: 2, geometry: 'ellipse' });

// Decision diamond
await page.addShape({ text: "Yes?", x: 4, y: 5, width: 2, height: 2, geometry: 'diamond' });

// Rounded rectangle (optional corner radius in inches)
await page.addShape({ text: "Process", x: 7, y: 5, width: 3, height: 2,
    geometry: 'rounded-rectangle', cornerRadius: 0.2 });

// Right-pointing triangle
await page.addShape({ text: "Extract", x: 1, y: 2, width: 2, height: 2, geometry: 'triangle' });

// Parallelogram (Data / IO shape)
await page.addShape({ text: "Input", x: 4, y: 2, width: 3, height: 1.5, geometry: 'parallelogram' });
```

Supported values: `'rectangle'` (default), `'ellipse'`, `'diamond'`, `'rounded-rectangle'`, `'triangle'`, `'parallelogram'`.

#### 24. Connector Styling
Control line appearance and routing algorithm on any connector.

```typescript
import { ArrowHeads } from 'ts-visio';

// Styled connector with crow's foot arrow and custom line
await shape1.connectTo(shape2, ArrowHeads.One, ArrowHeads.CrowsFoot, {
    lineColor: '#cc0000',   // Hex stroke color
    lineWeight: 1.5,        // Stroke width in points
    linePattern: 2,         // 1=solid, 2=dash, 3=dot, 4=dash-dot
    routing: 'curved',      // 'straight' | 'orthogonal' (default) | 'curved'
});

// Via page.connectShapes()
await page.connectShapes(a, b, undefined, undefined, { routing: 'straight' });
```

#### 25. Page Size & Orientation
Control the canvas dimensions using named paper sizes or raw inch values.

```typescript
import { PageSizes } from 'ts-visio';

// Use a named paper size (portrait by default)
page.setNamedSize('A4');                     // 8.268" × 11.693"
page.setNamedSize('Letter', 'landscape');   // 11" × 8.5"

// Set arbitrary dimensions in inches
page.setSize(13.33, 7.5);    // 13.33" × 7.5" widescreen

// Rotate the existing canvas without changing the paper size
page.setOrientation('landscape');  // swaps width and height if needed
page.setOrientation('portrait');

// Read current dimensions
console.log(page.pageWidth);    // e.g. 11
console.log(page.pageHeight);   // e.g. 8.5
console.log(page.orientation);  // 'landscape' | 'portrait'

// All size methods are chainable
page.setNamedSize('A3').setOrientation('landscape');
```

Available named sizes in `PageSizes`: `Letter`, `Legal`, `Tabloid`, `A3`, `A4`, `A5`.

---

#### 26. Document Metadata
Set and read document-level properties that appear in Visio's Document Properties dialog.

```typescript
// Write metadata (only supplied fields are changed)
doc.setMetadata({
    title:          'Network Topology',
    author:         'Alice',
    description:    'Data-centre interconnect diagram',
    keywords:       'network datacenter cloud',
    lastModifiedBy: 'CI pipeline',
    company:        'ACME Corp',
    manager:        'Bob',
    created:        new Date('2025-01-01T00:00:00Z'),
    modified:       new Date(),
});

// Read back all metadata fields
const meta = doc.getMetadata();
console.log(meta.title);    // 'Network Topology'
console.log(meta.author);   // 'Alice'
console.log(meta.company);  // 'ACME Corp'
console.log(meta.created);  // Date object
```

Fields map to OPC parts: `title`, `author`, `description`, `keywords`, `lastModifiedBy`,
`created`, `modified` → `docProps/core.xml`; `company`, `manager` → `docProps/app.xml`.

---

#### 27. Rich Text Formatting
Italic, underline, strikethrough, paragraph spacing, and text block margins are available both at shape-creation time and via `shape.setStyle()`.

```typescript
// At creation time
const shape = await page.addShape({
    text: 'Important',
    x: 2, y: 3, width: 3, height: 1,
    bold: true,
    italic: true,
    underline: true,
    strikethrough: false,
    // Paragraph spacing (in points)
    spaceBefore: 6,
    spaceAfter: 6,
    lineSpacing: 1.5,  // 1.5× line height
    // Text block margins (in inches)
    textMarginTop:    0.1,
    textMarginBottom: 0.1,
    textMarginLeft:   0.1,
    textMarginRight:  0.1,
});

// Post-creation via setStyle()
await shape.setStyle({
    italic: true,
    lineSpacing: 2.0,         // double spacing
    textMarginTop: 0.15,
    textMarginLeft: 0.05,
});
```

`lineSpacing` is a multiplier: `1.0` = single, `1.5` = 1.5×, `2.0` = double.
`spaceBefore` / `spaceAfter` are in **points**. Text margins are in **inches**.

---

#### 28. Named Connection Points
Define specific ports on shapes and connect to them precisely instead of relying on edge-intersection.

```typescript
import { StandardConnectionPoints } from 'ts-visio';

// 1. Add connection points at shape-creation time
const nodeA = await page.addShape({
    text: 'A', x: 2, y: 3, width: 2, height: 1,
    connectionPoints: StandardConnectionPoints.cardinal, // Top, Right, Bottom, Left
});

const nodeB = await page.addShape({
    text: 'B', x: 6, y: 3, width: 2, height: 1,
    connectionPoints: StandardConnectionPoints.cardinal,
});

// 2. Connect using named ports (Right of A → Left of B)
await page.connectShapes(nodeA, nodeB, undefined, undefined, undefined,
    { name: 'Right' },   // fromPort
    { name: 'Left' },    // toPort
);

// 3. Fluent Shape API
await nodeA.connectTo(nodeB, undefined, undefined, undefined,
    { name: 'Right' }, { name: 'Left' });

// 4. Add a point to an existing shape by index
const ix = nodeA.addConnectionPoint({
    name: 'Center',
    xFraction: 0.5, yFraction: 0.5,
    type: 'both',
});

// 5. Connect by zero-based index instead of name
await page.connectShapes(nodeA, nodeB, undefined, undefined, undefined,
    { index: 1 }, // Right (IX=1 in cardinal preset)
    { index: 3 }, // Left  (IX=3 in cardinal preset)
);

// 6. 'center' target (default behaviour) works alongside named ports
await page.connectShapes(nodeA, nodeB, undefined, undefined, undefined,
    'center', { name: 'Left' });
```

`StandardConnectionPoints.cardinal` — 4 points: `Top`, `Right`, `Bottom`, `Left`.
`StandardConnectionPoints.full` — 8 points: cardinal + `TopLeft`, `TopRight`, `BottomRight`, `BottomLeft`.
Unknown port names fall back gracefully to edge-intersection routing without throwing.

---

#### 29. StyleSheets (Document-Level Styles)
Define reusable named styles at the document level and apply them to shapes so they inherit line, fill, and text properties without repeating the same values on every shape.

```typescript
// 1. Create a document-level style
const headerStyle = doc.createStyle('Header', {
    fillColor:    '#4472C4',          // Blue fill
    lineColor:    '#2F5597',          // Dark-blue border
    lineWeight:   1.5,                // 1.5 pt stroke
    fontColor:    '#FFFFFF',          // White text
    fontSize:     14,                 // 14 pt
    bold:         true,
    fontFamily:   'Calibri',
    horzAlign:    'center',
    verticalAlign: 'middle',
});

const bodyStyle = doc.createStyle('Body', {
    fillColor:  '#DEEAF1',
    lineColor:  '#2F5597',
    fontSize:   11,
    horzAlign:  'left',
});

// 2. Apply at shape-creation time
const title = await page.addShape({
    text: 'System Architecture',
    x: 1, y: 8, width: 8, height: 1,
    styleId: headerStyle.id,          // sets LineStyle, FillStyle, TextStyle
});

// Fine-grained: apply only the fill from one style, line from another
const hybrid = await page.addShape({
    text: 'Hybrid',
    x: 1, y: 6, width: 3, height: 1,
    fillStyleId: headerStyle.id,
    lineStyleId: bodyStyle.id,
});

// 3. Apply (or change) style post-creation
const box = await page.addShape({ text: 'Server', x: 4, y: 4, width: 2, height: 1 });
box.applyStyle(bodyStyle.id);                // all three categories
box.applyStyle(headerStyle.id, 'fill');      // fill only — leaves line & text unchanged
box.applyStyle(headerStyle.id, 'text');      // text only

// 4. List all styles in the document
const styles = doc.getStyles();
// [ { id: 0, name: 'No Style' }, { id: 1, name: 'Normal' }, { id: 2, name: 'Header' }, … ]
```

`StyleProps` supports: `fillColor`, `lineColor`, `lineWeight` (pt), `linePattern`, `fontColor`, `fontSize` (pt), `bold`, `italic`, `underline`, `strikethrough`, `fontFamily`, `horzAlign`, `verticalAlign`, `spaceBefore`, `spaceAfter`, `lineSpacing`, `textMarginTop/Bottom/Left/Right` (in).
Local shape properties always override inherited stylesheet values.

#### 30. Reading Connectors Back
Enumerate connectors on a page — including those loaded from an existing `.vsdx` file — and inspect or delete them.

```typescript
import { VisioDocument } from 'ts-visio';

const doc  = await VisioDocument.load(buffer);   // or VisioDocument.create()
const page = doc.pages[0];

// Read all connector shapes
const connectors = page.getConnectors();

for (const conn of connectors) {
    console.log(`Connector ${conn.id}: ${conn.fromShapeId} → ${conn.toShapeId}`);
    console.log('  fromPort:', conn.fromPort);     // 'center' | { name } | { index }
    console.log('  toPort:',   conn.toPort);
    console.log('  style:',    conn.style);        // lineColor, lineWeight, linePattern, routing
    console.log('  arrows:',   conn.beginArrow, conn.endArrow);
}

// Delete a specific connector
await connectors[0].delete();

// After delete the connector is gone but the shapes remain
console.log(page.getConnectors().length);  // 0
console.log(page.getShapes().length);      // shapes still exist
```

#### 31. Color Palette
Register colors in the document's indexed color table (`<Colors>` in `document.xml`). Colors are identified by an integer IX that you can reference anywhere a hex color is accepted.

```typescript
// Register colors — returns the integer index (IX)
const blueIx  = doc.addColor('#4472C4');  // → 2 (first user color)
const redIx   = doc.addColor('#FF0000');  // → 3
const greenIx = doc.addColor('#00B050');  // → 4

// De-duplication: adding the same color returns the existing index
doc.addColor('#4472C4');  // → 2 (no duplicate created)

// Shorthand, missing #, and case variations are all normalised
doc.addColor('#FFF');        // same as #FFFFFF → returns 1 (built-in white)
doc.addColor('4472c4');      // same as #4472C4 → returns 2

// Read the full palette (sorted by index)
const palette = doc.getColors();
// [
//   { index: 0, rgb: '#000000' },  // built-in black
//   { index: 1, rgb: '#FFFFFF' },  // built-in white
//   { index: 2, rgb: '#4472C4' },
//   { index: 3, rgb: '#FF0000' },
//   { index: 4, rgb: '#00B050' },
// ]

// Look up an index by hex value
doc.getColorIndex('#4472C4');  // → 2
doc.getColorIndex('#123456');  // → undefined (not registered)
```

Built-in colors: IX 0 = `#000000` (black), IX 1 = `#FFFFFF` (white). User colors start at IX 2 and increment sequentially.

---

#### 32. Group Traversal (`shape.getChildren()`)
Access nested child shapes of a group without touching XML. Only **direct** children are returned — call `getChildren()` recursively to walk a deeper tree.

```typescript
// 1. Create a group with children
const group = await page.addShape({
    text: 'Container', x: 5, y: 5, width: 6, height: 6, type: 'Group'
});
const childA = await page.addShape({ text: 'A', x: 1, y: 1, width: 2, height: 1 }, group.id);
const childB = await page.addShape({ text: 'B', x: 1, y: 3, width: 2, height: 1 }, group.id);

// 2. Check if a shape is a group
console.log(group.isGroup);   // true
console.log(childA.isGroup);  // false
console.log(group.type);      // 'Group'

// 3. Get direct children
const children = group.getChildren();
// → [Shape('A'), Shape('B')]

for (const child of children) {
    console.log(child.text, child.id);
}

// 4. Recursively walk a nested tree
function walk(shape, depth = 0) {
    console.log(' '.repeat(depth * 2) + shape.text);
    for (const child of shape.getChildren()) {
        walk(child, depth + 1);
    }
}
walk(group);

// 5. Works on shapes loaded from an existing .vsdx file
const doc2   = await VisioDocument.load('existing.vsdx');
const shapes = doc2.pages[0].getShapes();
const groups = shapes.filter(s => s.isGroup);
for (const g of groups) {
    console.log(`Group "${g.text}" has ${g.getChildren().length} children`);
}
```

`getChildren()` returns `[]` for non-group shapes. Children are full `Shape` instances — all existing methods (`setStyle()`, `getProperties()`, `delete()`, etc.) work on them.

#### 33. Drawing Scale
Map page coordinates to real-world units. One `pageScale` `pageUnit` on the paper equals `drawingScale` `drawingUnit` in reality. Visio uses this to display rulers, grids, and shape dimensions in real-world terms.

```typescript
import type { LengthUnit, DrawingScaleInfo } from 'ts-visio';

// 1 inch on paper = 10 feet in the real world (architectural)
page.setDrawingScale(1, 'in', 10, 'ft');

// 1:100 metric (1 cm on paper = 100 cm = 1 m in reality)
page.setDrawingScale(1, 'cm', 100, 'cm');

// Engineering: 1 inch = 20 feet
page.setDrawingScale(1, 'in', 20, 'ft');

// Read the current scale back
const scale: DrawingScaleInfo | null = page.getDrawingScale();
if (scale) {
    console.log(`${scale.pageScale} ${scale.pageUnit} = ${scale.drawingScale} ${scale.drawingUnit}`);
    // → "1 in = 10 ft"
}

// Returns null when no custom scale is set
const freshPage = doc.pages[0];
console.log(freshPage.getDrawingScale()); // null

// Reset to 1:1 (no scale)
page.clearDrawingScale();
console.log(page.getDrawingScale()); // null

// All methods are chainable
page.setDrawingScale(1, 'mm', 1000, 'mm')
    .setNamedSize('A1');  // combine with page size changes
```

Supported `LengthUnit` values:
- Imperial: `'in'` (inches), `'ft'` (feet), `'yd'` (yards), `'mi'` (miles)
- Metric: `'mm'`, `'cm'`, `'m'`, `'km'`

The drawing scale does not affect the internal coordinate system — shape positions and sizes are still specified in page-inches. The scale is purely a display/annotation mapping applied by Visio's ruler and grid.

---

## Examples

Check out the [examples](./examples) directory for complete scripts.

- **[Simple Schema](./examples/simple-schema.ts)**: Generates a database schema ERD with tables, styling, and Crow's Foot connectors.
- **[Network Topology](./examples/network-diagram.ts)**: Demonstrates the **Fluent Shape Data API** to build a network map with hidden metadata, typed properties, and connections.
- **[Containers Demo](./examples/containers_demo.ts)**: Shows how to create Classic Visio containers and place shapes within them.
- **[Lists Demo](./examples/lists_demo.ts)**: Demonstrates Vertical and Horizontal List stacks.
- **[Hyperlinks Demo](./examples/hyperlinks_demo.ts)**: Demonstrates Internal and External navigation.
- **[Layers Demo](./examples/layers_demo.ts)**: Shows how to create layers and toggle visibility.
- **[Image Embedding Demo](./examples/images_demo.ts)**: Demonstrates how to embed PNG or JPEG images into a diagram.
- **[Swimlane Demo](./examples/swimlane_demo.ts)**: Demonstrates creating Cross-Functional Flowcharts with Pools and Lanes.

## Development

This project uses **Vitest** for testing.

```bash
npm install
npm test
```
