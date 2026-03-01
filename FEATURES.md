# ts-visio Feature Gap Analysis

## What the Library Currently Supports

- **Document**: Create, load (path/Buffer/ArrayBuffer/Uint8Array), save
- **Pages**: Add foreground/background pages, set background page, enumerate pages
- **Shapes**: Standard rectangles, foreign/image shapes, containers, lists, swimlane pools/lanes, tables, groups
- **Connectors**: Connect shapes with begin/end arrow styles
- **Shape positioning**: `placeRightOf`, `placeBelow`, absolute position update
- **Shape data**: Custom properties (typed: string, number, boolean, date)
- **Hyperlinks**: External URLs, internal page links
- **Layers**: Add layers, assign shapes, update visibility/lock/print properties
- **Z-order**: Send to front / send to back
- **Container operations**: `addMember`, `addListItem`, `resizeToFit`
- **Style**: `fillColor`, `fontColor`, `bold`, `lineColor`, `linePattern`, `lineWeight`
- **Reading**: Read shapes from existing VSDX files
- **Validation**: Structural validation of VSDX package (required files, rels integrity, shape IDs, Connect references, master references)
- **High-level builders**: `SchemaDiagram` (ER diagrams with 1:1 / 1:N relations)

---

## Feature Gaps

### 1. Deletion / Mutation

- ~~`shape.delete()`~~ — ✅ Implemented (removes shape, orphaned Connects, and container Relationships)
- ~~`doc.deletePage(page)`~~ — ✅ Implemented (removes page file, rels, pages.xml entry, Content Types override, BackPage refs)
- `layer.delete()`
- `connector.delete()`

---

### 2. Text Styling

`ShapeStyle` only covers `fillColor`, `fontColor`, and `bold`. Missing:

- ~~**Font family**~~ — ✅ Implemented (`fontFamily` prop, uses `FONT("name")` formula)
- ~~**Font size**~~ — ✅ Implemented (`fontSize` in points, stored as inches internally)
- ~~**Text alignment**~~ — ✅ Implemented (`horzAlign`: left/center/right/justify via Paragraph section; `verticalAlign`: top/middle/bottom as top-level shape cell)
- **Italic / Underline / Strikethrough** — `createCharacterSection` ignores the other Style bits (2=Italic, 4=Underline, 8=Strikethrough)
- **Text margins** — no padding/margin control (TxtWidth, TxtHeight, TxtPinX, TxtPinY)
- **Paragraph spacing** — no line spacing, space-before/after control (Paragraph section)

---

### 3. Shape Transformations

- ~~**Rotation**~~ — ✅ Implemented (`shape.rotate(degrees)` / `shape.angle` getter; stored as radians in `Angle` cell)
- ~~**Flip**~~ — ✅ Implemented (`shape.flipX(enabled?)` / `shape.flipY(enabled?)` via `FlipX`/`FlipY` cells)
- ~~**Resize**~~ — ✅ Implemented (`shape.resize(width, height)`; updates Width, Height, LocPinX, LocPinY, and Geometry cache)

---

### 4. Non-Rectangular Geometry

- ~~**Ellipse / circle**~~ — ✅ Implemented (Visio `Ellipse` row; `geometry: 'ellipse'`)
- ~~**Rounded rectangle**~~ — ✅ Implemented (`EllipticalArcTo` corners; `geometry: 'rounded-rectangle'`, optional `cornerRadius` in inches)
- ~~**Diamond / rhombus**~~ — ✅ Implemented (4-vertex polygon starting at top; `geometry: 'diamond'`)
- ~~**Flowchart primitives**~~ — ✅ Implemented: `'triangle'` (right-pointing), `'parallelogram'` (Data/IO shape, 20% skew)

---

### 5. Reading Data Back from Existing Shapes

`ShapeReader` parses `Sections` internally but they are not surfaced through the `Shape` public API:

- ~~`shape.getProperties()`~~ — ✅ Implemented (typed value coercion: String, Number, Boolean, Date)
- ~~`shape.getHyperlinks()`~~ — ✅ Implemented (address, subAddress, description, newWindow)
- ~~`shape.getLayerIndices()`~~ — ✅ Implemented (returns `number[]`)
- `page.getConnectors()` — read existing `<Connect>` elements from a loaded file
- `page.getLayers()` — read existing layers from a loaded file
- Sub-shapes of groups are parsed but not accessible (top-level only via `getShapes()`)

---

### 6. Page Operations

- `doc.getPage(name)` / `doc.findPage(name)` — no lookup by name or ID
- `page.rename(name)`
- `doc.deletePage(page)`
- Page reordering
- Page duplication
- **Page size / orientation** — hardcoded 8.5×11 in `PageManager`; no public API to change it
- **Drawing scale** — scale settings are fixed in the page template

---

### 7. Connector Styling

- ~~**Line color, weight, and pattern**~~ — ✅ Implemented (`ConnectorStyle.lineColor`, `lineWeight` in pt, `linePattern`)
- ~~**Routing style**~~ — ✅ Implemented (`ConnectorStyle.routing`: `'straight'` / `'orthogonal'` / `'curved'`; maps to Visio `ShapeRouteStyle` 2/1/16)
- Named connection points — connecting to a specific port on a shape rather than shape-to-shape center

---

### 8. Document-Level Properties

- **Document metadata** — title, author, description, keywords (`docProps/core.xml`)
- **StyleSheet** — document-level line/fill/text styles that shapes can inherit from
- **Color palette** — document-level color table

---

### 9. Masters / Stencils

`MasterManager` exists internally but has no public API:

- No public method to create or import master shapes
- No way to load a `.vssx` stencil file and use its masters
- The `masterId` mechanism works for referencing existing masters, but creating them from scratch is not supported

---

### 10. Missing Public Exports

- ~~`SchemaDiagram`~~ — ✅ Exported
- ~~`VisioValidator`~~ — ✅ Exported
- ~~`Layer`~~ — ✅ Exported

---

### 11. API Ergonomics

- ~~`page.findShapes(predicate)`~~ — ✅ Implemented (searches all shapes including nested group children)
- ~~`page.getShapeById(id)`~~ — ✅ Implemented (recursive search through group tree)
- ~~`doc.getPage(name)`~~ — ✅ Implemented (exact name match, case-sensitive)
- `shape.getChildren()` — access sub-shapes of groups/containers
- `Shape.setStyle` only accepts `ShapeStyle` which does not cover line style; no way to change border color post-creation via the public API

---

## Priority Ranking

| Priority | Gap |
|----------|-----|
| ✅ Done | Font size & family, text alignment |
| ✅ Done | `deleteShape`, `deletePage` |
| ✅ Done | `page.getShapeById`, `page.findShapes`, `doc.getPage(name)` |
| ✅ Done | Read properties/hyperlinks/layer indices back from existing shapes |
| ✅ Done | Rotation, flip, and resize via `Shape` API |
| ✅ Done | Non-rectangular geometry (ellipse, diamond, rounded-rect, triangle, parallelogram) |
| ✅ Done | Connector routing style and line styling |
| ✅ Done | Missing exports (`Layer`, `SchemaDiagram`, `VisioValidator`) |
| 🟡 Medium | Page size / orientation API |
| 🟢 Low | Document metadata |
| 🟢 Low | Masters / stencils public API |
| 🟢 Low | Rich text / paragraph formatting |
