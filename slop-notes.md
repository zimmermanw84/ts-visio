ts-visio Development Plan
Phase 1: Repo Setup & Core Types
• Goal: Initialize a TypeScript library with strict typing and necessary dependencies.

• Prompt to AI: "Initialize a new Node.js project for an npm library called `ts-visio`. Setup TypeScript with a `strict` config. Install `jszip` for file handling and `fast-xml-parser` for XML manipulation. Create a folder structure with `src/core`, `src/xml`, and `src/utils`."

Phase 2: The OPC Container (File I/O)
• Goal: Handle the `.vsdx` file format (Open Packaging Conventions), which is just a ZIP of XMLs.

• Prompt to AI: "Create a `VisioPackage` class in `src/core/VisioPackage.ts`. It should have a `load(buffer: Buffer)` method that uses JSZip to unzip the content. Store the file map in memory. Add a helper method `getFileText(path: string)` to read specific XML files from the internal zip structure."

Phase 3: The Reader (Parsing XML)
• Goal: Locate pages and shapes within the unzipped structure.

• Prompt to AI: "Implement a `PageManager` that reads `visio/pages/pages.xml`. It should return an array of available pages with their IDs and names. Then, create a `ShapeReader` that can parse a specific `visio/pages/pageX.xml` file and return a list of `<Shape>` elements using `fast-xml-parser`."

Phase 4: The Writer (Modification)
• Goal: Update text or properties and re-zip the file.

• Prompt to AI: "Create a method `updateShapeText(pageId: string, shapeId: string, newText: string)`. It needs to find the correct `pageX.xml`, locate the `<Text>` node for that shape, update the content, and update the internal JSZip object with the modified XML string."

• Prompt to AI: "Implement a `save()` method on the `VisioPackage` class. It should generate a new binary buffer from the JSZip instance so the user can write it to disk."

Phase 5: High-Level Abstractions (API Design)
• Goal: Hide the XML complexity from the end user.

• Prompt to AI: "Refactor the code to expose a clean public API. Create a `VisioDocument` class that wraps `VisioPackage`. It should have methods like `doc.getPages()`, `page.getShapes()`, and `shape.setText()`. Ensure no raw XML types leak to the public API."


-------------


Here are the specific, technical AI prompts you can use to build these features. I have organized them by the implementation order recommended in your analysis.


### Phase 1: Connectors (The "Critical" Path)

*Context:* You need to teach the AI about Visio's specific way of handling 1D shapes and the separate `<Connects>` collection.

**Prompt 1: Understanding the XML Structure**

> "I am building a raw Visio XML generator in TypeScript. I need to understand how 'Connectors' work in the `.vsdx` schema.
> 1. How is a 1D connector shape defined differently from a 2D box in the `<Shapes>` list?
> 2. Explain the role of the `<Connects>` collection in `page1.xml`.
> 3. Provide a sample XML snippet showing two box shapes linked by a dynamic connector."
>
>

**Prompt 2: Implementing `addConnector**`

> "Based on the standard Open XML Visio schema, write a TypeScript method `addConnector(pageId: string, fromShapeId: number, toShapeId: number)` for my `VisioPackage` class.
> The function needs to:
> 1. Create a new Shape element with 1D geometry (LineTo/MoveTo) in the `page.xml`.
> 2. Add two entries to the `<Connects>` collection: one linking the connector's 'BeginX' to the `fromShape` and one linking 'EndX' to the `toShape`.
> 3. Use `fast-xml-parser` logic or string template insertion."
>
>

---

### Phase 2: Styling (Fill & Text)

*Context:* Visio handles styles in "Sections" (Fill, Line, Character) within the ShapeSheet.

**Prompt 3: Background Colors (Fill)**

> "I need to add a `fillColor` property to my shape creation logic. In the Visio ShapeSheet XML, which `<Section>` and `<Cell>` handles the background color?
> Please write a helper function `createFillSection(hexColor: string)` that converts a standard Hex code (like #CCCCCC) into the Visio XML format required for the `FillForegnd` cell."

**Prompt 4: Font Styling (Bold/Color)**

> "I need to support bold text and font colors. This requires modifying the `<Character>` section of a shape.
> Write a TypeScript function that generates the XML for a `<Character>` section. It should accept an options object: `{ bold: boolean, color: string }`.
> *Note:* Explain how Visio maps boolean values for 'Bold' (is it '1' or 'TRUE'?) and how it handles font colors."

---

### Phase 3: Line Ends (Crow's Foot Notation)

*Context:* This is crucial for ER Diagrams. Visio uses integer IDs for arrowhead types.

**Prompt 5: Arrowheads**

> "I am generating an Entity Relationship diagram. I need to modify my connector shapes to show 'Crow's Foot' notation.
> 1. Which cells in the `<Line>` section of the ShapeSheet control the start and end arrowheads?
> 2. What are the specific integer IDs (values) for the 'Crow's Foot' symbol and the 'One' (single dash) symbol in Visio?
> 3. Update my `addConnector` function to accept `beginArrow` and `endArrow` parameters."
>
>

---

### Phase 4: Compound Shapes (Tables)

*Context:* Handling "Groups" is complex. A simpler "Stack" approach is often better for V1.

**Prompt 6: The "Stacking" Logic**

> "I need to generate a database table shape that consists of a Header Rectangle (grey) stacked perfectly on top of a Body Rectangle (white).
> Instead of using complex Visio Groups, write a high-level function `addTable(x, y, title, columns[])`.
> It should:
> 1. Calculate the height of the header and body based on the text.
> 2. Call `addShape` for the Header at `(x, y)`.
> 3. Call `addShape` for the Body at `(x, y - headerHeight)` (since Visio coordinates start from bottom-left).
> 4. Return the ID of the main container shape."
>
>