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
> 5. Update README.md with the new features.
> 6. Update tests to reflect the new features.


----

Here is an itemized markdown list of prompts designed to guide an AI coding assistant (like Copilot, Gemini, or ChatGPT) through these refactors.

You can copy and paste these directly into your chat window.

# AI Prompts for ts-visio Refactoring

## 1. Object-Oriented Return Types

*Context: We need to stop returning strings (IDs) and start returning rich objects.*

* [ ] **Prompt:** "Refactor the `Page.addShape` and `Page.addTable` methods. Currently, they return a `Promise<string>` (the Shape ID). Change them to return a `Promise<Shape>` object. The `Shape` object should encapsulate the ID and provide access to its own properties (like `x`, `y`, `width`, `height`). Update the `Shape` interface to support this."
* [ ] **Prompt:** "Scan the codebase for all instances where `addShape` is called. Refactor the calling code to handle the new object return type instead of the string ID string."

## 2. Fluent Chaining

*Context: We want to enable syntax like `shape.connectTo(other)`. This requires the `Shape` object to have methods.*

* [ ] **Prompt:** "I want to implement a fluent API for my `Shape` class. Add a method `connectTo(targetShape: Shape, beginArrow?: number, endArrow?: number)` to the `Shape` class. This method should internally call the page's connection logic using `this.id` and `targetShape.id`. It should return `this` (or a Promise of `this`) to allow for further chaining."
* [ ] **Prompt:** "Add a method `setStyle(styleProps: ShapeStyle)` to the `Shape` class that allows updating fill and text properties. Ensure it returns `this` to support chains like `myShape.connectTo(other).setStyle({...})`."

## 3. Shape Grouping (Complex XML)

*Context: Visio Groups act like mini-pages. This is the hardest implementation detail.*

* [ ] **Prompt:** "I need to implement native Visio 'Group' shapes. Explain the XML structure for a `<Shape Type='Group'>` in a `.vsdx` file. How does it differ from a standard 2D shape? specifically, how do I nest other shapes inside it?"
* [ ] **Prompt:** "Refactor `addTable` to use a Group shape.
1. Create a parent Group shape to hold the table.
2. Inside the Group, add the 'Header' rectangle and the 'Body' rectangle.
3. **Crucial:** Adjust the coordinates of the Header and Body to be relative to the Group's parent coordinates, not the Page coordinates."



## 4. Automatic Layout / Relative Positioning

*Context: We need simple math to calculate bounding boxes.*

* [ ] **Prompt:** "Implement a `placeRightOf(target: Shape, options: { gap: number })` method on the `Shape` class.
1. Retrieve the `x`, `y`, `width` of the `target` shape.
2. Calculate the new `x` for the current shape (`target.x + target.width + gap`).
3. Keep the same `y` coordinate (top-aligned).
4. Update the current shape's internal state and XML to reflect this new position."
5. Add tests for this method.
6. Update README.md with the new feature.


* [ ] **Prompt:** "Add a `placeBelow(target: Shape, options: { gap: number })` method that performs similar logic for vertical stacking."

## 5. Typed Schema Builder

*Context: This is a 'Facade' pattern that simplifies the generic API for a specific use case.*

* [ ] **Prompt:** "Create a new class `SchemaDiagram` that acts as a domain-specific wrapper around `VisioPage`.
1. It should have a method `addTable(tableName: string, columns: string[])`.
2. It should have a method `addRelation(fromTable: Shape, toTable: Shape, type: '1:1' | '1:N')`.
3. Map '1:1' to specific Visio arrow IDs (e.g., standard arrowheads) and '1:N' to Crow's Foot arrowheads."



---

### Recommended Execution Order

I suggest running the prompts in this exact order: **1 → 3 → 2 → 4 → 5**.

**Why?**

* You need the **Objects (1)** to exist before you can **Group (3)** them.
* You need the **Group (3)** logic working before you add **Chaining (2)** methods to it.
* **Layout (4)** relies on the unified coordinates of the Group, so it must come after.


---

This is a significant architectural shift. Moving from "drawing lines" to "instantiating masters" is the correct way to handle Visio files, as it dramatically reduces file size and enables the use of complex icons (like network gear or AWS symbols).

Here is the itemized prompt plan to execute this transition cleanly.

### Phase 1: Parsing Existing Masters (`masters.xml`)

*Goal: Before we can use a Master, we must identify what Masters already exist in the template file.*

**Prompt 1: The Master Manager**

> "Create a `MasterManager` class in `src/core/MasterManager.ts`.
> 1. It should accept the unzipped `VisioPackage`.
> 2. Implement `load()`: Read `visio/masters/masters.xml` to find all registered masters.
> 3. Return an array of objects: `{ id: string, name: string, type: string, xmlPath: string }`.
> 4. Use `fast-xml-parser` to handle the XML.
> 5. **Test Requirement:** Create `tests/MasterManager.test.ts` that mocks a `masters.xml` string and asserts that the parser correctly extracts the ID and Name (e.g., 'Rectangle')."
>
>

**Prompt 2: PR Generation (Phase 1)**

> "I have implemented the `MasterManager`. Please generate a Markdown PR description.
> * **Title:** feat: Core MasterManager for parsing stencil data
> * **Description:** Explain that this reads the `masters.xml` index.
> * **Changes:** List the new class and the interface changes.
> * **Verification:** Mention the unit tests."
>
>

---

### Phase 2: Instantiating a Master (The "Drop" Logic)

*Goal: Modify `addShape` to link to a Master ID instead of drawing raw geometry.*

**Prompt 3: Refactoring `addShape` Strategy**

> "I need to refactor `Page.addShape` to support 'Dropping' a master.
> 1. Update the signature to accept an optional `masterId: string`.
> 2. **Logic Change:**
> * **If `masterId` is present:** Create a `<Shape Master='ID'>` tag. **Do not** generate `<Geom>` or `<Section>` tags (geometry is inherited). Only set `PinX`, `PinY`, `Width`, and `Height`.
> * **If `masterId` is missing:** Keep the existing logic that draws the rectangle manually (Regression support).
>
>
> 3. **Test Requirement:** Write a test case `should drop master shape` that verifies the output XML contains the `Master` attribute and *excludes* the `<Geom>` section."
>
>

**Prompt 4: Regression Testing**

> "Write a regression test suite for `addShape`.
> 1. **Test A (Legacy):** Call `addShape` without a master ID. Assert that `<Geom>` and `<MoveTo>` tags still exist in the output.
> 2. **Test B (New):** Call `addShape` with a master ID. Assert `Master` attribute exists and `<Geom>` tags are absent.
> 3. Ensure both shapes can exist on the same page."
>
>

**Prompt 5: PR Generation (Phase 2)**

> "Generate a PR description for the `addShape` refactor.
> * **Title:** feat: Enable Shape Instantiation via Master ID
> * **Description:** Detail how this reduces file size by inheriting geometry.
> * **Breaking Changes:** None (backward compatible).
> * **Documentation:** Provide a code snippet for the `README.md` showing how to find a master by name and drop it."
>
>

---

### Phase 3: Handling Relationships (The `.rels` File)

*Goal: Visio is strict. You cannot just use a Master ID in `page1.xml`. You must also declare the dependency in `page1.xml.rels`.*

**Prompt 6: The Rels Manager**

> "Visio requires a relationship link between the Page and the Master.
> 1. Create `src/core/RelsManager.ts` to handle `.rels` files (OPC relationships).
> 2. When `addShape(masterId)` is called:
> * Check if `page1.xml.rels` already has a relationship to that Master's file (e.g., `../masters/master1.xml`).
> * If not, generate a new `Relationship` ID (e.g., `rId5`) and append it to the `.rels` file.
>
>
> 3. **Critical:** The `addShape` function must now wait for this Relationship ID and using *that* might be required in some contexts, though typically the Master ID `(e.g. '2')` is used in the Shape attribute. Verify this constraint: Does the Shape tag use the Master's *internal ID* or the *Relationship ID*?"
>
>

**Prompt 7: Integration Test (The Full Loop)**

> "Write a full integration test:
> 1. Load a template `.vsdx` that has a 'Router' master inside it.
> 2. Drop the 'Router' master onto Page 1.
> 3. Save the file.
> 4. **Verification:** Inspect the ZIP contents. Ensure `page1.xml` has a shape with `Master='X'` and `page1.xml.rels` is valid."
>
>

**Prompt 8: Final Documentation & PR**

> "Generate the final PR description for the `.rels` management.
> * **Title:** fix: Manage OPC Relationships for Masters
> * **Context:** Explain that without updating `.rels`, Visio considers the file corrupt.
> * **Updates:** Update the root `README.md` to include a complete example: 'Loading a Template & Dropping Stencils'."
>
>

---

### Recommended Execution Order

1. **Phase 1 (Reading)** allows you to see what you are working with.
2. **Phase 3 (Rels)** is actually required *before* Phase 2 will produce a valid file, but you can code Phase 2's logic first. I recommend doing **1 -> 3 -> 2** if you want the tests to pass immediately, or **1 -> 2 -> 3** if you want to implement the logic before the plumbing.


---

Here is the itemized prompt plan to implement **Multi-Page Support**. This logic is critical because if you miss registering a page in any of the three required locations (`[Content_Types]`, `pages.xml`, or `.rels`), Visio will declare the file corrupt.

### Phase 1: The Page Index (Reading `pages.xml`)

*Goal: Stop assuming `page1.xml` exists. We need to dynamically load the list of pages from the index.*

**Prompt 1: PageManager Implementation**

> "Create a `PageManager` class in `src/core/PageManager.ts`.
> 1. In the constructor, accept the `VisioPackage` (JSZip instance).
> 2. Implement `load()`: Parse `visio/pages/pages.xml` to build a registry of existing pages.
> 3. Store them as an array of objects: `{ id: number, name: string, relId: string, xmlPath: string }`.
> 4. **Important:** Determine the `xmlPath` by looking up the `relId` in `visio/pages/_rels/pages.xml.rels`.
> 5. **Test Requirement:** Create `tests/PageManager.test.ts` that loads a sample 3-page Visio file and correctly identifies the IDs and Names of all pages."
>
>

**Prompt 2: PR Generation (Phase 1)**

> "Generate a Markdown PR description for the `PageManager` read logic.
> * **Title:** feat: Parse dynamic page list from pages.xml
> * **Description:** Explain that we are removing the hardcoded `page1.xml` dependency.
> * **Changes:** New `PageManager` class and `rels` lookup logic.
> * **Verification:** Unit tests for parsing multi-page indexes."
>
>

---

### Phase 2: Page Creation (The "Corrupt File" Avoidance Protocol)

*Goal: Create the physical XML files and register them in the global Content Types. If this step fails, the file breaks.*

**Prompt 3: Generating the Page XML**

> "I need to implement `PageManager.createPage(name: string)`.
> 1. Calculate the next available Page ID (e.g., if ID 1 exists, use ID 2).
> 2. Generate a new empty page XML string (standard `<PageContents>` header).
> 3. Write this string to the zip at `visio/pages/page{ID}.xml`.
> 4. **Critical:** You must also register this new XML file in the root `[Content_Types].xml` file with `Override PartName='/visio/pages/page{ID}.xml' ContentType='application/vnd.ms-visio.page+xml'`.
> 5. **Test Requirement:** Write a test that adds a page and verifies the file exists in the ZIP and the entry exists in `[Content_Types].xml`."
>
>

**Prompt 4: Updating the Page Index & Rels**

> "Now link the new page file to the document structure.
> 1. In `visio/pages/_rels/pages.xml.rels`, add a new Relationship pointing to `page{ID}.xml`. Generate a new `rId`.
> 2. In `visio/pages/pages.xml`, append a new `<Page>` element.
> * Set `ID` to the new ID.
> * Set `Name` to the function argument.
> * Set `r:id` to the new relationship ID you just generated.
>
>
> 3. **Validation:** Ensure the `PageSheet` (page properties) inside the new XML inherits default dimensions."
>
>

**Prompt 5: PR Generation (Phase 2)**

> "Generate a PR description for the Page Creation logic.
> * **Title:** feat: Implement physical page creation and registration
> * **Context:** Creating a page requires updates to 4 different XML files.
> * **Risks:** Mention that `[Content_Types].xml` validation is included to prevent file corruption.
> * **Tests:** Unit tests verifying ZIP structure integrity."
>
>

---

### Phase 3: Public API & Regression Testing

*Goal: expose `doc.addPage()` and ensure we didn't break the single-page use case.*

**Prompt 6: High-Level API Refactor**

> "Refactor `VisioDocument` to use `PageManager`.
> 1. Expose `doc.addPage(name: string): Page`. This should return the object-oriented `Page` wrapper we built previously.
> 2. Expose `doc.getPages(): Page[]`.
> 3. **Regression Requirement:** Ensure `doc.pages[0]` still works for the default page so existing code examples don't break.
> 4. **Test Requirement:** Write a `MultiPage.test.ts` that creates a document, adds 'Architecture', adds 'Database', and draws a shape on both. Save and Validate."
>
>

**Prompt 7: Final Documentation & PR**

> "Generate the final PR description.
> * **Title:** feat: Public API for Multi-Page Documents
> * **Changes:** `VisioDocument` now exposes `addPage()`.
> * **Example:** Provide a markdown code snippet showing how to create a 2-page document.
> * **Checklist:** Confirm existing single-page tests pass (Regression)."
>
>

---

Here is the itemized prompt plan to implement **Shape Data (Custom Properties)**. This feature turns a drawing into a database by allowing shapes to hold invisible or visible metadata.

### Phase 1: The Property Section (Schema & Definition)

*Goal: Create the structure to hold data. In Visio, you cannot just "set a value"; you must first define the property row (Label, Type, Format) in the ShapeSheet.*

**Prompt 1: Understanding Property XML**

> "I need to implement Custom Properties (Shape Data).
> 1. Explain the XML structure of the `<Section N='Property'>` in the ShapeSheet.
> 2. specifically, detail the roles of these Cells:
> * `Prop.Name` (The row reference)
> * `Label` (The display name, e.g., 'IP Address')
> * `Value` (The data)
> * `Type` (String vs Int vs Date)
> * `Invisible` (Hidden metadata)
>
>
> 3. Provide a sample XML snippet of a Shape that has a visible 'Cost' property and a hidden 'ID' property."
>
>

**Prompt 2: Implementing `addPropertyDefinition**`

> "Implement a method `addPropertyDefinition` on the `Shape` class.
> 1. Arguments: `name: string` (internal key), `label: string` (display name), `type: VisioPropType`.
> 2. **Logic:**
> * Check if `<Section N='Property'>` exists. If not, create it.
> * Create a new `<Row>` with `N='Prop.{name}'`.
> * Add `<Cell N='Label' V='{label}'/>`.
> * Add `<Cell N='Type' V='{type}'/>`.
>
>
> 3. **Return:** The newly created Row object or ID."
>
>

**Prompt 3: PR Generation (Phase 1)**

> "Generate a Markdown PR description.
> * **Title:** feat: Core Logic for Shape Data Definitions
> * **Description:** Adds support for `<Section N='Property'>`.
> * **Technical Detail:** Explain how we map TypeScript types to Visio Property Types (0=String, 2=Number, etc.).
> * **Test Plan:** Verify XML output contains correct `<Row>` structure."
>
>

### Phase 2: Visibility & Data Types (The Logic)

*Goal: Handle the specific use case of "Hidden" data and ensuring Dates/Numbers are formatted correctly so Excel exports work.*

**Prompt 4: Handling Visibility (Hidden Data)**

> "Refactor `addPropertyDefinition` to accept an options object: `{ invisible: boolean }`.
> 1. If `invisible` is true, add the `<Cell N='Invisible' V='1'/>` to the property Row.
> 2. **Test Requirement:** Create a test case 'Hidden Metadata' where we add an 'EmployeeID' field with `invisible: true`. Assert that the XML has `Invisible` set to 1."
>
>

**Prompt 5: Setting Values**

> "Implement `setPropertyValue(key: string, value: any)`.
> 1. Locate the `<Row N='Prop.{key}'>`.
> 2. Update the `<Cell N='Value'>`.
> 3. **Critical:** Visio stores values as formulae or strings.
> * If string: `V='StringContent'`.
> * If number: `V='100'`.
> * If date: Visio uses a specific floating point date format or ISO string. How should we handle JS Date objects?
>
>
> 4. Implement the Date serialization logic for Visio."
>
>

**Prompt 6: PR Generation (Phase 2)**

> "Generate a PR description for Data Types & Visibility.
> * **Title:** feat: Support for Hidden Properties and Typed Values
> * **Context:** Necessary for 'Visual Database' use cases where metadata shouldn't clutter the diagram.
> * **Changes:** Added Invisible cell support and Date serialization helper."
> Double check to ensure we're not missing any test cases
> Ensure that the PR is copy pastable for github
>

---

### Phase 3: Developer Experience (Fluent API)

*Goal: Make it one line of code to add metadata.*

**Prompt 7: High-Level Abstraction**

> "Create a fluent API for this feature.
> 1. Create a `ShapeData` interface: `{ label?: string, value: string | number, hidden?: boolean }`.
> 2. Add a method `shape.addData(key: string, data: ShapeData)`.
> 3. **Example Usage:**
> ```typescript
> shape.addData('ip', { label: 'IP Address', value: '192.168.1.1' })
>      .addData('id', { value: 12345, hidden: true });
>
> ```
>
>
> 4. Ensure this calls the low-level definition and value setters we built in previous phases."
>
>

**Prompt 8: Integration Test & PR**

> "Write an integration test:
> 1. Create a shape.
> 2. Add 3 data fields (String, Number, Hidden).
> 3. Save the file.
> 4. **Verification:** Start a mock XML parser to read the saved file and verify that the `Property` section contains exactly 3 rows with the correct attributes.
> 5. Double check to ensure we're not missing any test cases
> 6. Generate the final PR description in github markdown format for easy copy pasting
>
>

Here is the itemized prompt plan to implement **Image Support**. This is complex because it bridges two worlds: binary file handling (saving the PNG/JPG) and XML referencing (displaying it).

### Phase 1: Binary Storage & Content Types (The Plumbing)

*Goal: Successfully save an image file into the ZIP container and ensure Visio recognizes the file format.*

**Prompt 1: The Media Manager**

> "Create a `MediaManager` class in `src/core/MediaManager.ts`.
> 1. Implement `addMedia(filename: string, data: Buffer): string`.
> 2. **Logic:**
> * Write the buffer to `visio/media/{filename}` in the JSZip object.
> * Return the internal path (e.g., `../media/image1.png`).
>
>
> 3. **Content Types:** Check `[Content_Types].xml`. Ensure that a `<Default Extension="png" ...>` (and jpg/jpeg) exists. If not, append it.
> 4. **Test Requirement:** Write a test that adds a PNG buffer, saves the ZIP, and verifies the file exists in the `visio/media/` folder."
>
>

**Prompt 2: PR Generation (Phase 1)**

> "Generate a Markdown PR description.
> * **Title:** feat: Core Media Handling & Content Type Registration
> * **Description:** Sets up the binary storage layer for images.
> * **Changes:** `MediaManager` class, updates to global `[Content_Types]`.
> * **Verification:** Binary integrity test."
>
>

---

### Phase 2: Relationships & The "Foreign" Shape

*Goal: Link the page to the image file and create the container shape.*

**Prompt 3: Linking Page to Media (.rels)**

> "I need to link a specific Page to a Media file.
> 1. Extend `PageManager` or `RelsManager`.
> 2. Create a method `addPageImageRel(pageId: string, mediaPath: string): string`.
> 3. **Logic:**
> * Open `visio/pages/_rels/page{id}.xml.rels`.
> * Create a new `<Relationship>` pointing to the media path.
> * Type: `http://schemas.microsoft.com/office/2006/relationships/image`.
> * Return the generated `rId` (e.g., `rId5`)."
>
>
>
>

**Prompt 4: The ForeignData Shape XML**

> "Implement the shape generation for images.
> 1. Create a helper `createImageShapeXML(id: number, rId: string, x, y, w, h)`.
> 2. **XML Structure:**
> * The `<Shape>` element must have `Type='Foreign'`.
> * It must contain a `<ForeignData r:id='{rId}'>` element.
> * It *usually* implies a specific line/fill style (invisible borders).
>
>
> 3. **Geometry:** Does a Foreign shape need a `<Geom>` section? Check the Visio schema and advise/implement if necessary."
>
>

**Prompt 5: PR Generation (Phase 2)**

> "Generate a PR description.
> * **Title:** feat: ForeignData Shape Support
> * **Context:** Images in Visio are treated as 'Foreign Objects'.
> * **Technical Detail:** Explain the relationship linkage (`rId`) between the Page XML and the Media file.
> * **Test Plan:** Verify the generated XML includes `<ForeignData>`."
>
>

---

### Phase 3: Public API & Regression

*Goal: A simple API that handles the complexity of "Upload -> Link -> Draw".*

**Prompt 6: The `addImage` API**

> "Refactor the `Page` class to expose `addImage`.
> 1. Signature: `addImage(data: Buffer, name: string, x: number, y: number, width: number, height: number)`.
> 2. **Orchestration:**
> * Call `MediaManager.addMedia()` -> Get Path.
> * Call `RelsManager.addRel()` -> Get rId.
> * Call `createImageShapeXML()` -> Get Shape.
>
>
> 3. **Return:** The new `Shape` object.
> 4. **Test Requirement:** Create `ImageEmbedding.test.ts`. Load a dummy buffer. Add it to a page. Assert that the resulting `.vsdx` file contains the image in the media folder AND the correct XML tags in the page."
>
>

**Prompt 7: Final Documentation & PR**

> "Generate the final PR description.
> * **Title:** feat: Public API for Image Embedding
> * **Usage:** Provide a code example:
> ```typescript
> const logo = fs.readFileSync('logo.png');
> page.addImage(logo, 'logo.png', 1, 10, 2, 1);
>
> ```
>
>
> * **Checklist:** Ensure regression tests for standard shapes (Boxes/Lines) still pass."
>
>

This is a sophisticated feature. Implementing **Containers** differs significantly from Groups because it relies on specific metadata tags (`User.msvStructureType`) that signal the Visio engine to take over behavior when the user opens the file.

Here is the itemized prompt plan to implement true Visio Containers.

### Phase 1: The Container Definition (ShapeSheet Magic)

*Goal: Create a shape that Visio recognizes as a Container, even if it’s currently empty.*

**Prompt 1: Container Metadata & User Cells**

> "I need to define a shape as a 'Container' according to the Visio specification.
> 1. Identify the specific `User-Defined Cells` required to activate container behavior. specifically, details on:
> * `User.msvStructureType` (Should be 'Container')
> * `User.msvSDContainerMargin`
> * `User.msvSDContainerResize`
>
>
> 2. Create a helper method `makeContainer(shapeId: string)` that injects these cells into the ShapeSheet of an existing rectangle.
> 3. **Critical:** Does a Container require a specific `Line` or `Fill` pattern to look like a standard Visio container (e.g., header at top)? Provide the XML for a 'Classic' container geometry."
>
>

**Prompt 2: PR Generation (Phase 1)**

> "Generate a Markdown PR description.
> * **Title:** feat: Core Container Metadata Support
> * **Description:** Implements the `User.msv*` cells that trigger Visio's native container engine.
> * **Verification:** Open the generated file in Visio and verify the 'Container Tools' tab appears when the shape is selected."
>
>

---

### Phase 2: Membership Logic (The Relationship)

*Goal: Tell Visio, 'These specific shapes belong to this container.'*

**Prompt 3: Defining Membership in XML**

> "I need to programmatically add existing shapes to a Container.
> 1. Unlike Groups (which nest `<Shape>` tags), Containers use flat relationships. Explain the specific `<Relationship>` tag attributes required in the `page.xml.rels` or the `relationships.xml` part.
> 2. Implement `Container.addMember(shape: Shape)`.
> 3. **Logic:**
> * Is there a specific `User.msvShapeCategories` tag needed on the *member* shapes to allow them to be contained?
> * How do we define the 'ContainerRelationship' in the XML to ensure that when the user moves the Container, the members move with it?"
>
>
>
>

**Prompt 4: Test Requirement (Membership)**

> "Write a test case:
> 1. Create a Container.
> 2. Create a Box.
> 3. Call `container.addMember(box)`.
> 4. **Assertion:** Verify the output XML contains the correct `Relationship` ID linking the two, or the specific ShapeSheet cell (like `User.msvContainerID`) if that is the preferred linkage method in OpenXML."
>
>

---

### Phase 3: The "Layout Engine" (Simulating Behavior)

*Goal: Since the Visio engine isn't running during file generation, the library must manually calculate the initial size of the container to fit its contents.*

**Prompt 5: Auto-Resize Logic**

> "Since Visio won't auto-resize the container until the user *opens* the file, my library must calculate the initial geometry.
> 1. Implement a `resizeToFit()` method on the `ContainerShape` class.
> 2. **Algorithm:**
> * Iterate through all member shapes.
> * Calculate the bounding box (MinX, MinY, MaxX, MaxY) of the collective members.
> * Add the standard padding defined in `User.msvSDContainerMargin`.
> * Update the Container's `PinX`, `PinY`, `Width`, and `Height` to wrap them perfectly.
> * **Constraint:** Ensure the Container's Z-Order is *behind* the members."
>
>
>
>

**Prompt 6: PR Generation (Phase 3)**

> "Generate a PR description.
> * **Title:** feat: Container Auto-Resize Logic
> * **Context:** Visio is a runtime engine; `ts-visio` is a static generator. We must simulate the 'Expand to fit' behavior at generation time.
> * **Changes:** Added BoundingBox calculation and Z-Order management."
>
>

---

### Phase 4: Lists (The "Table" Use Case)

*Goal: Address your specific use case ("Drag a new column"). In Visio, a Table is actually a "List" container, not just a generic Container.*

**Prompt 7: Implementing Lists (Ordered Containers)**

> "My use case is specifically for Database Tables (Columns). In Visio, this is technically a 'List' container (User.msvStructureType = 'List').
> 1. How does a 'List' differ from a 'Container' in the XML?
> 2. Implement `addListItem(item: Shape, position: index)`.
> 3. **Logic:**
> * Handle `User.msvSDListDirection` (Vertical vs Horizontal).
> * Automatically calculate the Y-position of the new item based on the previous item's height (stacking).
> * Update the List Container's height to include the new item."
>
>
>
>

**Prompt 8: Final Documentation & PR**

> "Generate the final PR description for List Containers.
> * **Title:** feat: Structured List Containers for Tables
> * **Usage Example:**
> ```typescript
> const table = page.addListContainer('Users');
> table.addListItem(new Shape('PK: ID'));
> table.addListItem(new Shape('Username'));
>
> ```
>
>
> * **Outcome:** The generated file behaves like a native Visio database table."
>
>

Here is the itemized prompt plan to implement **Hyperlinks**. This feature is essential for interactive diagrams, allowing shapes to act as navigation buttons for external documentation (Jira/Confluence) or internal navigation (Drill-downs).

### Phase 1: The Hyperlink Section (XML Structure)

*Goal: Create the ShapeSheet infrastructure to hold one or more links.*

**Prompt 1: Understanding Hyperlink XML**

> "I need to implement Hyperlinks in the ShapeSheet.
> 1. Explain the XML structure of `<Section N='Hyperlink'>`.
> 2. Detail the specific Cells required:
> * `Address` (The target URL).
> * `SubAddress` (Used for internal page anchors).
> * `Description` (The tooltip text).
> * `NewWindow` (Boolean).
>
>
> 3. **Multiple Links:** Visio supports multiple links on a single shape (right-click menu). How are multiple rows handled in this section? Provide a sample XML snippet."
>
>

**Prompt 2: Implementing `addHyperlink` Core**

> "Implement `Shape.addHyperlinkRow(address: string, description?: string)`.
> 1. Check if `<Section N='Hyperlink'>` exists. If not, create it.
> 2. Create a new `<Row>` (Visio uses named rows like `Hyperlink.Row_1`).
> 3. Set the `Address` cell to the provided string.
> 4. Set the `Description` cell if provided.
> 5. **Escape Logic:** Ensure the URL is properly XML-escaped (e.g., `&` becomes `&amp;`)."
>
>

**Prompt 3: PR Generation (Phase 1)**

> "Generate a Markdown PR description.
> * **Title:** feat: Core Hyperlink ShapeSheet Support
> * **Description:** Adds the ability to inject `<Section N='Hyperlink'>` into shapes.
> * **Verification:** Unit test checking XML string escaping for URLs with query parameters."
>
>

---

### Phase 2: Internal vs. External Navigation

*Goal: Handle the difference between "Go to Google" and "Go to Page 2".*

**Prompt 4: Internal Page Linking (SubAddress)**

> "I need to support linking to other pages within the `.vsdx` document.
> 1. Refactor `addHyperlinkRow` to accept a `type` or detect the input.
> 2. **Visio Rule:**
> * If External: Set `Address='https://...'`.
> * If Internal: Set `Address=''` and `SubAddress='Page-2'`.
>
>
> 3. **Page Name Lookup:** Since pages can be renamed, should this method accept a `Page` object and look up its name dynamically, or just a string ID?
> 4. Implement a helper `linkToPage(targetPage: Page)` that sets the `SubAddress` correctly."
>
>

**Prompt 5: Test Requirement (Navigation)**

> "Write a comprehensive test suite `Hyperlinks.test.ts`.
> 1. **Case A (External):** Shape links to '[https://jira.com](https://jira.com)'. Assert `Address` is set.
> 2. **Case B (Internal):** Shape links to 'Architecture Page'. Assert `Address` is empty and `SubAddress` is set.
> 3. **Case C (Mailto):** Shape links to 'mailto:support@company.com'.
> 4. **Validation:** Ensure `NewWindow` defaults to '0' (false) unless specified."
>
>

**Prompt 6: PR Generation (Phase 2)**

> "Generate a PR description.
> * **Title:** feat: Internal & External Navigation Logic
> * **Technical Context:** explains the usage of `SubAddress` for internal navigation vs `Address` for web links.
> * **Changes:** Added `linkToPage()` helper."
>
>

---

### Phase 3: Public API & Documentation

*Goal: A fluent, developer-friendly experience.*

**Prompt 7: Fluent API Design**

> "Refactor the `Shape` class to expose high-level methods:
> 1. `shape.toUrl(url: string, description?: string): Shape` (Chainable).
> 2. `shape.toPage(page: Page, description?: string): Shape` (Chainable).
> 3. **Documentation:** Update the JSDoc to explain that these create entries in the right-click menu of the shape in Visio.
> 4. **Example:**
> ```typescript
> const box = page.addShape(...);
> box.toUrl('https://jira.atlassian.com/browse/PROJ-1', 'Open Ticket');
> ```"
>
> ```
>
>
>
>

**Prompt 8: Final Documentation & PR**

> "Generate the final PR description.
> * **Title:** feat: Public Hyperlink API
> * **Usage:** Provide code snippets for linking to JIRA and linking to a 'Details' page.
> * **Readme Update:** Add a section 'Interactivity & Navigation'.
> * **Checklist:** Confirm tests pass for URL escaping."
>
>

Here is the itemized prompt plan to implement **Layers**. This feature allows developers to organize complex diagrams by toggling visibility, locking elements (e.g., "Background"), or separating printing content from viewing content.

### Phase 1: Layer Definitions (The Page Schema)

*Goal: Define the "buckets" that shapes will live in. Layers are defined at the Page level, not the Document level.*

**Prompt 1: Understanding Layer XML**

> "I need to implement Layers in `ts-visio`.
> 1. Explain the XML structure of the `<Layers>` collection within `<PageSheet>` or `<PageContents>`.
> 2. Detail the attributes for a `<Layer>` element:
> * `IX` (Index - is it 0-based or 1-based?)
> * `Name` (The display string)
> * `Visible` (Boolean)
> * `Lock` (Boolean - prevents selection)
> * `Print` (Boolean - visible on screen but not in print)
>
>
> 3. Provide a sample XML snippet of a page with a 'Background' layer (Locked) and a 'Comments' layer (Hidden)."
>
>

**Prompt 2: Implementing `addLayer**`

> "Implement `Page.addLayer(name: string, options?: LayerOptions)`.
> 1. **Index Management:** You must track the next available Layer Index (`IX`) for this page.
> 2. **XML Generation:** Append a new `<Layer>` element to the `<Layers>` collection in `page{N}.xml`.
> 3. **Return:** A `Layer` object (containing the Name and Index) that can be passed to shapes later.
> 4. **Test Requirement:** Write a test that adds 3 layers and asserts their Indices are sequential (0, 1, 2 or 1, 2, 3)."
>
>

**Prompt 3: PR Generation (Phase 1)**

> "Generate a Markdown PR description.
> * **Title:** feat: Core Page Layer Definitions
> * **Description:** Adds the infrastructure to define layers on a page.
> * **Technical Detail:** Explains the management of Layer Indices (`IX`).
> * **Verification:** Unit test checking XML structure for `<Layers>`."
>
>

---

### Phase 2: Layer Membership (Assigning Shapes)

*Goal: Put a shape onto a layer. Note: A shape can belong to multiple layers simultaneously.*

**Prompt 4: Understanding `<LayerMem>**`

> "I need to assign shapes to layers.
> 1. Explain the `<LayerMem>` (Layer Membership) section in the ShapeSheet.
> 2. Specifically, the `<Cell N='LayerMember' V='...'/>`.
> 3. **Syntax:** How does Visio store multiple layers? (e.g., is it semicolon-separated like `'1;4'`?).
> 4. If a shape inherits layers from a Master, how does `<LayerMem>` interact with that?"
>
>

**Prompt 5: Implementing `assignLayer**`

> "Implement `Shape.assignLayer(layer: Layer | number)`.
> 1. Check if the shape has a `<LayerMem>` section. If not, create it.
> 2. **Logic:**
> * Read the existing `LayerMember` value.
> * Append the new Layer Index (handling the separator syntax).
> * **Idempotency:** Ensure we don't add the same layer index twice.
>
>
> 3. **Test Requirement:** Create a shape, add it to 'Layer A' (Index 1) and 'Layer B' (Index 2). Assert the XML value is `V='1;2'`."
>
>

**Prompt 6: PR Generation (Phase 2)**

> "Generate a PR description.
> * **Title:** feat: Shape Layer Membership
> * **Context:** Shapes need to reference the Page's defined layers.
> * **Changes:** Implemented `<LayerMem>` parsing and updating.
> * **Tests:** Validated multi-layer assignment logic."
>
>

---

### Phase 3: Layer Management & Toggle API

*Goal: The ability to say "Hide all Comments" programmatically.*

**Prompt 7: Toggling Visibility & Locking**

> "Implement methods on the `Layer` class to update properties dynamically.
> 1. `layer.setVisible(visible: boolean)`: Updates the `Visible` cell in the Page XML.
> 2. `layer.setLocked(locked: boolean)`: Updates the `Lock` cell.
> 3. **Refactor:** Ensure the `PageManager` can locate the specific `<Layer>` node by its Index to apply these updates."
>
>

**Prompt 8: Public API Design**

> "Refactor the public API for ease of use.
> 1. **Scenario:**
> ```typescript
> const comments = page.addLayer('Comments');
> const box = page.addShape(...);
> box.addToLayer(comments);
>
> ```
>
>
>
>

> ```
> // Later:
> comments.hide(); // Sets Visible=0
> ```
>
> ```
>
>
> 2. Ensure the `Layer` object returned by `addLayer` holds a reference to the Page so it can trigger the XML update."
>
>

**Prompt 9: Final Documentation & PR**

> "Generate the final PR description.
> * **Title:** feat: Public Layer Management API
> * **Usage:** Provide a code example showing how to create a 'Wireframe' mode by hiding a specific layer.
> * **Readme Update:** Document the `Layer` class and `Shape.addToLayer`.
> * **Checklist:** Verify that standard shapes (without layers) still render correctly (Regression)."
>
>


We're building a "Mermaid to Visio" converter.

The biggest technical hurdle here is **Layout**. Mermaid syntax defines *relationships* ("A connects to B"), but Visio files require *explicit coordinates* ("Shape A is at x:1, y:2"). Your converter must act as the layout engine.

We will use `dagre` (Directed Graph Layout) to calculate the layout.
Use tests/assets/test-diagram.mmd as a sample input.

Here is the breakdown of tasks and AI prompts to build this integration.

### Phase 1: Parsing Mermaid (The Input)

*Goal: Turn a raw Mermaid string (e.g., `graph TD; A-->B;`) into a structured JSON object.*
*Note: Mermaid's official parser is browser-heavy. For a Node.js tool, you often need a lightweight AST parser.*

**Task 1: Select and Implement a Parser**

* **Context:** We need to extract nodes, edges, and subgraph groupings from the text.
* **AI Prompt:**
> "I need to parse Mermaid Flowchart syntax in a Node.js environment without using a headless browser.
> 1. Evaluate libraries like `mermaid-parser`, `langium-mermaid`, or creating a custom parser using `chevrotain`.
> 2. Write a function `parseMermaid(text: string)` that returns a generic graph object:
> `{ nodes: [{id, text, type}], edges: [{from, to, text, type}] }`.
> 3. Start with support for basic Flowcharts (`graph TD` or `flowchart LR`)."
>
>



### Phase 2: The Layout Engine (The Math)

*Goal: Calculate X/Y coordinates for the shapes. Visio won't do this for you automatically upon opening.*
*Solution: Use `dagre` (Directed Graph Layout), the same library Mermaid uses internally.*

**Task 2: Integrate Dagre for Coordinate Calculation**

* **Context:** We need to map the parsed nodes to a coordinate system.
* **AI Prompt:**
> "I need to calculate the layout for a flowchart in Node.js.
> 1. Install `dagre`.
> 2. Create a class `GraphLayout`.
> 3. Implement `calculateLayout(nodes, edges, direction)`.
> * Map the parsed nodes to `dagre.setNode(id, { width, height })`. *Note: Assume a default width/height for Visio shapes (e.g., 1x0.75 inches) for now.*
> * Map edges to `dagre.setEdge(from, to)`.
> * Run `dagre.layout(g)`.
>
>
> 4. Return the nodes with their calculated `x` and `y` coordinates."
>
>



### Phase 3: Shape Mapping (The Translation)

*Goal: Convert abstract Mermaid types to your `ts-visio` shapes.*

**Task 3: The Mapper Strategy**

* **Context:** Mermaid has distinct syntax for shapes: `[Box]`, `(Circle)`, `{Rhombus}`, `((Circle))`. You need to map these to Visio Masters or Geometry.
* **AI Prompt:**
> "I need to map Mermaid node syntax to Visio shapes.
> 1. Create a `ShapeMapper` class.
> 2. Implement a switch case based on Mermaid bracket syntax:
> * `[text]` -> Rectangle (Visio Master 'Process')
> * `(text)` -> Rounded Rectangle (Visio Master 'Terminator')
> * `{text}` -> Rhombus (Visio Master 'Decision')
> * `[(text)]` -> Cylinder (Visio Master 'Database')
>
>
> 3. If I am using `ts-visio`, how do I effectively swap the 'Master' ID based on these types?"
>
>



### Phase 4: Integration (The Builder)

*Goal: Orchestrate the Parse -> Layout -> Draw pipeline.*

**Task 4: The Converter Pipeline**

* **Context:** Putting it all together.
* **AI Prompt:**
> "Create a main class `MermaidToVisio`.
> 1. Implement `convert(mermaidText: string, outputPath: string)`.
> 2. **Pipeline Steps:**
> * Call `Parser.parse(text)`.
> * Call `GraphLayout.calculate(graph)`.
> * Initialize `VisioDocument` and `Page`.
> * Loop through nodes: Call `page.addShape()` using the calculated X/Y and mapped Master.
> * Loop through edges: Call `page.connectShapes()` (using the dynamic connector logic we built previously).
> * Save the file."
>
>
>
>



### Phase 5: Advanced Styling (Optional)

*Goal: Support Mermaid's styling syntax `style A fill:#f9f`.*

**Task 5: Parsing and Applying Styles**

* **Context:** Mermaid styles are usually distinct lines at the end of the graph definition.
* **AI Prompt:**
> "Extend the Mermaid parser to extract `style` definitions (e.g., `style id fill:#f9f,stroke:#333`).
> 1. Map these styles to `ts-visio` style objects.
> 2. When creating the shape in the integration loop, apply `.fillColor()` and `.strokeColor()` if a style exists for that Node ID."
>
>



### Recommended Task Order

1. **Phase 1 & 2 (Parser + Layout):** Build a script that just logs JSON with X/Y coordinates. Don't touch Visio yet.
2. **Phase 4 (Integration):** Feed that JSON into `ts-visio` using standard rectangles.
3. **Phase 3 (Shapes):** Make it look pretty with correct shapes.
4. **Phase 5 (Styles):** Add color.