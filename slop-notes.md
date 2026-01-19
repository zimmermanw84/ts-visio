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
> 5. Generate the final PR description."
>
>