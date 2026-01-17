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