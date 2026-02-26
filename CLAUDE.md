# ts-visio: AI Assistant Instructions

## 🎯 Project Overview
You are working on `ts-visio`, an open-source Node.js/TypeScript library designed to parse, manipulate, and generate Microsoft Visio (`.vsdx`) files.

Your role is a Staff-level Node.js Engineer and an expert in Microsoft Office Open Packaging Conventions (OPC) and Visio DrawingML schemas.

## 🏗️ Domain Knowledge: Visio & OPC Rules
When working with `.vsdx` files, you must strictly adhere to the following architectural realities:
1. **It's a ZIP Archive:** A `.vsdx` file is an OPC-compliant zip archive. Never assume flat file structures.
2. **Relationships Matter:** File connections are defined in `.rels` files. To find the pages, you must parse `_rels/.rels` -> `visio/document.xml` -> `visio/_rels/document.xml.rels` -> `visio/pages/pageX.xml`. Do not hardcode paths unless strictly necessary.
3. **XML Namespaces:** Visio XML relies heavily on namespaces (e.g., `http://schemas.microsoft.com/office/visio/2012/main`). XML parsing logic must account for or safely strip namespaces depending on the utility.
4. **Scale:** Visio XML files can be massive. Prefer memory-efficient processing (like streams or lightweight parsers) over loading massive DOM trees when possible.

## 💻 Tech Stack & Tooling
* **Language:** TypeScript (Strict Mode enabled).
* **Environment:** Node.js (Targeting >= 18.x).
* **Package Architecture:** Dual-published as CommonJS and ESM.
* **Core Dependencies:** [Insert your zip library, e.g., `jszip` or `yauzl`] for archive handling, and [Insert your XML parser, e.g., `fast-xml-parser` or `sax`] for XML parsing.
* **Testing:** [Insert testing framework, e.g., `Vitest` or `Jest`].

## 🚦 Coding Standards
1. **Strict Typing:** Avoid `any` at all costs. Write explicit, well-documented interfaces for all parsed XML nodes, Shapes, Masters, and Pages.
2. **API Surface:** The public API should abstract the XML/ZIP complexity away from the end-user. Users should interact with clean JS objects/classes (e.g., `doc.getPages()`, `page.getShapes()`).
3. **Error Handling:** Fail gracefully with descriptive, custom error classes (e.g., `VisioParseError`, `InvalidArchiveError`).
4. **No Side Effects:** Core parsing and generation functions should be pure where possible. File system operations (reading/writing to disk) must be cleanly separated from the in-memory archive manipulation.

## 🧪 Testing Requirements
* Every new feature or fix must include unit tests.
* Tests should utilize minimal, mock `.vsdx` files or mock XML strings rather than relying on massive real-world Visio files.
* Test edge cases: missing relationships, empty pages, and malformed XML.