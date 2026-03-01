# ts-visio Bug Report

Audit performed 2026-02-28. 23 bugs identified across Critical / High / Medium / Low severity tiers.

---

## Status Legend

| Symbol | Meaning |
|--------|---------|
| ✅ | Fixed |
| 🔴 | Open — High |
| 🟡 | Open — Medium |
| 🟢 | Open — Low |

---

## Critical (all fixed)

### ✅ BUG 1 — XML Declaration and Namespaces Stripped on Every Write
**Files:** `src/ShapeModifier.ts`, `src/core/PageManager.ts`, `src/core/RelsManager.ts`, `src/core/MediaManager.ts`

`XMLBuilder` was not configured to preserve the `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` declaration. Every write operation (`addShape`, `addPage`, `addConnector`, etc.) stripped it, producing files Visio refuses to open.

**Fix:** Created `src/utils/XmlHelper.ts` with `createXmlParser()` (`ignoreDeclaration: false`), `createXmlBuilder()` (`suppressBooleanAttributes: false`), and `buildXml()` (guarantees declaration is always prepended). All four classes updated to use the shared helper.

---

### ✅ BUG 2 — `ForeignData` Malformed — Images Never Render
**File:** `src/shapes/ForeignShapeBuilder.ts`

`ForeignData: { '@_r:id': rId }` serialized to `<ForeignData r:id="rId1"/>`. The Visio spec requires the `r:id` on a `<Rel>` child element and a `ForeignType` attribute on `ForeignData`.

**Fix:**
```typescript
ForeignData: {
    '@_ForeignType': 'Bitmap',
    Rel: { '@_r:id': rId }
}
```

---

### ✅ BUG 3 — Page ID-to-Path Hardcoded
**Files:** `src/ShapeModifier.ts`, `src/Page.ts`

`getPagePath(pageId)` returned `visio/pages/page${pageId}.xml`. In loaded files with deleted or reordered pages, page ID 3 might live in `page1.xml`, causing the wrong file to be read or written.

**Fix:** Added `pagePathRegistry` map and `registerPage()` to `ShapeModifier`. `Page` reads `internalPage.xmlPath` (populated from the .rels-resolved `PageEntry.xmlPath`) with an ID-derived fallback for newly created pages. `VisioDocument` threads the resolved path into all `Page` stubs.

---

### ✅ BUG 4 — `addContainer` / `addList` Bypass the Page Cache
**File:** `src/ShapeModifier.ts`

Both methods called `pkg.getFileText()` + `parser.parse()` directly instead of `getParsed()`, and wrote via `pkg.updateFile()` directly instead of `saveParsed()`. With `autoSave: false`, pending shape mutations in the cache were overwritten by `addContainer`/`addList` reading stale disk content, and `flush()` then wrote the old parsed object back — silently dropping either the shapes or the container.

**Fix:** Both methods now use `getParsed(pageId)` and `saveParsed(pageId, parsed)`, keeping them coherent with the shared page cache and dirty-page tracking.

---

## High

### 🔴 BUG 5 — `as any` Cast on JSZip `.async()` Type
**File:** `src/VisioPackage.ts:44`

`file.async(type as any)` suppresses TypeScript's ability to catch type mismatches on the JSZip API call.

**Fix:** Use a proper union cast: `file.async(type as 'string' | 'nodebuffer')`.

---

### 🔴 BUG 6 — `backPageId` Type Mismatch
**Files:** `src/core/PageManager.ts:12`, `src/types/VisioTypes.ts:50`

`PageEntry.backPageId` is typed as `number`; `VisioPage.backPageId` is typed as `string`. The assignment in `VisioDocument.pages` silently coerces between the two.

**Fix:** Settle on `string` throughout and convert at the parse boundary in `PageManager.load()`.

---

### 🔴 BUG 7 — Two Conflicting `PageManager` Classes
**Files:** `src/PageManager.ts`, `src/core/PageManager.ts`

`index.ts` exports the top-level `src/PageManager.ts` (a simpler implementation without relationship resolution or page creation). `VisioDocument` uses `src/core/PageManager.ts`. Consumers who import `{ PageManager }` from the package get the wrong class.

**Fix:** Either remove `src/PageManager.ts` or update `index.ts` to re-export from `src/core/PageManager.ts`.

---

### 🔴 BUG 8 — `placeRightOf` Position Calculation Wrong
**File:** `src/Shape.ts:71`

```typescript
// Bug — uses full width, not half-width
const newX = targetShape.x + targetShape.width + options.gap;

// Correct — PinX is center, so offset by half-widths on both sides
const newX = targetShape.x + (targetShape.width / 2) + options.gap + (this.width / 2);
```

The gap between shapes is always `targetWidth/2 + thisWidth/2` too wide.

---

### 🔴 BUG 9 — Geometry Rows Use Literal Coordinates — Resize Breaks
**Files:** `src/shapes/ShapeBuilder.ts:58-64`, `src/shapes/ForeignShapeBuilder.ts:36-41`

Geometry row coordinates store raw numbers (e.g. `Width=2`, `Height=1`). Visio requires formula references (`Width*1`, `Height*1`) so that geometry tracks the bounding box when a shape is resized. Shapes created by ts-visio cannot be resized in Visio Desktop.

**Fix:** Add `@_F` formula attributes (`'Width*1'`, `'Height*1'`) on geometry `LineTo` rows, or emit the standard Visio `Width` / `Height` cell references.

---

### 🔴 BUG 11 — `isBinaryExtension` Only Recognises Image Formats
**File:** `src/core/MediaConstants.ts:13-15`

Only `png`, `jpg`, `jpeg`, `gif`, `bmp`, `tiff` are treated as binary. Loading a `.vsdx` that contains `emf`, `wmf`, `svg`, `ole`, or other binary assets causes them to be decoded as UTF-8 strings, permanently corrupting the data when the file is saved.

**Fix:** Expand the binary extension list to include at minimum `emf`, `wmf`, `svg`, `ole`, `bin`, `ico`, `tif`, `wdp`. Consider defaulting to binary for unknown extensions.

---

## Medium

### 🟡 BUG 12 — New Pages Missing `NameU` Attribute
**File:** `src/core/PageManager.ts:170-174, 245-249`

Pages created by `createPage()` and `createBackgroundPage()` are given `@_ID` and `@_Name` but not `@_NameU` (universal/culture-invariant name), which Visio uses internally for cross-locale references.

**Fix:** Add `'@_NameU': name` to the page node object in both creation paths.

---

### 🟡 BUG 14 — `pageStub` Objects Cast as `any` in `VisioDocument`
**File:** `src/VisioDocument.ts:45, 61, 82`

`pageStub as any` bypasses type checking when constructing `VisioPage` objects. If the `VisioPage` interface changes, these casts will silently produce broken objects.

**Fix:** Introduce a typed factory function or use `Partial<VisioPage>` with explicit runtime validation.

---

### 🟡 BUG 15 — Every `Shape` / `Layer` Method Creates a New `ShapeModifier`
**Files:** `src/Shape.ts`, `src/Layer.ts`

Each mutating method instantiates `new ShapeModifier(this.pkg)` with its own empty `pageCache`. If `Page.modifier` has pending mutations in its cache and a `Shape` method then creates a separate modifier instance, the two caches diverge and can overwrite each other's changes.

**Fix:** Pass the owning page's `ShapeModifier` instance into `Shape` and `Layer` at construction time instead of creating new instances per method call.

---

### ✅ BUG 18 — `MasterManager.load()` Crashes on Empty `<Masters/>` Element
**File:** `src/core/MasterManager.ts:36-38`

When `parsedMasters.Masters.Master` is `undefined` (empty `<Masters/>` element), the normalization code wraps it as `[undefined]`. The subsequent `.map()` call then attempts `undefined['@_ID']`, throwing a `TypeError` at runtime.

**Fix:**
```typescript
masterNodes = masterNodes ? [masterNodes] : [];
```

---

### ✅ BUG 19 — `ConnectorBuilder.buildShapeHierarchy` Crashes on Empty `<Shapes/>`
**File:** `src/shapes/ConnectorBuilder.ts:66-68`

When a page has a `<Shapes/>` element with no `Shape` children, `parsed.PageContents.Shapes.Shape` is `undefined`. The code wrapped it as `[undefined]`, which caused `mapHierarchy` to crash accessing `undefined['@_ID']`.

**Fix:**
```typescript
const rawShapes = parsed.PageContents.Shapes?.Shape;
const topShapes = Array.isArray(rawShapes) ? rawShapes : rawShapes ? [rawShapes] : [];
```

---

### ✅ BUG 20 — `addConnector` on an Empty Page Produces `[undefined, connector]`
**File:** `src/ShapeModifier.ts` (`addConnector`)

When `parsed.PageContents.Shapes` exists but `.Shape` is `undefined` (parsed from `<Shapes/>`), the normalization in `addConnector`:
```typescript
parsed.PageContents.Shapes.Shape = [parsed.PageContents.Shapes.Shape]; // [undefined]
```
produced a phantom `undefined` entry before the connector shape, resulting in malformed XML.

**Fix:** Added null guard matching the pattern used in `addShape`, `addContainer`, and `addList`:
```typescript
parsed.PageContents.Shapes.Shape = parsed.PageContents.Shapes.Shape ? [parsed.PageContents.Shapes.Shape] : [];
```

---

## Low

### 🟢 BUG 21 — `Page` Methods Create New `ShapeModifier` Inconsistently
**File:** `src/Page.ts`

`addListItem`, `resizeToFit`, and `refreshLocalState` create `new ShapeModifier(this.pkg)` instead of reusing `this.modifier`, wasting cache and risking incoherence.

---

### 🟢 BUG 22 — `require('crypto')` Incompatible with Pure ESM
**File:** `src/core/MediaManager.ts:27, 49`

Synchronous `require('crypto')` fails in a pure ESM environment. The package is dual-published CJS/ESM.

**Fix:** `import { createHash } from 'node:crypto'` at the top of the file.

---

### 🟢 BUG 23 — Content-Type Extension Comparison Is Case-Sensitive
**File:** `src/core/MediaManager.ts:91`

```typescript
d['@_Extension']?.toLowerCase() === extension  // extension is NOT lowercased
```

A file named `Image.PNG` produces extension `PNG`; `'png' === 'PNG'` is false, causing a duplicate `<Default>` entry in `[Content_Types].xml`.

**Fix:** `extension.toLowerCase()` before comparison.

---

### 🟢 BUG 24 — Float Precision Loss in `placeRightOf` / `placeBelow`
**File:** `src/Shape.ts:78-82, 98-102`

Local state is updated via `newX.toString()`, then read back via `Number()`. Chained placement operations accumulate floating-point-to-string-to-float rounding errors.

---

### 🟢 BUG 25 — Background Page Template Missing `<Connects/>` Element
**File:** `src/core/PageManager.ts:213`

The foreground page template includes `<Connects/>` but the background page template does not, causing asymmetry that can trigger issues if connectors are added to a background page.

---

### 🟢 BUG 26 — `LineWeight` Cell Missing Unit Attribute
**File:** `src/utils/StyleHelpers.ts:87`

`LineWeight` is emitted as `'0.01'` with no `@_U` unit attribute. Visio defaults to inches, which is probably correct, but the missing unit is inconsistent with other cells and may behave unexpectedly in non-English locales.

---

### 🟢 BUG 27 — `ContainerBuilder` Sets TextXform Formula Strings as Values
**File:** `src/shapes/ContainerBuilder.ts:82-83`

```typescript
upsertCell('TxtPinY',    'Height',    'DY');
upsertCell('TxtLocPinY', 'TxtHeight', 'DY');
```

These set `@_V` (value) to the strings `'Height'` and `'TxtHeight'`, which Visio will interpret as `0` or a parse error. They should be `@_F` (formula) attributes.

**Fix:** Move the formula string to `@_F` and set `@_V` to the computed numeric value.

---

## Summary

| Severity | Total | Fixed | Open |
|----------|-------|-------|------|
| Critical | 4 | 4 | 0 |
| High | 6 | 1 | 5 |
| Medium | 6 | 3 | 3 |
| Low | 7 | 0 | 7 |
| **Total** | **23** | **8** | **15** |
