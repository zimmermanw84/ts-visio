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

### ✅ BUG 5 — `as any` Cast on JSZip `.async()` Type
**File:** `src/VisioPackage.ts:44`

`file.async(type as any)` suppresses TypeScript's ability to catch type mismatches on the JSZip API call.

**Fix:** Use a proper union cast: `file.async(type as 'string' | 'nodebuffer')`.

---

### ✅ BUG 6 — `backPageId` Type Mismatch
**Files:** `src/core/PageManager.ts:12`, `src/types/VisioTypes.ts:50`

`PageEntry.backPageId` is typed as `number`; `VisioPage.backPageId` is typed as `string`. The assignment in `VisioDocument.pages` silently coerces between the two.

**Fix:** Changed `PageEntry.backPageId` from `number` to `string`. Parse boundary in `PageManager.load()` now uses `.toString()` instead of `parseInt()`. Regression test added to `tests/PageManager.test.ts`.

---

### ✅ BUG 7 — Two Conflicting `PageManager` Classes
**Files:** `src/PageManager.ts`, `src/core/PageManager.ts`

`index.ts` exports the top-level `src/PageManager.ts` (a simpler implementation without relationship resolution or page creation). `VisioDocument` uses `src/core/PageManager.ts`. Consumers who import `{ PageManager }` from the package get the wrong class.

**Fix:** Deleted `src/PageManager.ts`. Updated `src/index.ts` to re-export `PageManager` and `PageEntry` from `src/core/PageManager.ts`. Updated `tests/BlankCreation.test.ts` to use the correct `load()` method and `PageEntry` field names.

---

### ✅ BUG 8 — `placeRightOf` Position Calculation Wrong
**File:** `src/Shape.ts:71`

```typescript
// Bug — uses full width, not half-width
const newX = targetShape.x + targetShape.width + options.gap;

// Correct — PinX is center, so offset by half-widths on both sides
const newX = targetShape.x + (targetShape.width / 2) + options.gap + (this.width / 2);
```

The gap between shapes was always `targetWidth/2 + thisWidth/2` too wide.

**Fix:** Applied the correct formula. Updated `tests/Layout.test.ts` expected values to match correct geometry.

---

### ✅ BUG 9 — Geometry Rows Use Literal Coordinates — Resize Breaks
**Files:** `src/shapes/ShapeBuilder.ts:58-64`, `src/shapes/ForeignShapeBuilder.ts:36-41`

Geometry row coordinates store raw numbers (e.g. `Width=2`, `Height=1`). Visio requires formula references (`Width*1`, `Height*1`) so that geometry tracks the bounding box when a shape is resized. Shapes created by ts-visio cannot be resized in Visio Desktop.

**Fix:** Added `@_F: 'Width'` / `@_F: 'Height'` formula attributes to the `LineTo` geometry rows in both `ShapeBuilder` and `ForeignShapeBuilder`. Static `@_V` values are retained for initial render.

---

### ✅ BUG 11 — `isBinaryExtension` Only Recognises Image Formats
**File:** `src/core/MediaConstants.ts:13-15`

Only `png`, `jpg`, `jpeg`, `gif`, `bmp`, `tiff` were treated as binary. Loading a `.vsdx` that contains `emf`, `wmf`, `svg`, `ole`, or other binary assets caused them to be decoded as UTF-8 strings, permanently corrupting the data when the file is saved.

**Fix:** Inverted the logic — introduced a `TEXT_EXTENSIONS = new Set(['xml', 'rels'])` allowlist and changed `isBinaryExtension()` to return `!TEXT_EXTENSIONS.has(ext.toLowerCase())`. Unknown extensions now default to binary, preventing corruption of any current or future binary asset type. Also expanded `MIME_TYPES` with `emf`, `wmf`, `svg`, `ico`, `tif`, `wdp`. The `type as any` cast in `VisioPackage.ts` was also replaced with the correct `'string' | 'nodebuffer'` union (BUG 5). Three regression tests added to `tests/VisioPackage.test.ts`.

---

## Medium

### ✅ BUG 12 — New Pages Missing `NameU` Attribute
**File:** `src/core/PageManager.ts:170-174, 245-249`

Pages created by `createPage()` and `createBackgroundPage()` were given `@_ID` and `@_Name` but not `@_NameU` (universal/culture-invariant name), which Visio uses internally for cross-locale references.

**Fix:** Added `'@_NameU': name` to the page node object in both `createPage()` and `createBackgroundPage()`. Two regression tests added to `tests/BackgroundPages.test.ts`.

---

### ✅ BUG 14 — `pageStub` Objects Cast as `any` in `VisioDocument`
**File:** `src/VisioDocument.ts:45, 61, 82`

`pageStub as any` bypasses type checking when constructing `VisioPage` objects. If the `VisioPage` interface changes, these casts will silently produce broken objects.

**Fix:** Typed stubs as `VisioPage` explicitly in all three construction sites (`addPage`, `pages` getter, `addBackgroundPage`). All `as any` casts removed.

---

### ✅ BUG 15 — Every `Shape` / `Layer` Method Creates a New `ShapeModifier`
**Files:** `src/Shape.ts`, `src/Layer.ts`

Each mutating method instantiates `new ShapeModifier(this.pkg)` with its own empty `pageCache`. If `Page.modifier` has pending mutations in its cache and a `Shape` method then creates a separate modifier instance, the two caches diverge and can overwrite each other's changes.

**Fix:** Added optional `modifier?: ShapeModifier` parameter to both `Shape` and `Layer` constructors (defaulting to a fresh instance for standalone use). `Page` now passes `this.modifier` to every `Shape` and `Layer` it constructs, ensuring all mutations share a single cache.

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

### ✅ BUG 21 — `Shape` Methods Create New `ShapeModifier` Inconsistently
**File:** `src/Shape.ts`

`addListItem`, `resizeToFit`, `refreshLocalState`, and all other mutating methods created `new ShapeModifier(this.pkg)` instead of reusing the shared instance, wasting cache and risking incoherence.

**Fix:** Resolved by BUG 15 fix — all methods in `Shape` now use `this.modifier` which is injected from `Page`.

---

### ✅ BUG 22 — `require('crypto')` Incompatible with Pure ESM
**File:** `src/core/MediaManager.ts:27, 49`

Synchronous `require('crypto')` fails in a pure ESM environment. The package is dual-published CJS/ESM.

**Fix:** Replaced both inline `require('crypto')` calls with a top-level `import { createHash } from 'node:crypto'`.

---

### ✅ BUG 23 — Content-Type Extension Comparison Is Case-Sensitive
**File:** `src/core/MediaManager.ts:91`

```typescript
d['@_Extension']?.toLowerCase() === extension  // extension is NOT lowercased
```

A file named `Image.PNG` produces extension `PNG`; `'png' === 'PNG'` is false, causing a duplicate `<Default>` entry in `[Content_Types].xml`.

**Fix:** Changed comparison to `extension.toLowerCase()` so both sides are normalised.

---

### ✅ BUG 24 — Float Precision Loss in `placeRightOf` / `placeBelow`
**File:** `src/Shape.ts:78-82, 98-102`

Local state is updated via `newX.toString()`, then read back via `Number()`. Chained placement operations accumulate floating-point-to-string-to-float rounding errors.

**Fix:** Introduced `fmtCoord(n)` helper that rounds to 10 decimal places before converting to string (`parseFloat(n.toFixed(10)).toString()`). All local-state coordinate writes now use this helper.

---

### ✅ BUG 25 — Background Page Template Missing `<Connects/>` Element
**File:** `src/core/PageManager.ts:213`

The foreground page template includes `<Connects/>` but the background page template does not, causing asymmetry that can trigger issues if connectors are added to a background page.

**Fix:** Added `<Connects/>` to the background page template string in `createBackgroundPage()`.

---

### ✅ BUG 26 — `LineWeight` Cell Missing Unit Attribute
**File:** `src/utils/StyleHelpers.ts:87`

`LineWeight` is emitted as `'0.01'` with no `@_U` unit attribute. Visio defaults to inches, which is probably correct, but the missing unit is inconsistent with other cells and may behave unexpectedly in non-English locales.

**Fix:** Added `'@_U': 'IN'` to the `LineWeight` cell in `createLineSection()`.

---

### ✅ BUG 27 — `ContainerBuilder` Sets TextXform Formula Strings as Values
**File:** `src/shapes/ContainerBuilder.ts:82-83`

```typescript
upsertCell('TxtPinY',    'Height',    'DY');
upsertCell('TxtLocPinY', 'TxtHeight', 'DY');
```

These set `@_V` (value) to the strings `'Height'` and `'TxtHeight'`, which Visio will interpret as `0` or a parse error. They should be `@_F` (formula) attributes.

**Fix:** `upsertCell` now accepts a `formula` and `val` parameter separately. `TxtPinY` gets `@_F: 'Height'` with `@_V` set to the shape's current height; `TxtLocPinY` gets `@_F: 'TxtHeight'` with `@_V: '0'`.

---

## Summary

| Severity | Total | Fixed | Open |
|----------|-------|-------|------|
| Critical | 4 | 4 | 0 |
| High | 6 | 6 | 0 |
| Medium | 6 | 6 | 0 |
| Low | 7 | 7 | 0 |
| **Total** | **23** | **23** | **0** |
