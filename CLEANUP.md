# Codebase Organization Plan

## 1. Split `ShapeModifier.ts` (1,450 lines → 5 focused files)

This is the biggest win. The file currently mixes 6+ distinct concerns:

| New File | Extracted Methods |
|---|---|
| `core/PageSheetEditor.ts` | `setPageSize`, `getPageDimensions`, `setDrawingScale`, `getDrawingScale`, `clearDrawingScale`, `ensurePageSheet`, `updateNextShapeId` |
| `core/LayerEditor.ts` | `addLayer`, `assignLayer`, `updateLayerProperty`, `getPageLayers`, `deleteLayer`, `getShapeLayerIndices` |
| `core/ContainerEditor.ts` | `addContainer`, `addList`, `addListItem`, `resizeContainerToFit`, `getContainerMembers`, `addRelationship`, `reorderShape` |
| `core/ConnectorEditor.ts` | `addConnector` |
| `ShapeModifier.ts` (slimmed) | Shape CRUD + style + geometry + properties + hyperlinks + shared XML cache infrastructure |

Each extracted service shares the same XML cache infrastructure (page parsing, saving, dirty tracking) via a base class or composed `PageXmlCache` object.

---

## 2. Move `ShapeStyle` out of `ShapeModifier.ts`

`ShapeStyle` is defined at the bottom of `ShapeModifier.ts` but is a public-facing type. Move it to `types/VisioTypes.ts` alongside `NewShapeProps` and `ConnectorStyle`.

Also move `HorzAlign` and `VertAlign` from `utils/StyleHelpers.ts` to `types/VisioTypes.ts` — they are fields on `ShapeStyle` and `NewShapeProps` but currently require an internal import to use them.

---

## 3. Extract shared shape-tree traversal to `utils/ShapeTreeUtils.ts`

Shape gathering logic is duplicated across three files: `ShapeModifier`, `ShapeReader`, and `VisioValidator`. Extract to a single utility module:

```ts
// src/utils/ShapeTreeUtils.ts
gatherAllShapes(root: any): any[]
findShapeById(root: any, id: string): any | undefined
buildShapeMap(root: any): Map<string, any>
```

---

## 4. Clean up `index.ts` public API surface

Currently exports internal implementation details. Proposed changes:

**Remove from public exports:**
- `VisioPackage` — raw OPC zip manipulation, not user-facing
- `PageManager` — internal orchestration
- `PageEntry` — internal OPC structure
- `ShapeModifier` — internal XML mutation service
- `ShapeReader` — internal XML read service

**Add to public exports:**
- `HorzAlign`, `VertAlign` — needed to type `ShapeStyle.horzAlign` / `verticalAlign`
- `ShapeStyle` — currently missing from exports even though it is part of the public surface

---

## 5. Move high-level diagram patterns out of `Page.ts`

`addTable()`, `addSwimlanePool()`, and `addSwimlaneLane()` are high-level composition patterns, not core page operations. Move them to a `src/diagrams/` layer:

```
src/diagrams/
  SchemaDiagram.ts    (move from src/ root)
  TablePattern.ts     (addTable logic from Page.ts)
  SwimlanePattern.ts  (addSwimlanePool / addSwimlaneLane logic from Page.ts)
```

`Page` retains thin wrapper methods that delegate to these patterns, or consumers call the pattern classes directly.

---

## 6. Collapse `PageManager` template duplication

`createPage()` and `createBackgroundPage()` share ~120 lines of identical XML template string, differing only by the presence of `@_Background='1'`. Extract to a single private `buildPageXml(name: string, isBackground: boolean)` method.

---

## Summary of affected files

```
src/
  core/
    PageSheetEditor.ts    ← NEW (extracted from ShapeModifier)
    LayerEditor.ts        ← NEW (extracted from ShapeModifier)
    ContainerEditor.ts    ← NEW (extracted from ShapeModifier)
    ConnectorEditor.ts    ← NEW (extracted from ShapeModifier)
    PageManager.ts        ← simplified (template dedup)
  utils/
    ShapeTreeUtils.ts     ← NEW (deduplicated from ShapeModifier, ShapeReader, VisioValidator)
    StyleHelpers.ts       ← remove HorzAlign / VertAlign (moved to types)
  types/
    VisioTypes.ts         ← add ShapeStyle, HorzAlign, VertAlign
  diagrams/
    SchemaDiagram.ts      ← MOVED from src/
    TablePattern.ts       ← NEW (from Page.ts)
    SwimlanePattern.ts    ← NEW (from Page.ts)
  ShapeModifier.ts        ← slimmed to ~600 lines
  Page.ts                 ← slimmed (diagram methods delegate to diagrams/)
  index.ts                ← rationalized public exports
```

---

## Suggested order of implementation

1. **`utils/ShapeTreeUtils.ts`** — pure extraction, no behavior change, unblocks everything else
2. **`index.ts` cleanup** — no runtime behavior change, improves API hygiene immediately
3. **`types/VisioTypes.ts`** — move `ShapeStyle`, `HorzAlign`, `VertAlign`; update all import sites
4. **`ShapeModifier.ts` split** — largest effort; do after shared infrastructure is in place
5. **`diagrams/` extraction** — move `SchemaDiagram`, `addTable`, swimlane logic from `Page.ts`
6. **`PageManager` dedup** — small, self-contained cleanup
