import { XMLParser, XMLBuilder } from 'fast-xml-parser';
import { VisioPackage } from '../VisioPackage';
import { SECTION_NAMES } from './VisioConstants';
import { createXmlParser, createXmlBuilder, buildXml } from '../utils/XmlHelper';

/**
 * Shared XML parsing, caching, and persistence infrastructure for all page editors.
 * Maintains a per-page parsed-object cache so multiple operations on the same page
 * avoid redundant parse/serialize round-trips.
 */
export class PageXmlCache {
    readonly parser: XMLParser;
    readonly builder: XMLBuilder;
    private pageCache: Map<string, { content: string; parsed: any }> = new Map();
    private dirtyPages: Set<string> = new Set();
    readonly shapeCache = new WeakMap<object, Map<string, any>>();
    private pagePathRegistry = new Map<string, string>();
    public autoSave: boolean = true;

    constructor(readonly pkg: VisioPackage) {
        this.parser = createXmlParser();
        this.builder = createXmlBuilder();
    }

    /**
     * Register the resolved OPC part path for a page ID.
     * Must be called before any operation on a loaded file to ensure the
     * correct file is targeted rather than the ID-derived fallback name.
     */
    registerPage(pageId: string, xmlPath: string): void {
        this.pagePathRegistry.set(pageId, xmlPath);
    }

    getPagePath(pageId: string): string {
        return this.pagePathRegistry.get(pageId) ?? `visio/pages/page${pageId}.xml`;
    }

    getShapeMap(parsed: any): Map<string, any> {
        if (!this.shapeCache.has(parsed)) {
            const map = new Map<string, any>();
            let topLevelShapes = parsed.PageContents.Shapes ? parsed.PageContents.Shapes.Shape : [];
            if (!Array.isArray(topLevelShapes)) {
                topLevelShapes = topLevelShapes ? [topLevelShapes] : [];
            }

            const gather = (shapeList: any[]): void => {
                for (const s of shapeList) {
                    map.set(s['@_ID'], s);
                    if (s.Shapes?.Shape) {
                        const children = Array.isArray(s.Shapes.Shape) ? s.Shapes.Shape : [s.Shapes.Shape];
                        gather(children);
                    }
                }
            };

            gather(topLevelShapes);
            this.shapeCache.set(parsed, map);
        }
        return this.shapeCache.get(parsed)!;
    }

    getAllShapes(parsed: any): any[] {
        return Array.from(this.getShapeMap(parsed).values());
    }

    getNextId(parsed: any): string {
        const shapeMap = this.getShapeMap(parsed);
        let maxId = 0;
        for (const s of shapeMap.values()) {
            const id = parseInt(s['@_ID']);
            if (!isNaN(id) && id > maxId) maxId = id;
        }
        const nextId = maxId + 1;
        // NextShapeID stores one beyond the currently assigned ID.
        this.updateNextShapeId(parsed, nextId + 1);
        return nextId.toString();
    }

    ensurePageSheet(parsed: any): void {
        if (!parsed.PageContents.PageSheet) {
            // Enforce element order: PageSheet must precede Shapes and Connects in the XML.
            const shapes = parsed.PageContents.Shapes;
            const connects = parsed.PageContents.Connects;
            const rels = parsed.PageContents.Relationships;

            if (shapes) delete parsed.PageContents.Shapes;
            if (connects) delete parsed.PageContents.Connects;
            if (rels) delete parsed.PageContents.Relationships;

            parsed.PageContents.PageSheet = { Cell: [] };

            if (shapes) parsed.PageContents.Shapes = shapes;
            if (connects) parsed.PageContents.Connects = connects;
            if (rels) parsed.PageContents.Relationships = rels;
        }

        if (!Array.isArray(parsed.PageContents.PageSheet.Cell)) {
            parsed.PageContents.PageSheet.Cell = parsed.PageContents.PageSheet.Cell
                ? [parsed.PageContents.PageSheet.Cell]
                : [];
        }
    }

    updateNextShapeId(parsed: any, nextVal: number): void {
        this.ensurePageSheet(parsed);
        const cells = parsed.PageContents.PageSheet.Cell;
        const cell = cells.find((c: any) => c['@_N'] === 'NextShapeID');
        if (cell) {
            cell['@_V'] = nextVal.toString();
        } else {
            cells.push({ '@_N': 'NextShapeID', '@_V': nextVal.toString() });
        }
    }

    getParsed(pageId: string): any {
        const pagePath = this.getPagePath(pageId);
        let content: string;
        try {
            content = this.pkg.getFileText(pagePath);
        } catch {
            throw new Error(`Could not find page file for ID ${pageId}. Expected at ${pagePath}`);
        }

        const cached = this.pageCache.get(pagePath);
        if (cached && cached.content === content) {
            return cached.parsed;
        }

        const parsed = this.parser.parse(content);
        this.pageCache.set(pagePath, { content, parsed });
        return parsed;
    }

    saveParsed(pageId: string, parsed: any): void {
        const pagePath = this.getPagePath(pageId);

        if (!this.autoSave) {
            this.dirtyPages.add(pagePath);
            return;
        }

        this.performSave(pagePath, parsed);
    }

    performSave(pagePath: string, parsed: any): void {
        const newXml = buildXml(this.builder, parsed);
        this.pkg.updateFile(pagePath, newXml);
        this.pageCache.set(pagePath, { content: newXml, parsed });
    }

    flush(): void {
        for (const pagePath of this.dirtyPages) {
            const cached = this.pageCache.get(pagePath);
            if (cached?.parsed) {
                this.performSave(pagePath, cached.parsed);
            }
        }
        this.dirtyPages.clear();
    }

    // ── Shape geometry helpers (used by multiple editors) ────────────────────

    getShapeGeometry(pageId: string, shapeId: string): { x: number; y: number; width: number; height: number } {
        const parsed = this.getParsed(pageId);
        const shape = this.getShapeMap(parsed).get(shapeId);
        if (!shape) throw new Error(`Shape ${shapeId} not found`);

        const getCellVal = (name: string) => {
            if (!shape.Cell) return 0;
            const cells = Array.isArray(shape.Cell) ? shape.Cell : [shape.Cell];
            const c = cells.find((cell: any) => cell['@_N'] === name);
            return c ? Number(c['@_V']) : 0;
        };

        return {
            x: getCellVal('PinX'),
            y: getCellVal('PinY'),
            width: getCellVal('Width'),
            height: getCellVal('Height'),
        };
    }

    async updateShapePosition(pageId: string, shapeId: string, x: number, y: number): Promise<void> {
        const parsed = this.getParsed(pageId);
        const shape = this.getShapeMap(parsed).get(shapeId);
        if (!shape) throw new Error(`Shape ${shapeId} not found on page ${pageId}`);

        if (!shape.Cell) {
            shape.Cell = [];
        } else if (!Array.isArray(shape.Cell)) {
            shape.Cell = [shape.Cell];
        }

        const upsert = (name: string, value: string) => {
            const cell = shape.Cell.find((c: any) => c['@_N'] === name);
            if (cell) cell['@_V'] = value;
            else shape.Cell.push({ '@_N': name, '@_V': value });
        };

        upsert('PinX', x.toString());
        upsert('PinY', y.toString());

        this.saveParsed(pageId, parsed);
    }

    async updateShapeDimensions(pageId: string, shapeId: string, w: number, h: number): Promise<void> {
        const parsed = this.getParsed(pageId);
        const shape = this.getShapeMap(parsed).get(shapeId);
        if (!shape) throw new Error(`Shape ${shapeId} not found`);

        if (!shape.Cell) shape.Cell = [];
        if (!Array.isArray(shape.Cell)) shape.Cell = [shape.Cell];

        const upsert = (name: string, val: string) => {
            const cell = shape.Cell.find((c: any) => c['@_N'] === name);
            if (cell) cell['@_V'] = val;
            else shape.Cell.push({ '@_N': name, '@_V': val });
        };

        upsert('Width', w.toString());
        upsert('Height', h.toString());

        this.saveParsed(pageId, parsed);
    }
}
