import { XMLParser } from 'fast-xml-parser';
import { VisioPackage } from './VisioPackage';
import { VisioShape, ConnectorRouting, ConnectionTarget } from './types/VisioTypes';
import { asArray, parseCells, parseSection } from './utils/VisioParsers';
import { ConnectorData } from './Connector';
import { SECTION_NAMES } from './core/VisioConstants';

export class ShapeReader {
    private parser: XMLParser;

    constructor(private pkg: VisioPackage) {
        this.parser = new XMLParser({
            ignoreAttributes: false,
            attributeNamePrefix: "@_"
        });
    }

    readShapes(path: string): VisioShape[] {
        let content: string;
        try {
            content = this.pkg.getFileText(path);
        } catch {
            return [];
        }

        const parsed = this.parser.parse(content);
        const shapesData = parsed.PageContents?.Shapes?.Shape;

        if (!shapesData) return [];

        return asArray<any>(shapesData).map(s => this.parseShape(s));
    }

    /**
     * Returns every shape on the page flattened into a single array,
     * including shapes nested inside groups at any depth.
     */
    readAllShapes(path: string): VisioShape[] {
        let content: string;
        try {
            content = this.pkg.getFileText(path);
        } catch {
            return [];
        }

        const parsed = this.parser.parse(content);
        const shapesData = parsed.PageContents?.Shapes?.Shape;
        if (!shapesData) return [];

        const result: VisioShape[] = [];
        this.gatherShapes(asArray<any>(shapesData), result);
        return result;
    }

    /**
     * Find a single shape by ID anywhere in the page tree (including nested groups).
     * Returns undefined if not found.
     */
    readShapeById(path: string, shapeId: string): VisioShape | undefined {
        let content: string;
        try {
            content = this.pkg.getFileText(path);
        } catch {
            return undefined;
        }

        const parsed = this.parser.parse(content);
        const shapesData = parsed.PageContents?.Shapes?.Shape;
        if (!shapesData) return undefined;

        return this.findShapeById(asArray<any>(shapesData), shapeId);
    }

    /**
     * Read all connector shapes from a page XML file.
     * A shape is considered a connector if it has `ObjType=2` or a `BeginX` cell,
     * and is referenced in the page's `<Connects>` section.
     */
    readConnectors(path: string): ConnectorData[] {
        let content: string;
        try {
            content = this.pkg.getFileText(path);
        } catch {
            return [];
        }

        const parsed = this.parser.parse(content);
        const shapesData = parsed.PageContents?.Shapes?.Shape;
        if (!shapesData) return [];

        // Build a flat map of all shapes by ID (including nested)
        const shapeMap = new Map<string, any>();
        const collectShapes = (list: any[]): void => {
            for (const s of list) {
                shapeMap.set(s['@_ID'], s);
                if (s.Shapes?.Shape) collectShapes(asArray<any>(s.Shapes.Shape));
            }
        };
        collectShapes(asArray<any>(shapesData));

        // Identify connector shapes by ObjType=2 or presence of BeginX cell
        const connectorShapes: any[] = [];
        for (const [, shape] of shapeMap) {
            const cells = parseCells(shape);
            if (cells['ObjType']?.V === '2' || cells['BeginX'] !== undefined) {
                connectorShapes.push(shape);
            }
        }
        if (connectorShapes.length === 0) return [];

        // Group <Connect> elements by connector ID (FromSheet)
        const connectsRaw = parsed.PageContents?.Connects?.Connect;
        const connects = asArray<any>(connectsRaw);
        const connectsByConnector = new Map<string, { beginConnect?: any; endConnect?: any }>();
        for (const c of connects) {
            const fromSheet = c['@_FromSheet'];
            const fromCell  = c['@_FromCell'];
            if (!connectsByConnector.has(fromSheet)) {
                connectsByConnector.set(fromSheet, {});
            }
            const entry = connectsByConnector.get(fromSheet)!;
            if (fromCell === 'BeginX') entry.beginConnect = c;
            else if (fromCell === 'EndX') entry.endConnect = c;
        }

        const ROUTING_BY_VALUE: Record<string, ConnectorRouting> = {
            '1': 'orthogonal',
            '2': 'straight',
            '16': 'curved',
        };

        const result: ConnectorData[] = [];

        for (const connShape of connectorShapes) {
            const connId = connShape['@_ID'];
            const entry  = connectsByConnector.get(connId);
            if (!entry?.beginConnect || !entry?.endConnect) continue;

            const { beginConnect, endConnect } = entry;
            const fromShapeId = beginConnect['@_ToSheet'];
            const toShapeId   = endConnect['@_ToSheet'];

            const fromPort = this.decodeToPart(beginConnect['@_ToPart']);
            const toPort   = this.decodeToPart(endConnect['@_ToPart']);

            // Extract line style from the connector's Line section
            const sections = asArray<any>(connShape.Section);
            let lineColor: string | undefined;
            let lineWeight: number | undefined;
            let linePattern: number | undefined;
            for (const sec of sections) {
                if (sec['@_N'] === SECTION_NAMES.Line) {
                    const lineCells = parseCells(sec);
                    if (lineCells['LineColor']?.V)  lineColor   = lineCells['LineColor'].V;
                    if (lineCells['LineWeight']?.V)  lineWeight  = parseFloat(lineCells['LineWeight'].V) * 72; // in→pt
                    if (lineCells['LinePattern']?.V) linePattern = parseInt(lineCells['LinePattern'].V, 10);
                }
            }

            const cells = parseCells(connShape);
            const routeStyleVal = cells['ShapeRouteStyle']?.V;
            const routing = routeStyleVal ? ROUTING_BY_VALUE[routeStyleVal] : undefined;

            const style: import('./types/VisioTypes').ConnectorStyle = {};
            if (lineColor  !== undefined) style.lineColor  = lineColor;
            if (lineWeight !== undefined) style.lineWeight  = lineWeight;
            if (linePattern !== undefined) style.linePattern = linePattern;
            if (routing    !== undefined) style.routing     = routing;

            result.push({
                id: connId,
                fromShapeId,
                toShapeId,
                fromPort,
                toPort,
                style,
                beginArrow: cells['BeginArrow']?.V ?? '0',
                endArrow:   cells['EndArrow']?.V   ?? '0',
            });
        }

        return result;
    }

    /** Decode a Visio ToPart integer string to a ConnectionTarget. */
    private decodeToPart(toPart: string | undefined): ConnectionTarget {
        const part = toPart !== undefined ? parseInt(toPart, 10) : 3;
        if (isNaN(part) || part === 3) return 'center';
        if (part >= 100) return { index: part - 100 };
        return 'center';
    }

    private gatherShapes(rawShapes: any[], result: VisioShape[]): void {
        for (const s of rawShapes) {
            result.push(this.parseShape(s));
            if (s.Shapes?.Shape) {
                this.gatherShapes(asArray<any>(s.Shapes.Shape), result);
            }
        }
    }

    /**
     * Return the direct child shapes of a group or container shape.
     * Returns an empty array if the shape has no children or does not exist.
     */
    readChildShapes(path: string, parentId: string): VisioShape[] {
        let content: string;
        try {
            content = this.pkg.getFileText(path);
        } catch {
            return [];
        }

        const parsed = this.parser.parse(content);
        const shapesData = parsed.PageContents?.Shapes?.Shape;
        if (!shapesData) return [];

        const rawParent = this.findRawShape(asArray<any>(shapesData), parentId);
        if (!rawParent?.Shapes?.Shape) return [];

        return asArray<any>(rawParent.Shapes.Shape).map((s: any) => this.parseShape(s));
    }

    private findShapeById(rawShapes: any[], shapeId: string): VisioShape | undefined {
        const raw = this.findRawShape(rawShapes, shapeId);
        return raw ? this.parseShape(raw) : undefined;
    }

    private findRawShape(rawShapes: any[], shapeId: string): any | undefined {
        for (const s of rawShapes) {
            if (s['@_ID'] === shapeId) return s;
            if (s.Shapes?.Shape) {
                const found = this.findRawShape(asArray<any>(s.Shapes.Shape), shapeId);
                if (found) return found;
            }
        }
        return undefined;
    }

    private parseShape(s: any): VisioShape {
        const shape: VisioShape = {
            ID: s['@_ID'],
            Name: s['@_Name'],
            NameU: s['@_NameU'],
            Type: s['@_Type'],
            Master: s['@_Master'],
            Text: s.Text?.['#text'] || (typeof s.Text === 'string' ? s.Text : undefined),
            Cells: parseCells(s),
            Sections: {}
        };

        const sections = asArray(s.Section);
        for (const sec of sections) {
            const section = sec as any;
            if (section['@_N']) {
                shape.Sections[section['@_N']] = parseSection(section);
            }
        }

        return shape;
    }
}
