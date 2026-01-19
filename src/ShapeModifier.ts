import { XMLParser, XMLBuilder } from 'fast-xml-parser';
import { VisioPackage } from './VisioPackage';
import { RelsManager } from './core/RelsManager';
import { createFillSection, createCharacterSection, createLineSection } from './utils/StyleHelpers';

export interface NewShapeProps {
    text: string;
    x: number;
    y: number;
    width: number;
    height: number;
    id?: string;
    fillColor?: string;
    fontColor?: string;
    bold?: boolean;
    type?: string;
    masterId?: string;
}

export class ShapeModifier {
    private parser: XMLParser;
    private builder: XMLBuilder;
    private relsManager: RelsManager;

    constructor(private pkg: VisioPackage) {
        this.parser = new XMLParser({
            ignoreAttributes: false,
            attributeNamePrefix: "@_"
        });
        this.builder = new XMLBuilder({
            ignoreAttributes: false,
            attributeNamePrefix: "@_",
            format: true
        });
        this.relsManager = new RelsManager(pkg);
    }

    private getPagePath(pageId: string): string {
        return `visio/pages/page${pageId}.xml`;
    }

    private getAllShapes(parsed: any): any[] {
        let topLevelShapes = parsed.PageContents.Shapes ? parsed.PageContents.Shapes.Shape : [];
        if (!Array.isArray(topLevelShapes)) {
            topLevelShapes = topLevelShapes ? [topLevelShapes] : [];
        }

        const gather = (shapeList: any[]): any[] => {
            let all: any[] = [];
            for (const s of shapeList) {
                all.push(s);
                if (s.Shapes && s.Shapes.Shape) {
                    const children = Array.isArray(s.Shapes.Shape) ? s.Shapes.Shape : [s.Shapes.Shape];
                    all = all.concat(gather(children));
                }
            }
            return all;
        };

        return gather(topLevelShapes);
    }

    private getNextId(parsed: any): string {
        const allShapes = this.getAllShapes(parsed);
        let maxId = 0;
        for (const s of allShapes) {
            const id = parseInt(s['@_ID']);
            if (!isNaN(id) && id > maxId) maxId = id;
        }
        return (maxId + 1).toString();
    }
    private getCellVal(shape: any, name: string): string {
        if (!shape || !shape.Cell) return '0';
        const cell = shape.Cell.find((c: any) => c['@_N'] === name);
        return cell ? cell['@_V'] : '0';
    }

    private buildShapeHierarchy(parsed: any): Map<string, { shape: any; parent: any }> {
        const shapeHierarchy = new Map<string, { shape: any; parent: any }>();
        const mapHierarchy = (shapes: any[], parent: any | null) => {
            for (const s of shapes) {
                shapeHierarchy.set(s['@_ID'], { shape: s, parent });
                if (s.Shapes && s.Shapes.Shape) {
                    const children = Array.isArray(s.Shapes.Shape) ? s.Shapes.Shape : [s.Shapes.Shape];
                    mapHierarchy(children, s);
                }
            }
        };
        const topShapes = parsed.PageContents.Shapes ?
            (Array.isArray(parsed.PageContents.Shapes.Shape) ? parsed.PageContents.Shapes.Shape : [parsed.PageContents.Shapes.Shape])
            : [];
        mapHierarchy(topShapes, null);
        return shapeHierarchy;
    }

    private getAbsolutePos(id: string, shapeHierarchy: Map<string, { shape: any; parent: any }>): { x: number, y: number } {
        const entry = shapeHierarchy.get(id);
        if (!entry) return { x: 0, y: 0 };

        const shape = entry.shape;
        const pinX = parseFloat(this.getCellVal(shape, 'PinX'));
        const pinY = parseFloat(this.getCellVal(shape, 'PinY'));

        if (!entry.parent) {
            return { x: pinX, y: pinY };
        }

        const parentPos = this.getAbsolutePos(entry.parent['@_ID'], shapeHierarchy);
        const parentLocPinX = parseFloat(this.getCellVal(entry.parent, 'LocPinX'));
        const parentLocPinY = parseFloat(this.getCellVal(entry.parent, 'LocPinY'));

        return {
            x: (parentPos.x - parentLocPinX) + pinX,
            y: (parentPos.y - parentLocPinY) + pinY
        };
    }

    private getEdgePoint(cx: number, cy: number, w: number, h: number, targetX: number, targetY: number): { x: number, y: number } {
        const dx = targetX - cx;
        const dy = targetY - cy;

        if (dx === 0 && dy === 0) return { x: cx, y: cy };

        const rad = Math.atan2(dy, dx);
        const rw = w / 2;
        const rh = h / 2;

        const tx = dx !== 0 ? (dx > 0 ? rw : -rw) / Math.cos(rad) : Infinity;
        const ty = dy !== 0 ? (dy > 0 ? rh : -rh) / Math.sin(rad) : Infinity;

        const t = Math.min(Math.abs(tx), Math.abs(ty));

        return {
            x: cx + t * Math.cos(rad),
            y: cy + t * Math.sin(rad)
        };
    }

    private calculateConnectorLayout(
        fromShapeId: string,
        toShapeId: string,
        shapeHierarchy: Map<string, { shape: any; parent: any }>
    ) {
        let beginX = 0, beginY = 0, endX = 0, endY = 0;
        let sourceGeom: { x: number, y: number, w: number, h: number } | null = null;
        let targetGeom: { x: number, y: number, w: number, h: number } | null = null;

        const sourceEntry = shapeHierarchy.get(fromShapeId);
        const targetEntry = shapeHierarchy.get(toShapeId);

        if (sourceEntry) {
            const abs = this.getAbsolutePos(fromShapeId, shapeHierarchy);
            const w = parseFloat(this.getCellVal(sourceEntry.shape, 'Width'));
            const h = parseFloat(this.getCellVal(sourceEntry.shape, 'Height'));
            sourceGeom = { x: abs.x, y: abs.y, w, h };
            beginX = abs.x;
            beginY = abs.y;
        }

        if (targetEntry) {
            const abs = this.getAbsolutePos(toShapeId, shapeHierarchy);
            const w = parseFloat(this.getCellVal(targetEntry.shape, 'Width'));
            const h = parseFloat(this.getCellVal(targetEntry.shape, 'Height'));
            targetGeom = { x: abs.x, y: abs.y, w, h };
            endX = abs.x;
            endY = abs.y;
        }

        if (sourceGeom && targetGeom) {
            const startNode = this.getEdgePoint(sourceGeom.x, sourceGeom.y, sourceGeom.w, sourceGeom.h, targetGeom.x, targetGeom.y);
            const endNode = this.getEdgePoint(targetGeom.x, targetGeom.y, targetGeom.w, targetGeom.h, sourceGeom.x, sourceGeom.y);
            beginX = startNode.x;
            beginY = startNode.y;
            endX = endNode.x;
            endY = endNode.y;
        }

        const dx = endX - beginX;
        const dy = endY - beginY;
        const width = Math.sqrt(dx * dx + dy * dy);
        const angle = Math.atan2(dy, dx);

        return { beginX, beginY, endX, endY, width, angle };
    }

    private createConnectorShapeObject(id: string, layout: any, beginArrow?: string, endArrow?: string) {
        const { beginX, beginY, endX, endY, width, angle } = layout;

        return {
            '@_ID': id,
            '@_NameU': 'Dynamic connector',
            '@_Name': 'Dynamic connector',
            '@_Type': 'Shape',
            Cell: [
                { '@_N': 'BeginX', '@_V': beginX.toString() },
                { '@_N': 'BeginY', '@_V': beginY.toString() },
                { '@_N': 'EndX', '@_V': endX.toString() },
                { '@_N': 'EndY', '@_V': endY.toString() },
                { '@_N': 'PinX', '@_V': ((beginX + endX) / 2).toString(), '@_F': '(BeginX+EndX)/2' },
                { '@_N': 'PinY', '@_V': ((beginY + endY) / 2).toString(), '@_F': '(BeginY+EndY)/2' },
                { '@_N': 'Width', '@_V': width.toString(), '@_F': 'SQRT((EndX-BeginX)^2+(EndY-BeginY)^2)' },
                { '@_N': 'Height', '@_V': '0' },
                { '@_N': 'Angle', '@_V': angle.toString(), '@_F': 'ATAN2(EndY-BeginY,EndX-BeginX)' },
                { '@_N': 'LocPinX', '@_V': (width * 0.5).toString(), '@_F': 'Width*0.5' },
                { '@_N': 'LocPinY', '@_V': '0', '@_F': 'Height*0.5' },
                { '@_N': 'ObjType', '@_V': '2' },
                { '@_N': 'ShapePermeableX', '@_V': '0' },
                { '@_N': 'ShapePermeableY', '@_V': '0' },
                { '@_N': 'ShapeRouteStyle', '@_V': '1' },
                { '@_N': 'ConFixedCode', '@_V': '0' }
            ],
            Section: [
                createLineSection({
                    color: '#000000',
                    weight: '0.01',
                    beginArrow: beginArrow || '0',
                    beginArrowSize: '2',
                    endArrow: endArrow || '0',
                    endArrowSize: '2'
                }),
                {
                    '@_N': 'Geometry',
                    '@_IX': '0',
                    Row: [
                        { '@_T': 'MoveTo', '@_IX': '1', Cell: [{ '@_N': 'X', '@_V': '0' }, { '@_N': 'Y', '@_V': '0' }] },
                        { '@_T': 'LineTo', '@_IX': '2', Cell: [{ '@_N': 'X', '@_V': width.toString(), '@_F': 'Width' }, { '@_N': 'Y', '@_V': '0', '@_F': 'Height*0' }] }
                    ]
                }
            ]
        };
    }

    private addConnectorToConnects(parsed: any, connectorId: string, fromShapeId: string, toShapeId: string) {
        if (!parsed.PageContents.Connects) {
            parsed.PageContents.Connects = { Connect: [] };
        }

        let connectCollection = parsed.PageContents.Connects.Connect;
        // Ensure it's an array if it was a single object or undefined
        if (!Array.isArray(connectCollection)) {
            connectCollection = connectCollection ? [connectCollection] : [];
            parsed.PageContents.Connects.Connect = connectCollection;
        }

        connectCollection.push({
            '@_FromSheet': connectorId,
            '@_FromCell': 'BeginX',
            '@_FromPart': '9',
            '@_ToSheet': fromShapeId,
            '@_ToCell': 'PinX',
            '@_ToPart': '3'
        });

        connectCollection.push({
            '@_FromSheet': connectorId,
            '@_FromCell': 'EndX',
            '@_FromPart': '12',
            '@_ToSheet': toShapeId,
            '@_ToCell': 'PinX',
            '@_ToPart': '3'
        });
    }

    async addConnector(pageId: string, fromShapeId: string, toShapeId: string, beginArrow?: string, endArrow?: string): Promise<string> {
        const pagePath = `visio/pages/page${pageId}.xml`;
        let content = '';

        try {
            content = this.pkg.getFileText(pagePath);
        } catch {
            throw new Error(`Could not find page file for ID ${pageId}. Expected at ${pagePath}`);
        }

        const parsed = this.parser.parse(content);

        // Ensure Shapes collection exists
        if (!parsed.PageContents.Shapes) {
            parsed.PageContents.Shapes = { Shape: [] };
        }
        if (!Array.isArray(parsed.PageContents.Shapes.Shape)) {
            parsed.PageContents.Shapes.Shape = [parsed.PageContents.Shapes.Shape];
        }

        const newId = this.getNextId(parsed);
        const shapeHierarchy = this.buildShapeHierarchy(parsed);

        const layout = this.calculateConnectorLayout(fromShapeId, toShapeId, shapeHierarchy);
        const connectorShape = this.createConnectorShapeObject(newId, layout, beginArrow, endArrow);

        const topLevelShapes = parsed.PageContents.Shapes.Shape;
        topLevelShapes.push(connectorShape);

        this.addConnectorToConnects(parsed, newId, fromShapeId, toShapeId);

        // Save back
        const newXml = this.builder.build(parsed);
        this.pkg.updateFile(pagePath, newXml);

        return newId;
    }

    async addShape(pageId: string, props: NewShapeProps, parentId?: string): Promise<string> {
        const pagePath = this.getPagePath(pageId);
        let content: string;
        try {
            content = this.pkg.getFileText(pagePath);
        } catch {
            throw new Error(`Could not find page file for ID ${pageId}. Expected at ${pagePath}`);
        }

        const parsed = this.parser.parse(content);

        // Ensure Shapes container exists
        if (!parsed.PageContents.Shapes) {
            parsed.PageContents.Shapes = { Shape: [] };
        }
        let topLevelShapes = parsed.PageContents.Shapes.Shape;
        if (!Array.isArray(topLevelShapes)) {
            topLevelShapes = topLevelShapes ? [topLevelShapes] : [];
            parsed.PageContents.Shapes.Shape = topLevelShapes;
        }

        const allShapes = this.getAllShapes(parsed);

        // Auto-generate ID if not provided
        let newId = props.id;
        if (!newId) {
            newId = this.getNextId(parsed);
        }

        const newShape: any = {
            '@_ID': newId,
            '@_Name': `Sheet.${newId}`,
            '@_Type': props.type || 'Shape', // Allow specifying Group type
            Cell: [
                { '@_N': 'PinX', '@_V': props.x.toString() },
                { '@_N': 'PinY', '@_V': props.y.toString() },
                { '@_N': 'Width', '@_V': props.width.toString() },
                { '@_N': 'Height', '@_V': props.height.toString() },
                { '@_N': 'LocPinX', '@_V': (props.width / 2).toString() },
                { '@_N': 'LocPinY', '@_V': (props.height / 2).toString() }
            ],
            Text: { '#text': props.text },
            Section: []
        };

        if (props.masterId) {
            newShape['@_Master'] = props.masterId;

            // Phase 3: Ensure Relationship
            // We assume the Page needs a link to the central Masters part to resolve IDs
            // Target path is relative to the *package root* mostly, but in .rels it's relative to the page folder?
            // "visio/pages/_rels/page1.xml.rels" -> Target="../masters/masters.xml"
            // Standard Visio relationship to Masters:
            await this.relsManager.ensureRelationship(
                `visio/pages/page${pageId}.xml`,
                '../masters/masters.xml',
                'http://schemas.microsoft.com/visio/2010/relationships/masters'
            );
        }

        // Only add Geometry if NOT a Group AND NOT a Master Instance
        // Groups should be pure containers for the Table parts (Header/Body)
        // Master instances inherit geometry from the Stencil
        if (props.type !== 'Group' && !props.masterId) {
            newShape.Section.push({
                '@_N': 'Geometry',
                '@_IX': '0',
                Row: [
                    { '@_T': 'MoveTo', '@_IX': '1', Cell: [{ '@_N': 'X', '@_V': '0' }, { '@_N': 'Y', '@_V': '0' }] },
                    { '@_T': 'LineTo', '@_IX': '2', Cell: [{ '@_N': 'X', '@_V': props.width.toString() }, { '@_N': 'Y', '@_V': '0' }] },
                    { '@_T': 'LineTo', '@_IX': '3', Cell: [{ '@_N': 'X', '@_V': props.width.toString() }, { '@_N': 'Y', '@_V': props.height.toString() }] },
                    { '@_T': 'LineTo', '@_IX': '4', Cell: [{ '@_N': 'X', '@_V': '0' }, { '@_N': 'Y', '@_V': props.height.toString() }] },
                    { '@_T': 'LineTo', '@_IX': '5', Cell: [{ '@_N': 'X', '@_V': '0' }, { '@_N': 'Y', '@_V': '0' }] }
                ]
            });
        }

        if (props.fillColor) {
            // Add Fill Section
            newShape.Section.push(createFillSection(props.fillColor));
        }
        // Removed: Explicit NoFill for Group is valid, but removing Geometry is cleaner.

        if (props.fontColor || props.bold) {
            newShape.Section.push(createCharacterSection({
                bold: props.bold,
                color: props.fontColor
            }));
        }

        if (parentId) {
            // Add to Parent Group
            const parent = allShapes.find((s: any) => s['@_ID'] == parentId);
            if (!parent) {
                throw new Error(`Parent shape ${parentId} not found`);
            }

            // Ensure Parent has Shapes collection
            if (!parent.Shapes) {
                parent.Shapes = { Shape: [] };
            }
            if (!Array.isArray(parent.Shapes.Shape)) {
                parent.Shapes.Shape = parent.Shapes.Shape ? [parent.Shapes.Shape] : [];
            }

            // Mark parent as Group if not already
            if (parent['@_Type'] !== 'Group') {
                parent['@_Type'] = 'Group';
            }

            parent.Shapes.Shape.push(newShape);
        } else {
            // Add to Page
            topLevelShapes.push(newShape);
        }

        const newXml = this.builder.build(parsed);
        this.pkg.updateFile(pagePath, newXml);

        return newId;
    }

    async updateShapeText(pageId: string, shapeId: string, newText: string): Promise<void> {
        const pagePath = this.getPagePath(pageId);
        let content: string;

        try {
            content = this.pkg.getFileText(pagePath);
        } catch {
            // Fallback: This is a simplification.
            // In a real robust app, we should use the page mapping logic to find file path.
            // For now, if exact ID match fails, we throw.
            // A more complex lookup would require re-accessing PageManager or iterating keys.
            throw new Error(`Could not find page file for ID ${pageId}. Expected at ${pagePath}`);
        }

        const parsed = this.parser.parse(content);
        let found = false;

        // Helper to recursively find and update shape
        const findAndUpdate = (shapes: any[]) => {
            for (const shape of shapes) {
                if (shape['@_ID'] == shapeId) {
                    shape.Text = {
                        '#text': newText
                    };
                    found = true;
                    return;
                }
                // If shapes can be nested (Groups), check for sub-shapes - future improvement
                // But typically basic Text is on top level or group level.
            }
        };

        const shapesData = parsed.PageContents?.Shapes?.Shape;
        if (shapesData) {
            const shapesArray = Array.isArray(shapesData) ? shapesData : [shapesData];
            findAndUpdate(shapesArray);
        }

        if (!found) {
            throw new Error(`Shape ${shapeId} not found on page ${pageId}`);
        }

        const newXml = this.builder.build(parsed);
        this.pkg.updateFile(pagePath, newXml);
    }
    async updateShapeStyle(pageId: string, shapeId: string, style: ShapeStyle): Promise<void> {
        const pagePath = this.getPagePath(pageId);
        let content: string;
        try {
            content = this.pkg.getFileText(pagePath);
        } catch {
            throw new Error(`Could not find page file for ID ${pageId}`);
        }

        const parsed = this.parser.parse(content);
        let found = false;

        const findAndUpdate = (shapes: any[]) => {
            for (const shape of shapes) {
                if (shape['@_ID'] == shapeId) {
                    found = true;
                    // Ensure Section array exists
                    if (!shape.Section) {
                        shape.Section = [];
                    } else if (!Array.isArray(shape.Section)) {
                        shape.Section = [shape.Section];
                    }

                    // Update/Add Fill
                    if (style.fillColor) {
                        // Remove existing Fill section if any (simplified: assuming IX=0)
                        shape.Section = shape.Section.filter((s: any) => s['@_N'] !== 'Fill');
                        shape.Section.push(createFillSection(style.fillColor));
                    }

                    // Update/Add Character (Font/Text Style)
                    if (style.fontColor || style.bold !== undefined) {
                        // Remove existing Character section if any
                        shape.Section = shape.Section.filter((s: any) => s['@_N'] !== 'Character');
                        shape.Section.push(createCharacterSection({
                            bold: style.bold,
                            color: style.fontColor
                        }));
                    }
                    return;
                }
            }
        };

        const shapesData = parsed.PageContents?.Shapes?.Shape;
        if (shapesData) {
            const shapesArray = Array.isArray(shapesData) ? shapesData : [shapesData];
            findAndUpdate(shapesArray);
        }

        if (!found) {
            throw new Error(`Shape ${shapeId} not found on page ${pageId}`);
        }

        const newXml = this.builder.build(parsed);
        this.pkg.updateFile(pagePath, newXml);
    }
    async updateShapePosition(pageId: string, shapeId: string, x: number, y: number): Promise<void> {
        const pagePath = this.getPagePath(pageId);
        let content: string;
        try {
            content = this.pkg.getFileText(pagePath);
        } catch {
            throw new Error(`Could not find page file for ID ${pageId}`);
        }

        const parsed = this.parser.parse(content);
        let found = false;

        const findAndUpdate = (shapes: any[]) => {
            for (const shape of shapes) {
                if (shape['@_ID'] == shapeId) {
                    found = true;
                    // Ensure Cell array exists
                    if (!shape.Cell) {
                        shape.Cell = [];
                    } else if (!Array.isArray(shape.Cell)) {
                        shape.Cell = [shape.Cell];
                    }

                    // Helper to update specific cell
                    const updateCell = (name: string, value: string) => {
                        const cell = shape.Cell.find((c: any) => c['@_N'] === name);
                        if (cell) {
                            cell['@_V'] = value;
                        } else {
                            shape.Cell.push({ '@_N': name, '@_V': value });
                        }
                    };

                    updateCell('PinX', x.toString());
                    updateCell('PinY', y.toString());
                    return;
                }
            }
        };

        const shapesData = parsed.PageContents?.Shapes?.Shape;
        if (shapesData) {
            const shapesArray = Array.isArray(shapesData) ? shapesData : [shapesData];
            findAndUpdate(shapesArray);
        }

        if (!found) {
            throw new Error(`Shape ${shapeId} not found on page ${pageId}`);
        }

        const newXml = this.builder.build(parsed);
        this.pkg.updateFile(pagePath, newXml);
    }
}

export interface ShapeStyle {
    fillColor?: string;
    fontColor?: string;
    bold?: boolean;
}
