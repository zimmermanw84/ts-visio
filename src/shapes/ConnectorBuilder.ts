
import { createLineSection } from '../utils/StyleHelpers';

export class ConnectorBuilder {
    private static getCellVal(shape: any, name: string): string {
        if (!shape || !shape.Cell) return '0';
        const cells = Array.isArray(shape.Cell) ? shape.Cell : [shape.Cell];
        const cell = cells.find((c: any) => c['@_N'] === name);
        return cell ? cell['@_V'] : '0';
    }

    private static getAbsolutePos(id: string, shapeHierarchy: Map<string, { shape: any; parent: any }>): { x: number, y: number } {
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

    private static getEdgePoint(cx: number, cy: number, w: number, h: number, targetX: number, targetY: number): { x: number, y: number } {
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

    static buildShapeHierarchy(parsed: any): Map<string, { shape: any; parent: any }> {
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

    static calculateConnectorLayout(
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

    static createConnectorShapeObject(id: string, layout: any, beginArrow?: string, endArrow?: string) {
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

    static addConnectorToConnects(parsed: any, connectorId: string, fromShapeId: string, toShapeId: string) {
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
}
