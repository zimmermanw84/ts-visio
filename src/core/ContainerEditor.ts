import { PageXmlCache } from './PageXmlCache';
import { SECTION_NAMES, STRUCT_RELATIONSHIP_TYPES } from './VisioConstants';
import { ContainerBuilder } from '../shapes/ContainerBuilder';
import type { NewShapeProps } from '../types/VisioTypes';

/**
 * Manages containers, lists, and their member relationships.
 */
export class ContainerEditor {
    constructor(private cache: PageXmlCache) {}

    async addContainer(pageId: string, props: NewShapeProps): Promise<string> {
        const parsed = this.cache.getParsed(pageId);

        if (!parsed.PageContents.Shapes) parsed.PageContents.Shapes = { Shape: [] };
        let topLevelShapes = parsed.PageContents.Shapes.Shape;
        if (!Array.isArray(topLevelShapes)) {
            topLevelShapes = topLevelShapes ? [topLevelShapes] : [];
            parsed.PageContents.Shapes.Shape = topLevelShapes;
        }

        const newId = props.id || this.cache.getNextId(parsed);
        const containerShape = ContainerBuilder.createContainerShape(newId, props);

        topLevelShapes.push(containerShape);
        this.cache.getShapeMap(parsed).set(newId, containerShape);

        this.cache.saveParsed(pageId, parsed);
        return newId;
    }

    async addList(
        pageId: string,
        props: NewShapeProps,
        direction: 'vertical' | 'horizontal' = 'vertical',
    ): Promise<string> {
        const parsed = this.cache.getParsed(pageId);

        if (!parsed.PageContents.Shapes) parsed.PageContents.Shapes = { Shape: [] };
        let topLevelShapes = parsed.PageContents.Shapes.Shape;
        if (!Array.isArray(topLevelShapes)) {
            topLevelShapes = topLevelShapes ? [topLevelShapes] : [];
            parsed.PageContents.Shapes.Shape = topLevelShapes;
        }

        const newId = props.id || this.cache.getNextId(parsed);
        const listShape = ContainerBuilder.createContainerShape(newId, props);
        ContainerBuilder.makeList(listShape, direction);

        topLevelShapes.push(listShape);
        this.cache.getShapeMap(parsed).set(newId, listShape);

        this.cache.saveParsed(pageId, parsed);
        return newId;
    }

    async addRelationship(pageId: string, shapeId: string, relatedShapeId: string, type: string): Promise<void> {
        const parsed = this.cache.getParsed(pageId);

        if (!parsed.PageContents.Relationships) {
            parsed.PageContents.Relationships = { Relationship: [] };
        }
        if (!Array.isArray(parsed.PageContents.Relationships.Relationship)) {
            parsed.PageContents.Relationships.Relationship = parsed.PageContents.Relationships.Relationship
                ? [parsed.PageContents.Relationships.Relationship]
                : [];
        }

        const relationships = parsed.PageContents.Relationships.Relationship;
        const exists = relationships.find((r: any) =>
            r['@_Type'] === type &&
            r['@_ShapeID'] === shapeId &&
            r['@_RelatedShapeID'] === relatedShapeId,
        );

        if (!exists) {
            relationships.push({
                '@_Type': type,
                '@_ShapeID': shapeId,
                '@_RelatedShapeID': relatedShapeId,
            });
            this.cache.saveParsed(pageId, parsed);
        }
    }

    getContainerMembers(pageId: string, containerId: string): string[] {
        const parsed = this.cache.getParsed(pageId);
        const rels = parsed.PageContents?.Relationships?.Relationship;
        if (!rels) return [];

        const relsArray = Array.isArray(rels) ? rels : [rels];
        return relsArray
            .filter((r: any) => r['@_Type'] === STRUCT_RELATIONSHIP_TYPES.Container && r['@_ShapeID'] === containerId)
            .map((r: any) => r['@_RelatedShapeID']);
    }

    async reorderShape(pageId: string, shapeId: string, position: 'front' | 'back'): Promise<void> {
        const parsed = this.cache.getParsed(pageId);
        const shapesContainer = parsed.PageContents?.Shapes;
        if (!shapesContainer?.Shape) return;

        let shapes = shapesContainer.Shape;
        if (!Array.isArray(shapes)) shapes = [shapes];

        const idx = shapes.findIndex((s: any) => s['@_ID'] == shapeId);
        if (idx === -1) return;

        const shape = shapes[idx];
        shapes.splice(idx, 1);

        if (position === 'back') {
            shapes.unshift(shape); // Back of Z-Order
        } else {
            shapes.push(shape); // Front of Z-Order
        }

        shapesContainer.Shape = shapes;
        this.cache.saveParsed(pageId, parsed);
    }

    async addListItem(pageId: string, listId: string, itemId: string): Promise<void> {
        const parsed = this.cache.getParsed(pageId);
        const listShape = this.cache.getShapeMap(parsed).get(listId);
        if (!listShape) throw new Error(`List ${listId} not found`);

        const getUserVal = (name: string, def: string) => {
            if (!listShape.Section) return def;
            const userSec = listShape.Section.find((s: any) => s['@_N'] === SECTION_NAMES.User);
            if (!userSec?.Row) return def;
            const rows = Array.isArray(userSec.Row) ? userSec.Row : [userSec.Row];
            const row = rows.find((r: any) => r['@_N'] === name);
            if (!row?.Cell) return def;
            const valCell = Array.isArray(row.Cell) ? row.Cell.find((c: any) => c['@_N'] === 'Value') : row.Cell;
            return valCell ? valCell['@_V'] : def;
        };

        const direction = parseInt(getUserVal('msvSDListDirection', '1')); // 1=Vert, 0=Horiz
        const spacing = parseFloat(getUserVal('msvSDListSpacing', '0.125').replace(/[^0-9.]/g, ''));

        const memberIds = this.getContainerMembers(pageId, listId);
        const itemGeo  = this.cache.getShapeGeometry(pageId, itemId);
        const listGeo  = this.cache.getShapeGeometry(pageId, listId);

        let newX = listGeo.x;
        let newY = listGeo.y;

        if (memberIds.length > 0) {
            const lastId  = memberIds[memberIds.length - 1];
            const lastGeo = this.cache.getShapeGeometry(pageId, lastId);

            if (direction === 1) { // Vertical (Stack Down)
                const lastBottom = lastGeo.y - (lastGeo.height / 2);
                newY = lastBottom - spacing - (itemGeo.height / 2);
                newX = lastGeo.x; // Align Centers
            } else { // Horizontal (Stack Right)
                const lastRight = lastGeo.x + (lastGeo.width / 2);
                newX = lastRight + spacing + (itemGeo.width / 2);
                newY = lastGeo.y; // Align Centers
            }
        }

        await this.cache.updateShapePosition(pageId, itemId, newX, newY);
        await this.addRelationship(pageId, listId, itemId, STRUCT_RELATIONSHIP_TYPES.Container);
        await this.resizeContainerToFit(pageId, listId, 0.25);
    }

    async resizeContainerToFit(pageId: string, containerId: string, padding: number = 0.25): Promise<void> {
        const memberIds = this.getContainerMembers(pageId, containerId);
        if (memberIds.length === 0) return;

        let minX = Infinity,  minY = Infinity;
        let maxX = -Infinity, maxY = -Infinity;

        for (const mid of memberIds) {
            const geo = this.cache.getShapeGeometry(pageId, mid);
            // Visio PinX/PinY is center. Bounding box needs Left/Bottom/Right/Top.
            const left   = geo.x - (geo.width  / 2);
            const right  = geo.x + (geo.width  / 2);
            const bottom = geo.y - (geo.height / 2);
            const top    = geo.y + (geo.height / 2);

            if (left   < minX) minX = left;
            if (right  > maxX) maxX = right;
            if (bottom < minY) minY = bottom;
            if (top    > maxY) maxY = top;
        }

        minX -= padding; maxX += padding;
        minY -= padding; maxY += padding;

        const newWidth  = maxX - minX;
        const newHeight = maxY - minY;
        const newPinX   = minX + (newWidth  / 2);
        const newPinY   = minY + (newHeight / 2);

        await this.cache.updateShapePosition(pageId, containerId, newPinX, newPinY);
        await this.cache.updateShapeDimensions(pageId, containerId, newWidth, newHeight);
        await this.reorderShape(pageId, containerId, 'back');
    }
}
