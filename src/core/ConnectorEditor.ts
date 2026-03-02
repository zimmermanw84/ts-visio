import { PageXmlCache } from './PageXmlCache';
import { RelsManager } from './RelsManager';
import { RELATIONSHIP_TYPES } from './VisioConstants';
import { ConnectorBuilder } from '../shapes/ConnectorBuilder';
import type { ConnectorStyle, ConnectionTarget } from '../types/VisioTypes';

/**
 * Creates connector shapes between two shapes on a page.
 */
export class ConnectorEditor {
    constructor(
        private cache: PageXmlCache,
        private relsManager: RelsManager,
    ) {}

    async addConnector(
        pageId: string,
        fromShapeId: string,
        toShapeId: string,
        beginArrow?: string,
        endArrow?: string,
        style?: ConnectorStyle,
        fromPort?: ConnectionTarget,
        toPort?: ConnectionTarget,
    ): Promise<string> {
        const parsed = this.cache.getParsed(pageId);

        if (!parsed.PageContents.Shapes) {
            parsed.PageContents.Shapes = { Shape: [] };
        }
        if (!Array.isArray(parsed.PageContents.Shapes.Shape)) {
            parsed.PageContents.Shapes.Shape = parsed.PageContents.Shapes.Shape
                ? [parsed.PageContents.Shapes.Shape]
                : [];
        }

        const newId = this.cache.getNextId(parsed);
        const shapeHierarchy = ConnectorBuilder.buildShapeHierarchy(parsed);

        // Validate arrow values (Visio supports 0-45)
        const validateArrow = (val?: string): string => {
            if (!val) return '0';
            const num = parseInt(val);
            if (isNaN(num) || num < 0 || num > 45) return '0';
            return val;
        };

        const layout = ConnectorBuilder.calculateConnectorLayout(
            fromShapeId, toShapeId, shapeHierarchy, fromPort, toPort,
        );
        const connectorShape = ConnectorBuilder.createConnectorShapeObject(
            newId, layout, validateArrow(beginArrow), validateArrow(endArrow), style,
        );

        parsed.PageContents.Shapes.Shape.push(connectorShape);
        this.cache.getShapeMap(parsed).set(newId, connectorShape);

        ConnectorBuilder.addConnectorToConnects(
            parsed, newId, fromShapeId, toShapeId, shapeHierarchy, fromPort, toPort,
        );

        await this.relsManager.ensureRelationship(
            `visio/pages/page${pageId}.xml`,
            '../masters/masters.xml',
            RELATIONSHIP_TYPES.MASTERS,
        );

        this.cache.saveParsed(pageId, parsed);
        return newId;
    }
}
