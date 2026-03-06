import { ConnectorStyle, ConnectionTarget } from './types/VisioTypes';
import { ShapeModifier } from './ShapeModifier';

/**
 * Raw data extracted from the page XML for a single connector shape.
 * Returned by `ShapeReader.readConnectors()` and wrapped by the `Connector` class.
 */
export interface ConnectorData {
    /** Visio shape ID of the connector (1D shape). */
    id: string;
    /** ID of the shape at the connector's begin-point (BeginX). `undefined` if the begin endpoint is not connected to any shape. */
    fromShapeId: string | undefined;
    /** ID of the shape at the connector's end-point (EndX). `undefined` if the end endpoint is not connected to any shape. */
    toShapeId: string | undefined;
    /** Connection point used on the from-shape. */
    fromPort: ConnectionTarget;
    /** Connection point used on the to-shape. */
    toPort: ConnectionTarget;
    /** Line style extracted from the connector's Line section. */
    style: ConnectorStyle;
    /** Begin-arrow head value (Visio ArrowHeads integer string). */
    beginArrow: string;
    /** End-arrow head value (Visio ArrowHeads integer string). */
    endArrow: string;
}

/**
 * A read/delete wrapper around a connector shape found on a page.
 * Obtain instances via `page.getConnectors()`.
 */
export class Connector {
    /** Visio shape ID of this connector. */
    readonly id: string;
    /** ID of the shape this connector starts from. `undefined` if the begin endpoint is not connected. */
    readonly fromShapeId: string | undefined;
    /** ID of the shape this connector ends at. `undefined` if the end endpoint is not connected. */
    readonly toShapeId: string | undefined;
    /** Connection point used on the from-shape ('center', `{ name }`, or `{ index }`). */
    readonly fromPort: ConnectionTarget;
    /** Connection point used on the to-shape ('center', `{ name }`, or `{ index }`). */
    readonly toPort: ConnectionTarget;
    /** Line style (color, weight, pattern, routing) for this connector. */
    readonly style: ConnectorStyle;
    /** Begin-arrow head value string (e.g. `'0'` = none, `'1'` = standard). */
    readonly beginArrow: string;
    /** End-arrow head value string (e.g. `'0'` = none, `'1'` = standard). */
    readonly endArrow: string;

    constructor(
        data: ConnectorData,
        private readonly pageId: string,
        private readonly modifier: ShapeModifier,
    ) {
        this.id = data.id;
        this.fromShapeId = data.fromShapeId;
        this.toShapeId = data.toShapeId;
        this.fromPort = data.fromPort;
        this.toPort = data.toPort;
        this.style = data.style;
        this.beginArrow = data.beginArrow;
        this.endArrow = data.endArrow;
    }

    /**
     * Delete this connector from the page (removes the shape and its Connect entries).
     */
    async delete(): Promise<void> {
        await this.modifier.deleteShape(this.pageId, this.id);
    }
}
