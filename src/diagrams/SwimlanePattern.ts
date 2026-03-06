import type { Page } from '../Page';
import type { Shape } from '../Shape';
import type { NewShapeProps } from '../types/VisioTypes';

/**
 * Swimlane composition helpers.
 *
 * A swimlane diagram is modelled as a vertical List (the pool) containing
 * one or more Containers (the lanes). Used by `page.addSwimlanePool()` and
 * `page.addSwimlaneLane()`.
 */
export class SwimlanePattern {
    static async addPool(page: Page, props: NewShapeProps): Promise<Shape> {
        return page.addList(props, 'vertical');
    }

    static async addLane(page: Page, pool: Shape, props: NewShapeProps): Promise<Shape> {
        const lane = await page.addContainer(props);
        await pool.addListItem(lane);
        return lane;
    }
}
