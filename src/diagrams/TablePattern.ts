import type { Page } from '../Page';
import type { Shape } from '../Shape';

/**
 * High-level table composition: a group shape containing a shaded header row
 * and a body row whose text lists the column names.
 *
 * Used by `page.addTable()` and `SchemaDiagram.addTable()`.
 */
export class TablePattern {
    static async add(page: Page, x: number, y: number, title: string, columns: string[]): Promise<Shape> {
        const width = 3;
        const headerHeight = 0.5;
        const lineItemHeight = 0.25;
        const bodyHeight = Math.max(0.5, columns.length * lineItemHeight + 0.1);
        const totalHeight = headerHeight + bodyHeight;

        // Group contains a header row and a body row; child coords are relative to the group origin.
        const groupShape = await page.addShape({
            text: '',
            x, y, width, height: totalHeight,
            type: 'Group',
        });

        const headerCenterY = bodyHeight + (headerHeight / 2);
        await page.addShape({
            text: title,
            x: width / 2,
            y: headerCenterY,
            width, height: headerHeight,
            fillColor: '#DDDDDD',
            bold: true,
        }, groupShape.id);

        const bodyCenterY = bodyHeight / 2;
        await page.addShape({
            text: columns.join('\n'),
            x: width / 2,
            y: bodyCenterY,
            width, height: bodyHeight,
            fillColor: '#FFFFFF',
            fontColor: '#000000',
        }, groupShape.id);

        return groupShape;
    }
}
