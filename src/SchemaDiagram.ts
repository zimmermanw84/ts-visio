import { Page } from './Page';
import { Shape } from './Shape';

export type RelationType = '1:1' | '1:N';

export class SchemaDiagram {
    constructor(private page: Page) { }

    /**
     * Adds a table entity to the diagram.
     * @param tableName Name of the table
     * @param columns List of column names (e.g., "id: int")
     * @param x X coordinate
     * @param y Y coordinate
     */
    async addTable(tableName: string, columns: string[], x: number = 0, y: number = 0): Promise<Shape> {
        return this.page.addTable(x, y, tableName, columns);
    }

    /**
     * Connects two tables with a specific relationship type.
     * @param fromTable Source table shape
     * @param toTable Target table shape
     * @param type Relationship type ('1:1' or '1:N')
     */
    async addRelation(fromTable: Shape, toTable: Shape, type: RelationType): Promise<void> {
        let beginArrow = '0'; // No arrow at start
        let endArrow = '1';   // Default standard arrow

        if (type === '1:1') {
            endArrow = '1'; // Standard Arrow
        } else if (type === '1:N') {
            endArrow = '29'; // Crow's Foot
        }

        await this.page.connectShapes(fromTable, toTable, beginArrow, endArrow);
    }
}
