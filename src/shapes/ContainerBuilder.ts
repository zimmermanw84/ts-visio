import { NewShapeProps } from '../types/VisioTypes';
import { ShapeBuilder } from './ShapeBuilder';
import { createLineSection } from '../utils/StyleHelpers';

export class ContainerBuilder {
    static createContainerShape(id: string, props: NewShapeProps): any {
        // Reuse basic shape creation (transform, text, etc)
        const shape = ShapeBuilder.createStandardShape(id, props);

        // Styling Override for "Classic Container":
        // Usually containers have a Header implementation or specific Geometry.
        // For Phase 1, we will stick to a basic Rectangle with the Metadata.
        // If the user wants "classic" look, we might need 2 geometries (Header + Body).

        // 1. Inject Container Metadata (User Cells)
        this.makeContainer(shape);

        // 2. Adjust styling if necessary
        // Containers often have 'NoFill' for the body so you can see behind,
        // or they have a Group structure.
        // For a simple single-shape container, we rely on the Geometry.

        return shape;
    }

    static makeContainer(shape: any) {
        // Ensure Section array
        if (!shape.Section) shape.Section = [];
        if (!Array.isArray(shape.Section)) shape.Section = [shape.Section];

        // 1. User Section (Metadata)
        let userSection = shape.Section.find((s: any) => s['@_N'] === 'User');
        if (!userSection) {
            userSection = { '@_N': 'User', Row: [] };
            shape.Section.push(userSection);
        }

        // Ensure Row is array (Critical for XML parser edge cases)
        if (!userSection.Row) userSection.Row = [];
        if (!Array.isArray(userSection.Row)) userSection.Row = [userSection.Row];

        // Helper to add/update user row
        const addUserRow = (name: string, value: string) => {
            const rowIdx = userSection.Row.findIndex((r: any) => r['@_N'] === name);
            const newRow = {
                '@_N': name,
                Cell: [{ '@_N': 'Value', '@_V': value }]
            };

            if (rowIdx >= 0) {
                userSection.Row[rowIdx] = newRow;
            } else {
                userSection.Row.push(newRow);
            }
        };

        // Critical Metadata
        addUserRow('msvStructureType', '"Container"');
        addUserRow('msvSDContainerMargin', '10 mm');

        // 2. Adjust Text Position (Simulate Header)
        // Find or create TextXform
        let textXform = shape.Section.find((s: any) => s['@_N'] === 'TextXform');
        if (!textXform) {
            textXform = { '@_N': 'TextXform', Cell: [] };
            shape.Section.push(textXform);
        }

        // Ensure Cell is array
        if (!textXform.Cell) textXform.Cell = [];
        if (!Array.isArray(textXform.Cell)) textXform.Cell = [textXform.Cell];

        const upsertCell = (name: string, val: string, unit: string) => {
            const idx = textXform.Cell.findIndex((c: any) => c['@_N'] === name);
            const cell = { '@_N': name, '@_V': val, '@_U': unit };
            if (idx >= 0) textXform.Cell[idx] = cell;
            else textXform.Cell.push(cell);
        };

        // Move text to top of shape: TxtPinY = Height, TxtLocPinY = TxtHeight
        // Removed redundant '*1' from formulas
        upsertCell('TxtPinY', 'Height', 'DY');
        upsertCell('TxtLocPinY', 'TxtHeight', 'DY');
    }

    static makeList(shape: any, direction: 'vertical' | 'horizontal' = 'vertical') {
        // 1. Convert basic container to List
        this.makeContainer(shape);

        // 2. User Section Override
        let userSection = shape.Section.find((s: any) => s['@_N'] === 'User');

        // Helper to add/update user row
        const upsertUserRow = (name: string, value: string) => {
            const rowIdx = userSection.Row.findIndex((r: any) => r['@_N'] === name);
            const newRow = {
                '@_N': name,
                Cell: [{ '@_N': 'Value', '@_V': value }]
            };

            if (rowIdx >= 0) userSection.Row[rowIdx] = newRow;
            else userSection.Row.push(newRow);
        };

        upsertUserRow('msvStructureType', '"List"'); // Override "Container"
        upsertUserRow('msvSDListDirection', direction === 'vertical' ? '1' : '0'); // 1=Vert, 0=Horiz
        upsertUserRow('msvSDListSpacing', '0.125'); // Default spacing (can be param later)
        upsertUserRow('msvSDListAlignment', '1'); // 1=Center
    }
}

