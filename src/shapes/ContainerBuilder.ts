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
        // Ensure User Section exists
        if (!shape.Section) shape.Section = [];
        let userSection = shape.Section.find((s: any) => s['@_N'] === 'User');
        if (!userSection) {
            userSection = { '@_N': 'User', Row: [] };
            shape.Section.push(userSection);
        }

        // Helper to add/update user row
        const addUserRow = (name: string, value: string, type?: string) => {
            const rowIdx = userSection.Row.findIndex((r: any) => r['@_N'] === name);
            const newRow = {
                '@_N': name,
                Cell: [{ '@_N': 'Value', '@_V': value }]
            };
            // Add Type? typically internal types are fine.

            if (rowIdx >= 0) {
                userSection.Row[rowIdx] = newRow;
            } else {
                userSection.Row.push(newRow);
            }
        };

        // Critical Metadata
        addUserRow('msvStructureType', '"Container"'); // Quotes are important for string values in Type=0?
        // Visio Formula Strings often need quotes.

        addUserRow('msvSDContainerMargin', '10 mm');
        // Default margin

        // Optional: Resize behavior
        // addUserRow('msvSDContainerResize', '0');

        // 2. Adjust Text Position (Simulate Header)
        // Move text to top of shape:
        // TxtPinY = Height
        // TxtLocPinY = TxtHeight (Top of text block matches Top of shape)
        if (!shape.Section.find((s: any) => s['@_N'] === 'TextXform')) {
            shape.Section.push({
                '@_N': 'TextXform',
                Cell: [
                    { '@_N': 'TxtPinY', '@_V': 'Height*1', '@_U': 'DY' }, // Using Formula to link to Height
                    { '@_N': 'TxtLocPinY', '@_V': 'TxtHeight*1', '@_U': 'DY' }
                ]
            });
        }
    }
}
