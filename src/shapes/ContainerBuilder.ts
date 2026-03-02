import { NewShapeProps } from '../types/VisioTypes';
import { ShapeBuilder } from './ShapeBuilder';
import { createLineSection } from '../utils/StyleHelpers';

export class ContainerBuilder {
    static createContainerShape(id: string, props: NewShapeProps): any {
        const shape = ShapeBuilder.createStandardShape(id, props);
        this.makeContainer(shape);
        return shape;
    }

    static makeContainer(shape: any) {
        if (!shape.Section) shape.Section = [];
        if (!Array.isArray(shape.Section)) shape.Section = [shape.Section];

        let userSection = shape.Section.find((s: any) => s['@_N'] === 'User');
        if (!userSection) {
            userSection = { '@_N': 'User', Row: [] };
            shape.Section.push(userSection);
        }

        if (!userSection.Row) userSection.Row = [];
        if (!Array.isArray(userSection.Row)) userSection.Row = [userSection.Row];

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

        addUserRow('msvStructureType', '"Container"');
        addUserRow('msvSDContainerMargin', '10 mm');

        let textXform = shape.Section.find((s: any) => s['@_N'] === 'TextXform');
        if (!textXform) {
            textXform = { '@_N': 'TextXform', Cell: [] };
            shape.Section.push(textXform);
        }

        if (!textXform.Cell) textXform.Cell = [];
        if (!Array.isArray(textXform.Cell)) textXform.Cell = [textXform.Cell];

        const heightVal = (shape.Cell as any[])?.find((c: any) => c['@_N'] === 'Height')?.['@_V'] ?? '1';

        const upsertCell = (name: string, formula: string, unit: string, val: string) => {
            const idx = textXform.Cell.findIndex((c: any) => c['@_N'] === name);
            const cell = { '@_N': name, '@_V': val, '@_F': formula, '@_U': unit };
            if (idx >= 0) textXform.Cell[idx] = cell;
            else textXform.Cell.push(cell);
        };

        // Move text to top of shape: formula drives dynamic sizing; @_V is the static initial value.
        // TxtHeight is Visio-computed, so its static value is left as '0'.
        upsertCell('TxtPinY',    'Height',    'DY', heightVal);
        upsertCell('TxtLocPinY', 'TxtHeight', 'DY', '0');
    }

    static makeList(shape: any, direction: 'vertical' | 'horizontal' = 'vertical') {
        this.makeContainer(shape);

        let userSection = shape.Section.find((s: any) => s['@_N'] === 'User');

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

