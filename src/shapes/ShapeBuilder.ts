
import { NewShapeProps } from '../types/VisioTypes';
import { createFillSection, createCharacterSection, createLineSection } from '../utils/StyleHelpers';

export class ShapeBuilder {
    static createStandardShape(id: string, props: NewShapeProps): any {
        // Validate dimensions
        if (props.width <= 0 || props.height <= 0) {
            throw new Error('Shape dimensions must be positive numbers');
        }

        const shape: any = {
            '@_ID': id,
            '@_NameU': `Sheet.${id}`,
            '@_Name': `Sheet.${id}`,
            '@_Type': props.type || 'Shape',
            Cell: [
                { '@_N': 'PinX', '@_V': props.x.toString() },
                { '@_N': 'PinY', '@_V': props.y.toString() },
                { '@_N': 'Width', '@_V': props.width.toString() },
                { '@_N': 'Height', '@_V': props.height.toString() },
                { '@_N': 'LocPinX', '@_V': (props.width / 2).toString() },
                { '@_N': 'LocPinY', '@_V': (props.height / 2).toString() }
            ],
            Section: []
            // Text added at end by caller or we can do it here if props.text is final
        };

        if (props.masterId) {
            shape['@_Master'] = props.masterId;
        }

        // Add Styles
        if (props.fillColor) {
            shape.Section.push(createFillSection(props.fillColor));

            // Standard Line for fillable shapes
            shape.Section.push(createLineSection({
                color: props.lineColor || '#000000',
                weight: props.lineWeight || '0.01',
                pattern: props.linePattern || '1'
            }));
        }

        if (props.fontColor || props.bold) {
            shape.Section.push(createCharacterSection({
                bold: props.bold,
                color: props.fontColor
            }));
        }

        // Add Geometry
        // Only if NOT a Group AND NOT a Master Instance
        if (props.type !== 'Group' && !props.masterId) {
            shape.Section.push({
                '@_N': 'Geometry',
                '@_IX': '0',
                Cell: [{ '@_N': 'NoFill', '@_V': props.fillColor ? '0' : '1' }],
                Row: [
                    { '@_T': 'MoveTo', '@_IX': '1', Cell: [{ '@_N': 'X', '@_V': '0' }, { '@_N': 'Y', '@_V': '0' }] },
                    { '@_T': 'LineTo', '@_IX': '2', Cell: [{ '@_N': 'X', '@_V': props.width.toString() }, { '@_N': 'Y', '@_V': '0' }] },
                    { '@_T': 'LineTo', '@_IX': '3', Cell: [{ '@_N': 'X', '@_V': props.width.toString() }, { '@_N': 'Y', '@_V': props.height.toString() }] },
                    { '@_T': 'LineTo', '@_IX': '4', Cell: [{ '@_N': 'X', '@_V': '0' }, { '@_N': 'Y', '@_V': props.height.toString() }] },
                    { '@_T': 'LineTo', '@_IX': '5', Cell: [{ '@_N': 'X', '@_V': '0' }, { '@_N': 'Y', '@_V': '0' }] }
                ]
            });
        }

        // Handle Text if provided
        if (props.text !== undefined && props.text !== null) {
            shape.Text = { '#text': props.text };
        }

        return shape;
    }
}
