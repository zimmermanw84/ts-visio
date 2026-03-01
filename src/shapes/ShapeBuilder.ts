
import { NewShapeProps } from '../types/VisioTypes';
import { createFillSection, createCharacterSection, createLineSection, createParagraphSection, vertAlignValue } from '../utils/StyleHelpers';
import { GeometryBuilder } from './GeometryBuilder';

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
                { '@_N': 'LocPinX', '@_V': (props.width / 2).toString(), '@_F': 'Width*0.5' },
                { '@_N': 'LocPinY', '@_V': (props.height / 2).toString(), '@_F': 'Height*0.5' }
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
                weight: '0.01'
            }));
        }

        if (props.fontColor || props.bold || props.fontSize !== undefined || props.fontFamily !== undefined) {
            shape.Section.push(createCharacterSection({
                bold: props.bold,
                color: props.fontColor,
                fontSize: props.fontSize,
                fontFamily: props.fontFamily,
            }));
        }

        if (props.horzAlign !== undefined) {
            shape.Section.push(createParagraphSection(props.horzAlign));
        }

        if (props.verticalAlign !== undefined) {
            (shape.Cell as any[]).push({ '@_N': 'VerticalAlign', '@_V': vertAlignValue(props.verticalAlign) });
        }

        // Add Geometry
        // Only if NOT a Group AND NOT a Master Instance
        if (props.type !== 'Group' && !props.masterId) {
            shape.Section.push(GeometryBuilder.build(props));
        }

        // Handle Text if provided
        if (props.text !== undefined && props.text !== null) {
            shape.Text = { '#text': props.text };
        }

        return shape;
    }
}
