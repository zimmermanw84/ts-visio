
import { NewShapeProps } from '../types/VisioTypes';
import { createFillSection, createCharacterSection, createLineSection, createParagraphSection, createTextBlockSection, vertAlignValue } from '../utils/StyleHelpers';
import { GeometryBuilder } from './GeometryBuilder';
import { ConnectionPointBuilder } from './ConnectionPointBuilder';

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

        // Apply document-level stylesheet references
        if (props.styleId !== undefined) {
            shape['@_LineStyle'] = props.styleId.toString();
            shape['@_FillStyle'] = props.styleId.toString();
            shape['@_TextStyle'] = props.styleId.toString();
        } else {
            if (props.lineStyleId !== undefined) shape['@_LineStyle'] = props.lineStyleId.toString();
            if (props.fillStyleId !== undefined) shape['@_FillStyle'] = props.fillStyleId.toString();
            if (props.textStyleId !== undefined) shape['@_TextStyle'] = props.textStyleId.toString();
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

        const hasCharProps = props.fontColor !== undefined
            || props.bold !== undefined
            || props.italic !== undefined
            || props.underline !== undefined
            || props.strikethrough !== undefined
            || props.fontSize !== undefined
            || props.fontFamily !== undefined;

        if (hasCharProps) {
            shape.Section.push(createCharacterSection({
                bold: props.bold,
                italic: props.italic,
                underline: props.underline,
                strikethrough: props.strikethrough,
                color: props.fontColor,
                fontSize: props.fontSize,
                fontFamily: props.fontFamily,
            }));
        }

        const hasParagraphProps = props.horzAlign !== undefined
            || props.spaceBefore !== undefined
            || props.spaceAfter !== undefined
            || props.lineSpacing !== undefined;

        if (hasParagraphProps) {
            shape.Section.push(createParagraphSection({
                horzAlign: props.horzAlign,
                spaceBefore: props.spaceBefore,
                spaceAfter: props.spaceAfter,
                lineSpacing: props.lineSpacing,
            }));
        }

        const hasTextBlockProps = props.textMarginTop !== undefined
            || props.textMarginBottom !== undefined
            || props.textMarginLeft !== undefined
            || props.textMarginRight !== undefined;

        if (hasTextBlockProps) {
            shape.Section.push(createTextBlockSection({
                topMargin:    props.textMarginTop,
                bottomMargin: props.textMarginBottom,
                leftMargin:   props.textMarginLeft,
                rightMargin:  props.textMarginRight,
            }));
        }

        if (props.verticalAlign !== undefined) {
            (shape.Cell as any[]).push({ '@_N': 'VerticalAlign', '@_V': vertAlignValue(props.verticalAlign) });
        }

        // Connection points
        if (props.connectionPoints && props.connectionPoints.length > 0) {
            shape.Section.push(ConnectionPointBuilder.buildConnectionSection(props.connectionPoints));
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
