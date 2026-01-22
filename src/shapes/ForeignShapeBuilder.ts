
import { NewShapeProps } from '../types/VisioTypes';

export class ForeignShapeBuilder {
    static createImageShapeObject(id: string, rId: string, props: NewShapeProps): any {
        return {
            '@_ID': id,
            '@_NameU': `Sheet.${id}`,
            '@_Name': `Sheet.${id}`,
            '@_Type': 'Foreign',
            ForeignData: { '@_r:id': rId },
            Cell: [
                { '@_N': 'PinX', '@_V': props.x.toString() },
                { '@_N': 'PinY', '@_V': props.y.toString() },
                { '@_N': 'Width', '@_V': props.width.toString() },
                { '@_N': 'Height', '@_V': props.height.toString() },
                { '@_N': 'LocPinX', '@_V': (props.width / 2).toString() },
                { '@_N': 'LocPinY', '@_V': (props.height / 2).toString() }
            ],
            Section: [
                // Foreign shapes typically have no border (LinePattern=0)
                {
                    '@_N': 'Line',
                    Cell: [
                        { '@_N': 'LinePattern', '@_V': '0' }, // 0 = Null/No Line
                        { '@_N': 'LineColor', '@_V': '#000000' },
                        { '@_N': 'LineWeight', '@_V': '0' }
                    ]
                },
                // Geometry is required for selection bounds
                {
                    '@_N': 'Geometry',
                    '@_IX': '0',
                    Cell: [{ '@_N': 'NoFill', '@_V': '1' }], // Images usually don't have a fill behind them
                    Row: [
                        { '@_T': 'MoveTo', '@_IX': '1', Cell: [{ '@_N': 'X', '@_V': '0' }, { '@_N': 'Y', '@_V': '0' }] },
                        { '@_T': 'LineTo', '@_IX': '2', Cell: [{ '@_N': 'X', '@_V': props.width.toString() }, { '@_N': 'Y', '@_V': '0' }] },
                        { '@_T': 'LineTo', '@_IX': '3', Cell: [{ '@_N': 'X', '@_V': props.width.toString() }, { '@_N': 'Y', '@_V': props.height.toString() }] },
                        { '@_T': 'LineTo', '@_IX': '4', Cell: [{ '@_N': 'X', '@_V': '0' }, { '@_N': 'Y', '@_V': props.height.toString() }] },
                        { '@_T': 'LineTo', '@_IX': '5', Cell: [{ '@_N': 'X', '@_V': '0' }, { '@_N': 'Y', '@_V': '0' }] }
                    ]
                }
            ]
        };
    }
}
