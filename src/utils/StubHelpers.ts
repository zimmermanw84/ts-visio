
import { VisioShape, VisioCell } from '../types/VisioTypes';

export function createVisioShapeStub(props: {
    ID: string,
    Name?: string,
    Text?: string,
    Cells?: Record<string, string | number>
}): VisioShape {
    return {
        ID: props.ID,
        Name: props.Name || `Sheet.${props.ID}`,
        Type: 'Shape',
        Text: props.Text,
        Cells: Object.entries(props.Cells || {}).reduce((acc, [k, v]) => {
            acc[k] = { N: k, V: v.toString() };
            return acc;
        }, {} as { [name: string]: VisioCell }),
        Sections: {}
    };
}
