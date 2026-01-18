import { VisioCell, VisioSection, VisioRow } from '../types/VisioTypes';

export const asArray = <T>(obj: any): T[] => {
    if (!obj) return [];
    return Array.isArray(obj) ? obj : [obj];
};

export const parseCells = (container: any): { [name: string]: VisioCell } => {
    const cells: { [name: string]: VisioCell } = {};
    const cellData = asArray(container.Cell);

    for (const c of cellData) {
        const cell = c as any;
        if (cell['@_N']) {
            cells[cell['@_N']] = {
                N: cell['@_N'],
                V: cell['@_V'],
                U: cell['@_U'],
                F: cell['@_F']
            };
        }
    }
    return cells;
};

export const parseSection = (sectionData: any): VisioSection => {
    const rows: VisioRow[] = [];
    const rowData = asArray(sectionData.Row);

    for (const r of rowData) {
        const row = r as any;
        rows.push({
            T: row['@_T'],
            IX: row['@_IX'],
            Cells: parseCells(row)
        });
    }

    return {
        N: sectionData['@_N'],
        Rows: rows,
        Cells: parseCells(sectionData) // Attempt to parse direct cells
    };
};
