export interface VisioSection {
    '@_N': string;
    '@_IX'?: string;
    Row?: any[];
    Cell?: any[];
}

export function createFillSection(hexColor: string): VisioSection {
    // Visio uses FillForegnd for the main background color.
    // Ideally we should sanitize hexColor to be #RRGGBB.
    return {
        '@_N': 'Fill',
        '@_IX': '0', // Standard index for unique sections
        Cell: [
            { '@_N': 'FillForegnd', '@_V': hexColor },
            { '@_N': 'FillBkgnd', '@_V': '#FFFFFF' }, // Default background pattern color usually white
            { '@_N': 'FillPattern', '@_V': '1' }      // 1 = Solid fill
        ]
    };
}
