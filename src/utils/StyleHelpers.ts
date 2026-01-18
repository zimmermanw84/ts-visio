export interface VisioSection {
    '@_N': string;
    '@_IX': string;
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

export const ArrowHeads = {
    None: '0',
    Standard: '1',
    Open: '2',
    Stealth: '3',
    Diamond: '4',
    Oneway: '5',
    CrowsFoot: '29', // Visio "Many"
    One: '24',       // Visio "One" (Dash) - Approximate, or '26'
    // There are many variants, but 29 is the standard "Fork"
};

export function createCharacterSection(props: { bold?: boolean; color?: string }): VisioSection {
    // Visio Character Section
    // N="Character"
    // Row T="Character"
    //   Cell N="Color" V="#FF0000"
    //   Cell N="Style" V="1" (1=Bold, 2=Italic, 4=Underline) - Bitwise

    // Visio booleans are often 0 or 1.
    // Style=1 (Bold)

    // Default Style is 0 (Normal)
    let styleVal = 0;
    if (props.bold) {
        styleVal += 1; // Add Bold bit
    }

    // Default Color is usually 0 (Black) or specific hex
    const colorVal = props.color || '#000000';

    return {
        '@_N': 'Character',
        '@_IX': '0',
        Row: [
            {
                '@_T': 'Character',
                '@_IX': '0',
                Cell: [
                    { '@_N': 'Color', '@_V': colorVal },
                    { '@_N': 'Style', '@_V': styleVal.toString() },
                    // Size, Font, etc could go here
                    { '@_N': 'Font', '@_V': '1' } // Default font (Calibri usually)
                ]
            }
        ]
    };
}

export function createLineSection(props: { color?: string; pattern?: string; weight?: string }): VisioSection {
    return {
        '@_N': 'Line',
        '@_IX': '0',
        Cell: [
            { '@_N': 'LineColor', '@_V': props.color || '#000000' },
            { '@_N': 'LinePattern', '@_V': props.pattern || '1' }, // 1 = Solid
            { '@_N': 'LineWeight', '@_V': props.weight || '0.01' } // ~0.72pt
        ]
    };
}
