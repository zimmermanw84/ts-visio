import { XMLParser } from 'fast-xml-parser';
import { VisioPackage } from './VisioPackage';
import { VisioShape } from './types/VisioTypes';
import { asArray, parseCells, parseSection } from './utils/VisioParsers';

export class ShapeReader {
    private parser: XMLParser;

    constructor(private pkg: VisioPackage) {
        this.parser = new XMLParser({
            ignoreAttributes: false,
            attributeNamePrefix: "@_"
        });
    }

    readShapes(path: string): VisioShape[] {
        let content: string;
        try {
            content = this.pkg.getFileText(path);
        } catch {
            return [];
        }

        const parsed = this.parser.parse(content);
        // Supports PageContents -> Shapes -> Shape or just Shapes -> Shape depending on structure
        const shapesData = parsed.PageContents?.Shapes?.Shape;

        if (!shapesData) return [];

        const shapesArray = asArray<any>(shapesData);
        return shapesArray.map(s => this.parseShape(s));
    }

    private parseShape(s: any): VisioShape {
        const shape: VisioShape = {
            ID: s['@_ID'],
            Name: s['@_Name'],
            NameU: s['@_NameU'],
            Type: s['@_Type'],
            Master: s['@_Master'],
            Text: s.Text?.['#text'] || (typeof s.Text === 'string' ? s.Text : undefined),
            Cells: parseCells(s),
            Sections: {}
        };

        const sections = asArray(s.Section);
        for (const sec of sections) {
            const section = sec as any;
            if (section['@_N']) {
                shape.Sections[section['@_N']] = parseSection(section);
            }
        }

        return shape;
    }
}
