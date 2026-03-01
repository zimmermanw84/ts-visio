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
        const shapesData = parsed.PageContents?.Shapes?.Shape;

        if (!shapesData) return [];

        return asArray<any>(shapesData).map(s => this.parseShape(s));
    }

    /**
     * Returns every shape on the page flattened into a single array,
     * including shapes nested inside groups at any depth.
     */
    readAllShapes(path: string): VisioShape[] {
        let content: string;
        try {
            content = this.pkg.getFileText(path);
        } catch {
            return [];
        }

        const parsed = this.parser.parse(content);
        const shapesData = parsed.PageContents?.Shapes?.Shape;
        if (!shapesData) return [];

        const result: VisioShape[] = [];
        this.gatherShapes(asArray<any>(shapesData), result);
        return result;
    }

    /**
     * Find a single shape by ID anywhere in the page tree (including nested groups).
     * Returns undefined if not found.
     */
    readShapeById(path: string, shapeId: string): VisioShape | undefined {
        let content: string;
        try {
            content = this.pkg.getFileText(path);
        } catch {
            return undefined;
        }

        const parsed = this.parser.parse(content);
        const shapesData = parsed.PageContents?.Shapes?.Shape;
        if (!shapesData) return undefined;

        return this.findShapeById(asArray<any>(shapesData), shapeId);
    }

    private gatherShapes(rawShapes: any[], result: VisioShape[]): void {
        for (const s of rawShapes) {
            result.push(this.parseShape(s));
            if (s.Shapes?.Shape) {
                this.gatherShapes(asArray<any>(s.Shapes.Shape), result);
            }
        }
    }

    private findShapeById(rawShapes: any[], shapeId: string): VisioShape | undefined {
        for (const s of rawShapes) {
            if (s['@_ID'] === shapeId) return this.parseShape(s);
            if (s.Shapes?.Shape) {
                const found = this.findShapeById(asArray<any>(s.Shapes.Shape), shapeId);
                if (found) return found;
            }
        }
        return undefined;
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
