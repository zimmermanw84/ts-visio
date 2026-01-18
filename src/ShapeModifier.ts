import { XMLParser, XMLBuilder } from 'fast-xml-parser';
import { VisioPackage } from './VisioPackage';

export interface NewShapeProps {
    text: string;
    x: number;
    y: number;
    width: number;
    height: number;
    id?: string;
}

export class ShapeModifier {
    private parser: XMLParser;
    private builder: XMLBuilder;

    constructor(private pkg: VisioPackage) {
        this.parser = new XMLParser({
            ignoreAttributes: false,
            attributeNamePrefix: "@_"
        });
        this.builder = new XMLBuilder({
            ignoreAttributes: false,
            attributeNamePrefix: "@_",
            format: true
        });
    }

    private getPagePath(pageId: string): string {
        return `visio/pages/page${pageId}.xml`;
    }

    async addShape(pageId: string, props: NewShapeProps): Promise<string> {
        const pagePath = this.getPagePath(pageId);
        let content: string;
        try {
            content = this.pkg.getFileText(pagePath);
        } catch {
            throw new Error(`Could not find page file for ID ${pageId}. Expected at ${pagePath}`);
        }

        const parsed = this.parser.parse(content);

        // Ensure Shapes container exists
        if (!parsed.PageContents.Shapes) {
            parsed.PageContents.Shapes = { Shape: [] };
        }
        let shapes = parsed.PageContents.Shapes.Shape;
        if (!Array.isArray(shapes)) {
            // If single object or undefined, normalize to array
            shapes = shapes ? [shapes] : [];
            parsed.PageContents.Shapes.Shape = shapes;
        }

        // Auto-generate ID if not provided
        let newId = props.id;
        if (!newId) {
            let maxId = 0;
            for (const s of shapes) {
                const id = parseInt(s['@_ID']);
                if (!isNaN(id) && id > maxId) maxId = id;
            }
            newId = (maxId + 1).toString();
        }

        const newShape = {
            '@_ID': newId,
            '@_Name': `Sheet.${newId}`,
            '@_Type': 'Shape',
            Cell: [
                { '@_N': 'PinX', '@_V': props.x.toString() },
                { '@_N': 'PinY', '@_V': props.y.toString() },
                { '@_N': 'Width', '@_V': props.width.toString() },
                { '@_N': 'Height', '@_V': props.height.toString() },
                { '@_N': 'LocPinX', '@_V': (props.width / 2).toString() },
                { '@_N': 'LocPinY', '@_V': (props.height / 2).toString() }
            ],
            Text: { '#text': props.text },
            Section: [
                {
                    '@_N': 'Geometry',
                    '@_IX': '0',
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

        shapes.push(newShape);

        const newXml = this.builder.build(parsed);
        this.pkg.updateFile(pagePath, newXml);

        return newId;
    }

    async updateShapeText(pageId: string, shapeId: string, newText: string): Promise<void> {
        const pagePath = this.getPagePath(pageId);
        let content: string;

        try {
            content = this.pkg.getFileText(pagePath);
        } catch {
            // Fallback: This is a simplification.
            // In a real robust app, we should use the page mapping logic to find file path.
            // For now, if exact ID match fails, we throw.
            // A more complex lookup would require re-accessing PageManager or iterating keys.
            throw new Error(`Could not find page file for ID ${pageId}. Expected at ${pagePath}`);
        }

        const parsed = this.parser.parse(content);
        let found = false;

        // Helper to recursively find and update shape
        const findAndUpdate = (shapes: any[]) => {
            for (const shape of shapes) {
                if (shape['@_ID'] == shapeId) {
                    shape.Text = {
                        '#text': newText
                    };
                    found = true;
                    return;
                }
                // If shapes can be nested (Groups), check for sub-shapes - future improvement
                // But typically basic Text is on top level or group level.
            }
        };

        const shapesData = parsed.PageContents?.Shapes?.Shape;
        if (shapesData) {
            const shapesArray = Array.isArray(shapesData) ? shapesData : [shapesData];
            findAndUpdate(shapesArray);
        }

        if (!found) {
            throw new Error(`Shape ${shapeId} not found on page ${pageId}`);
        }

        const newXml = this.builder.build(parsed);
        this.pkg.updateFile(pagePath, newXml);
    }
}
