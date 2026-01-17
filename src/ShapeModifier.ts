import { XMLParser, XMLBuilder } from 'fast-xml-parser';
import { VisioPackage } from './VisioPackage';

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

    async updateShapeText(pageId: string, shapeId: string, newText: string): Promise<void> {
        // Strategy: Try common naming patterns first, then fallback to iterating
        let pagePath = `visio/pages/page${pageId}.xml`;
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
