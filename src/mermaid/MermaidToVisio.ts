import * as fs from 'fs';
import { MermaidParser } from './MermaidParser';
import { GraphLayout } from '../layout/GraphLayout';
import { ShapeMapper } from './ShapeMapper';
import { VisioDocument } from '../VisioDocument';
import { Shape } from '../Shape';
import { NewShapeProps } from '../types/VisioTypes';

export class MermaidToVisio {
    private parser: MermaidParser;
    private layout: GraphLayout;

    constructor() {
        this.parser = new MermaidParser();
        this.layout = new GraphLayout();
    }

    async convert(mermaidText: string, outputPath: string, options?: { pageHeight?: number, pageWidth?: number }): Promise<void> {
        // 1. Parse
        const parsedData = this.parser.parse(mermaidText);

        // 2. Layout
        // Default scale: 1.5 inch width, 1 inch height for spacing calculation
        const positionedNodes = this.layout.calculateLayout(parsedData, {
            nodeWidth: 1.5,
            nodeHeight: 1
        });

        // 3. Initialize Visio
        const doc = await VisioDocument.create();
        const page = doc.pages[0];

        // TODO: Update Page Sheet with Width/Height if provided in options
        const pageHeight = options?.pageHeight || 11; // Default to 11 inches (Letter)

        // Map to keep track of created Visio shapes by Mermaid ID
        const shapeMap = new Map<string, Shape>();

        // 4. Create Shapes
        for (const node of positionedNodes) {
            const mapperProps = ShapeMapper.getShapeProps(node);

            // Calculate Y based on Visio coordinate system (Bottom-Left Origin)
            // Dagre is Top-Left.
            // We assume the drawing starts near top of the page.
            // Visio Y = PageHeight - (DagreY + Margin)
            // Let's assume DagreY is distance from top.
            const visioY = pageHeight - (node.y + (node.height / 2));

            const shapeProps: NewShapeProps = {
                text: mapperProps.text || node.text,
                x: node.x,
                y: visioY,
                width: node.width,
                height: node.height,
                ...mapperProps
            };

            const shape = await page.addShape(shapeProps);
            shapeMap.set(node.id, shape);
        }

        // 5. Connect Shapes
        for (const edge of parsedData.edges) {
            const fromShape = shapeMap.get(edge.from);
            const toShape = shapeMap.get(edge.to);

            if (fromShape && toShape) {
                // Map arrow types?
                // Arrow mapping logic could go to ShapeMapper or here.
                // Defaulting to standard arrow.
                const connector = await page.connectShapes(fromShape, toShape);

                if (edge.text) {
                    await connector.setText(edge.text);
                }
            }
        }

        // 6. Save
        const buffer = await doc.save();
        fs.writeFileSync(outputPath, buffer);
    }
}
