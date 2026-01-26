import { GraphNode } from './MermaidParser';
import { NewShapeProps } from '../types/VisioTypes';

export class ShapeMapper {
    /**
     * Maps a Mermaid GraphNode to Visio Shape Properties.
     * Determines whether to use a Master or geometry.
     */
    static getShapeProps(node: GraphNode): Partial<NewShapeProps> {
        const props: Partial<NewShapeProps> = {
            text: node.text
        };

        // Master Mappings based on Requirements
        // [text] -> Rectangle -> Master: Process
        // (text) -> Rounded Rectangle -> Master: Terminator (or Start/End)
        // {text} -> Rhombus -> Master: Decision
        // [(text)] -> Cylinder -> Master: Database
        // ((text)) -> Circle -> Master: Start/End (or On-page Reference)
        // [/text/] -> Parallelogram -> Master: Data

        switch (node.type) {
            case 'square':
                props.masterId = 'Process'; // Visio Standard: Process
                break;
            case 'round':
                props.masterId = 'Terminator'; // Visio Standard: Terminator
                break;
            case 'rhombus':
                props.masterId = 'Decision'; // Visio Standard: Decision
                break;
            case 'cylinder':
                props.masterId = 'Database'; // Visio Standard: Database
                break;
            case 'circle':
                props.masterId = 'Start/End'; // Visio Standard: Start/End
                break;
            case 'parallelogram':
                props.masterId = 'Data'; // Visio Standard: Data (Input/Output)
                break;
            default:
                props.masterId = 'Process';
        }

        // Apply Styles
        if (node.style) {
            if (node.style.fill) props.fillColor = node.style.fill;
            if (node.style.stroke) props.lineColor = node.style.stroke;

            if (node.style['stroke-width']) {
                // Parse "2px" or "4px". Visio uses "pt" often but generic string value is safest.
                // 1px approx 0.75pt.
                // If value is just "2px", let's keep it or strip "px" and append "pt" for clarity if purely numeric?
                // Visio accepts "2px" in formulas sometimes if it parses units, but for safety in XML
                // we might want "2 pt" or just pass it through if Visio handles it.
                // Let's pass it through as is, Visio is usually good with units.
                props.lineWeight = node.style['stroke-width'];
            }

            if (node.style['stroke-dasharray']) {
                // Map dasharray to LinePattern.
                // Mermaid: "5 5" etc.
                // Visio: 0=NoLine, 1=Solid, 2=Dash, 3=Dot, 4=DashDot...
                // Simple mapping: if present, assume Dash (2) or try to detect.
                // For now, if any dasharray is set, set Pattern to 2 (Dash).
                props.linePattern = '2';
            }
        }

        return props;
    }
}
