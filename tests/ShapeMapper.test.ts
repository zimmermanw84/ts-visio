import { describe, it, expect } from 'vitest';
import { ShapeMapper } from '../src/mermaid/ShapeMapper';
import { GraphNode } from '../src/mermaid/MermaidParser';

describe('ShapeMapper', () => {
    it('should map square to Process', () => {
        const node: GraphNode = { id: 'A', text: 'Task', type: 'square' };
        const props = ShapeMapper.getShapeProps(node);
        expect(props.masterId).toBe('Process');
    });

    it('should map round to Terminator', () => {
        const node: GraphNode = { id: 'A', text: 'End', type: 'round' };
        const props = ShapeMapper.getShapeProps(node);
        expect(props.masterId).toBe('Terminator');
    });

    it('should map rhombus to Decision', () => {
        const node: GraphNode = { id: 'A', text: 'If', type: 'rhombus' };
        const props = ShapeMapper.getShapeProps(node);
        expect(props.masterId).toBe('Decision');
    });

    it('should map cylinder to Database', () => {
        const node: GraphNode = { id: 'A', text: 'DB', type: 'cylinder' };
        const props = ShapeMapper.getShapeProps(node);
        expect(props.masterId).toBe('Database');
    });

    it('should map parallelogram to Data', () => {
        const node: GraphNode = { id: 'A', text: 'Input', type: 'parallelogram' };
        const props = ShapeMapper.getShapeProps(node);
        expect(props.masterId).toBe('Data');
    });

    it('should apply styles', () => {
        const node: GraphNode = {
            id: 'A',
            text: 'Styled',
            type: 'square',
            style: {
                fill: '#ff0000',
                stroke: '#0000ff',
                'stroke-width': '2px',
                'stroke-dasharray': '5 5'
            }
        };
        const props = ShapeMapper.getShapeProps(node);
        expect(props.fillColor).toBe('#ff0000');
        expect(props.lineColor).toBe('#0000ff');
        expect(props.lineWeight).toBe('2px');
        expect(props.linePattern).toBe('2');
    });
});
