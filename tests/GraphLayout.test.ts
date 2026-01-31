import { describe, it, expect } from 'vitest';
import { GraphLayout } from '../src/layout/GraphLayout';
import { GraphData } from '../src/mermaid/MermaidParser';

describe('GraphLayout', () => {
    const layout = new GraphLayout();

    it('should assign coordinates to nodes', () => {
        const input: GraphData = {
            nodes: [
                { id: 'A', text: 'Node A', type: 'square' },
                { id: 'B', text: 'Node B', type: 'square' }
            ],
            edges: [
                { from: 'A', to: 'B', type: 'arrow' }
            ]
        };

        const result = layout.calculateLayout(input);

        expect(result).toHaveLength(2);
        const nodeA = result.find(n => n.id === 'A');
        const nodeB = result.find(n => n.id === 'B');

        expect(nodeA).toBeDefined();
        expect(nodeB).toBeDefined();

        expect(typeof nodeA?.x).toBe('number');
        expect(typeof nodeA?.y).toBe('number');
        expect(typeof nodeB?.x).toBe('number');
        expect(typeof nodeB?.y).toBe('number');

        // Since it's Top-Bottom, A should be above B (smaller Y)
        expect(nodeA!.y).toBeLessThan(nodeB!.y);
    });

    it('should respect manual scale/units if provided', () => {
        const input: GraphData = {
            nodes: [{ id: 'A', text: 'A', type: 'square' }],
            edges: []
        };

        const result = layout.calculateLayout(input, { nodeWidth: 100, nodeHeight: 50 });
        expect(result[0].width).toBe(100);
        expect(result[0].height).toBe(50);
    });
});
