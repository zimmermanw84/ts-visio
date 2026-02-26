import dagre from 'dagre';
import { GraphData, GraphNode } from '../mermaid/MermaidParser';

export interface LayoutOptions {
    direction: 'TB' | 'BT' | 'LR' | 'RL';
    nodeWidth: number;
    nodeHeight: number;
    rankSep: number;  // Vertical separation
    nodeSep: number;  // Horizontal separation
}

export interface PositionedNode extends GraphNode {
    x: number;
    y: number;
    width: number;
    height: number;
}

export class GraphLayout {
    calculateLayout(graph: GraphData, options: Partial<LayoutOptions> = {}): PositionedNode[] {
        const g = new dagre.graphlib.Graph();

        const opts: LayoutOptions = {
            direction: 'TB',
            nodeWidth: 1, // Default Visio width in inches? Dagre uses pixels usually.
            // Visio pixels? 96 DPI?
            // Let's assume input is inches, but Dagre works in abstract units.
            // If we treat units as inches directly, Dagre might need scaling or we just use small numbers.
            // Visio units "1" = 1 inch often in internal cells if using "IN".
            // Let's use 100 for width, 75 for height as "pixels" and scale later?
            // Or just pass 1, 0.75. Dagre supports float.
            nodeHeight: 0.75,
            rankSep: 0.5,
            nodeSep: 0.5,
            ...options
        };

        g.setGraph({
            rankdir: opts.direction,
            ranksep: opts.rankSep,
            nodesep: opts.nodeSep
        });

        g.setDefaultEdgeLabel(() => ({}));

        // Add Nodes
        graph.nodes.forEach(node => {
            // Future: Calculate specific dimensions based on text length?
            // For now, fixed.
            g.setNode(node.id, { label: node.id, width: opts.nodeWidth, height: opts.nodeHeight });
        });

        // Add Edges
        graph.edges.forEach(edge => {
            g.setEdge(edge.from, edge.to);
        });

        // Calculate Layout
        dagre.layout(g);

        // Map back to PositionedNodes
        const output: PositionedNode[] = [];

        g.nodes().forEach(id => {
            const nodeInfo = g.node(id);
            const original = graph.nodes.find(n => n.id === id);

            if (original && nodeInfo) {
                output.push({
                    ...original,
                    x: nodeInfo.x,
                    y: nodeInfo.y,
                    width: nodeInfo.width,
                    height: nodeInfo.height
                });
            }
        });

        return output;
    }
}
