import { describe, it, expect } from 'vitest';
import * as fs from 'fs';
import * as path from 'path';
import { MermaidParser } from '../src/mermaid/MermaidParser';

describe('MermaidParser', () => {
    const parser = new MermaidParser();

    it('should parse simple node and edge', () => {
        const text = `
        graph TD
        A --> B
        `;
        const result = parser.parse(text);
        expect(result.nodes).toContainEqual(expect.objectContaining({ id: 'A' }));
        expect(result.nodes).toContainEqual(expect.objectContaining({ id: 'B' }));
        expect(result.edges).toHaveLength(1);
        expect(result.edges[0]).toEqual({ from: 'A', to: 'B', type: 'arrow', text: undefined });
    });

    it('should parse nodes with text and shapes', () => {
        const text = `
        A[Square] --> B(Round)
        C{Rhombus} --> D((Circle))
        E[/Parallelogram/]
        `;
        const result = parser.parse(text);

        const findNode = (id: string) => result.nodes.find(n => n.id === id);

        expect(findNode('A')).toEqual(expect.objectContaining({ type: 'square', text: 'Square' }));
        expect(findNode('B')).toEqual(expect.objectContaining({ type: 'round', text: 'Round' }));
        expect(findNode('C')).toEqual(expect.objectContaining({ type: 'rhombus', text: 'Rhombus' }));
        expect(findNode('D')).toEqual(expect.objectContaining({ type: 'circle', text: 'Circle' }));
        expect(findNode('E')).toEqual(expect.objectContaining({ type: 'parallelogram', text: 'Parallelogram' }));
    });

    it('should parse styles', () => {
        const text = `
        style A fill:#f9f,stroke:#333
        A --> B
        `;
        const result = parser.parse(text);
        const nodeA = result.nodes.find(n => n.id === 'A');
        expect(nodeA?.style).toEqual({ fill: '#f9f', stroke: '#333' });
    });

    it('should parse the test-diagram.mmd asset', () => {
        const assetPath = path.join(__dirname, 'assets/test-diagram.mmd');
        const content = fs.readFileSync(assetPath, 'utf-8');
        const result = parser.parse(content);

        // Verify Core Nodes
        const findNode = (id: string) => result.nodes.find(n => n.id === id);

        // Start([Start Order])
        expect(findNode('Start')).toEqual(expect.objectContaining({ id: 'Start', text: 'Start Order', type: 'round' }));

        // Input[/Receive JSON/]
        expect(findNode('Input')).toEqual(expect.objectContaining({ id: 'Input', text: 'Receive JSON', type: 'parallelogram' }));

        // Auth{Is Authorized?}
        expect(findNode('Auth')).toEqual(expect.objectContaining({ id: 'Auth', text: 'Is Authorized?', type: 'rhombus' }));

        // Reject[Reject Request]:::failure
        const reject = findNode('Reject');
        expect(reject).toEqual(expect.objectContaining({ id: 'Reject', text: 'Reject Request', type: 'square' }));
        expect(reject?.style).toEqual(expect.objectContaining({ className: 'failure' }));

        // Verify Edges
        // Auth -- No --> Reject
        const noEdge = result.edges.find(e => e.from === 'Auth' && e.to === 'Reject');
        expect(noEdge).toBeDefined();
        expect(noEdge?.text).toBe('No');

        // Verify Layout Complexity
        // Check total counts approximately
        expect(result.nodes.length).toBeGreaterThan(5);
        expect(result.edges.length).toBeGreaterThan(5);
    });
});
