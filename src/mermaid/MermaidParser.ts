export interface GraphNode {
    id: string;
    text: string;
    type: 'square' | 'round' | 'rhombus' | 'circle' | 'cylinder' | 'parallelogram';
    style?: any;
}

export interface GraphEdge {
    from: string;
    to: string;
    text?: string;
    type: 'arrow' | 'dotted';
}

export interface GraphData {
    nodes: GraphNode[];
    edges: GraphEdge[];
}

export class MermaidParser {
    parse(text: string): GraphData {
        const nodes: Map<string, GraphNode> = new Map();
        const edges: GraphEdge[] = [];
        const styles: Map<string, any> = new Map();
        const classes: Map<string, any> = new Map();

        const lines = text.split('\n').map(l => l.trim()).filter(l => l && !l.startsWith('%%') && !l.startsWith('graph') && !l.startsWith('flowchart'));

        // Regex patterns
        // Node: ID[Text] or ID((Text)) etc.
        // Captures: 1=ID, 2=OpenBracket, 3=Text, 4=CloseBracket
        const nodeRegex = /^([A-Za-z0-9_]+)(\[|\(|\{\{|\{|\(\(|\[\/)(.*?)(\]|\)|\}\}|\}|\)\)|\])$/;

        // Edge: A --> B or A -- Text --> B
        // Captures: 1=From, 2=LinkType(---,-->), 3=Text(optional), 4=To
        // Simplified: Split by arrow types

        // Command: style ID fill:#f9f...
        const styleRegex = /^style\s+([A-Za-z0-9_]+)\s+(.*)$/;

        // ClassDef: classDef className style...
        const classDefRegex = /^classDef\s+([A-Za-z0-9_]+)\s+(.*)$/;

        for (const line of lines) {
            // Style
            if (line.match(styleRegex)) {
                const match = line.match(styleRegex);
                if (match) {
                    styles.set(match[1], this.parseStyle(match[2]));
                }
                continue;
            }

            // ClassDef
            if (line.match(classDefRegex)) {
                const match = line.match(classDefRegex);
                if (match) {
                    classes.set(match[1], this.parseStyle(match[2]));
                }
                continue;
            }

            // Edges (and implicit nodes)
            // Supported Arrows: -->, ---, -- Text -->
            if (line.includes('-->') || line.includes('---')) {
                this.parseEdgeLine(line, nodes, edges);
                continue;
            }

            // Explicit Node Definition might look like "A[Label]" on its own line
            // But often combined with edges.
            // If it matches node regex directly:
            const nodeMatch = line.match(nodeRegex);
            if (nodeMatch) {
                this.getOrAddNode(nodeMatch[1], nodeMatch[0], nodes);
            }
        }

        // Apply styles to nodes
        for (const node of nodes.values()) {
            // Direct style
            if (styles.has(node.id)) {
                node.style = { ...node.style, ...styles.get(node.id) };
            }
            // Class style (if we parsed classes application ":::className") - NOTE: regex above needs to handle :::
        }

        return {
            nodes: Array.from(nodes.values()),
            edges
        };
    }

    private parseStyle(styleStr: string): any {
        const style: any = {};
        const parts = styleStr.split(',');
        for (const part of parts) {
            const [key, val] = part.split(':');
            if (key && val) style[key.trim()] = val.trim();
        }
        return style;
    }

    private parseEdgeLine(line: string, nodes: Map<string, GraphNode>, edges: GraphEdge[]) {
        // Splitting complex lines like "A --> B --> C"
        // Regex to separate nodes/edges is tricky.
        // Let's use a tokenizer approach or simplified splitting.

        // Example: Start([Start Order]) --> Input[/Receive JSON/]
        // Tokens: "Start([Start Order])", "-->", "Input[/Receive JSON/]"

        // Split by arrow patterns, keeping delimiters
        const parts = line.split(/(\s*--\s.*?\s-->\s*|\s*-->\s*|\s*---\s*)/).map(p => p.trim()).filter(p => p);

        // Expect format: Node, Arrow, Node, [Arrow, Node...]
        for (let i = 0; i < parts.length - 1; i += 2) {
            const fromRaw = parts[i];
            const arrow = parts[i + 1];
            const toRaw = parts[i + 2];

            if (!toRaw) break;

            const fromNode = this.parseNodeString(fromRaw, nodes);
            const toNode = this.parseNodeString(toRaw, nodes);

            let edgeText: string | undefined;
            if (arrow.includes('--')) {
                // specific check for "-- Text -->"
                const textMatch = arrow.match(/--\s(.*?)\s-->/);
                if (textMatch) edgeText = textMatch[1];
            }

            edges.push({
                from: fromNode.id,
                to: toNode.id,
                text: edgeText,
                type: 'arrow' // Defaulting to arrow for now
            });
        }
    }

    private getOrAddNode(id: string, raw: string, nodes: Map<string, GraphNode>): GraphNode {
        // Reuse parseNodeString logic
        const parsed = this.parseNodeString(raw, nodes);
        // parseNodeString already adds/updates the node in the map, so we are good.
        return parsed;
    }

    private parseNodeString(raw: string, nodes: Map<string, GraphNode>): GraphNode {
        // Handle Class Suffix ":::className"
        let className: string | undefined;
        if (raw.includes(':::')) {
            const parts = raw.split(':::');
            raw = parts[0];
            className = parts[1];
        }

        // Match Node Syntax
        // A, A[Text], A((Text)), etc.
        const idRegex = /^([A-Za-z0-9_\.]+)/;
        const match = raw.match(idRegex);
        if (!match) {
            // Fallback: Use raw as ID if no special chars found?
            // But we shouldn't get here usually if logic above is sound.
            // throw new Error(`Invalid node format: ${raw}`);
            // Actually, if we pass "Start([Start Order])", it matches "Start".
        }

        // If match failed, likely means bad input or previous parsing stage error.
        if (!match) throw new Error(`Invalid node format: ${raw}`);

        const id = match[1];

        // If it has brackets, register/update the node definition
        // Types: [ ] square, ( ) round, (( )) circle, { } rhombus, [/ /] parallelogram
        let text = id;
        let type: GraphNode['type'] = 'square'; // Default

        // Detection needs to be specific order to avoid partial matches
        // e.g. check [/ first, then [
        // check (( first, then (
        // check [( first, then [

        if (raw.includes('[( ') || raw.includes('[(')) {
            type = 'cylinder';
            text = this.extractText(raw, '\\[\\(', '\\)\\]');
        } else if (raw.includes('[/') && raw.includes('/]')) {
            type = 'parallelogram';
            text = this.extractText(raw, '\\[\\/', '\\/\\]');
        } else if (raw.includes('[[') && raw.includes(']]')) {
            // subrotine? Let's just treat as square or similar for now if not specified?
            // Prompt didn't specify. Let's stick to prompt.
            type = 'square';
            text = this.extractText(raw, '\\[\\[', '\\]\\]');
        } else if (raw.includes('([') && raw.includes('])')) {
            type = 'round'; // Stadium -> Mapped to Round/Terminator
            text = this.extractText(raw, '\\(\\[', '\\]\\)');
        } else if (raw.includes('((') && raw.includes('))')) {
            type = 'circle';
            text = this.extractText(raw, '\\(\\(', '\\)\\)');
        } else if (raw.includes('{{') && raw.includes('}}')) {
            type = 'rhombus'; // Hexagon usually. Mapping to Rhombus/Decision for fallback?
            text = this.extractText(raw, '\\{\\{', '\\}\\}');
        } else if (raw.includes('{') && raw.includes('}')) {
            type = 'rhombus';
            text = this.extractText(raw, '\\{', '\\}');
        } else if (raw.includes('(') && raw.includes(')')) {
            type = 'round';
            text = this.extractText(raw, '\\(', '\\)');
        } else if (raw.includes('[') && raw.includes(']')) {
            type = 'square';
            text = this.extractText(raw, '\\[', '\\]');
        }

        // Update or Create Node
        // If node exists, only update properties if this usage is more specific (has text/type)
        let node = nodes.get(id);
        if (!node) {
            node = { id, text, type };
            nodes.set(id, node);
        } else {
            // Upsert: if we found meaningful text/type, update it.
            // Avoid overwriting a specific type with a default 'square' if we just encountered 'A' -> 'B' later.
            // But here 'raw' usually contains the definition.
            // If raw is just "A", type stays default square, text "A".
            // If we parse "A", we shouldn't overwrite "A[Box]".

            // A definition is considered "specific" if it contains more than just the ID
            const isSpecificDefinition = raw.length > id.length;
            if (isSpecificDefinition) {
                // Only update text if it's not just the ID itself
                if (text !== id) node.text = text;
                // Always update type if a specific shape syntax was found
                node.type = type;
            }
        }

        // Store class reference
        if (className) {
            node.style = { ...node.style, className };
        }

        return node;
    }

    private extractText(raw: string, open: string, close: string): string {
        const regex = new RegExp(`${open}(.*?)${close}`);
        const match = raw.match(regex);
        return match ? match[1] : raw;
    }
}
