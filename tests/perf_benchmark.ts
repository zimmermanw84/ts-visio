
import { performance } from 'perf_hooks';

// Mock data generator (Iterative to avoid stack overflow during generation)
function generateNestedShapes(depth: number, width: number): any {
    let currentLevelShapes: any[] = [];
    // Create leaves
    for (let i = 0; i < Math.pow(width, depth); i++) { // Note: for width=1 this is just 1
         // For width > 1 this approach is hard without recursion or complex logic.
         // But for width=1 (linear), we can just build bottom up.
    }

    if (width === 1) {
        let node = { '@_ID': 'leaf', Shapes: null };
        for (let i = 1; i <= depth; i++) {
            node = {
                '@_ID': `node-${i}`,
                Shapes: {
                    Shape: [node]
                }
            };
        }
        return node;
    }

    // Fallback to recursive for small depth/width (won't work for deep check if width > 1)
    if (depth === 0) return { '@_ID': 'leaf', Shapes: null };
    const shapes: any[] = [];
    for (let i = 0; i < width; i++) {
        shapes.push(generateNestedShapes(depth - 1, width));
    }
    return {
        '@_ID': `node-${depth}`,
        Shapes: {
            Shape: shapes
        }
    };
}

// Prepare data
console.log('Generating test data...');
const DEPTH = 15; // Deep nesting
const WIDTH = 2;  // Branching factor
// Total nodes approx WIDTH^DEPTH. 2^15 is ~32768.
// Note: Deep nesting (e.g. Depth 20000, Width 1) causes Stack Overflow in Recursive versions,
// while Iterative version handles it easily.

const topLevel = [];
for (let i=0; i<WIDTH; i++) {
    topLevel.push(generateNestedShapes(DEPTH, WIDTH));
}

const parsed = {
    PageContents: {
        Shapes: {
            Shape: topLevel
        }
    }
};
console.log('Test data generated.');

// Variant 1: Concat (The Anti-Pattern)
function getAllShapes_Concat(parsed: any): any[] {
    let topLevelShapes = parsed.PageContents.Shapes ? parsed.PageContents.Shapes.Shape : [];
    if (!Array.isArray(topLevelShapes)) {
        topLevelShapes = topLevelShapes ? [topLevelShapes] : [];
    }

    const gather = (shapeList: any[]): any[] => {
        let all: any[] = [];
        for (const s of shapeList) {
            all.push(s);
            if (s.Shapes && s.Shapes.Shape) {
                const children = Array.isArray(s.Shapes.Shape) ? s.Shapes.Shape : [s.Shapes.Shape];
                all = all.concat(gather(children));
            }
        }
        return all;
    };

    return gather(topLevelShapes);
}

// Variant 2: Recursive Closure (Current Implementation)
function getAllShapes_Recursive(parsed: any): any[] {
    let topLevelShapes = parsed.PageContents.Shapes ? parsed.PageContents.Shapes.Shape : [];
    if (!Array.isArray(topLevelShapes)) {
        topLevelShapes = topLevelShapes ? [topLevelShapes] : [];
    }

    const all: any[] = [];
    const gather = (shapeList: any[]): void => {
        for (const s of shapeList) {
            all.push(s);
            if (s.Shapes && s.Shapes.Shape) {
                const children = Array.isArray(s.Shapes.Shape) ? s.Shapes.Shape : [s.Shapes.Shape];
                gather(children);
            }
        }
    };

    gather(topLevelShapes);
    return all;
}

// Variant 3: Iterative Stack (Proposed Optimization)
function getAllShapes_Iterative(parsed: any): any[] {
    let topLevelShapes = parsed.PageContents.Shapes ? parsed.PageContents.Shapes.Shape : [];
    if (!Array.isArray(topLevelShapes)) {
        topLevelShapes = topLevelShapes ? [topLevelShapes] : [];
    }

    const all: any[] = [];
    // Use a stack for DFS. Push in reverse order to maintain Pre-order traversal order (same as recursive)
    const stack: any[] = [];
    for (let i = topLevelShapes.length - 1; i >= 0; i--) {
        stack.push(topLevelShapes[i]);
    }

    while (stack.length > 0) {
        const s = stack.pop();
        all.push(s);

        if (s.Shapes && s.Shapes.Shape) {
            const children = Array.isArray(s.Shapes.Shape) ? s.Shapes.Shape : [s.Shapes.Shape];
            // Push children in reverse order so they are popped in original order
            for (let i = children.length - 1; i >= 0; i--) {
                stack.push(children[i]);
            }
        }
    }
    return all;
}

// Run Benchmarks
function runBenchmark(name: string, fn: (p: any) => any[]) {
    try {
        const start = performance.now();
        const result = fn(parsed);
        const end = performance.now();
        console.log(`${name}: ${(end - start).toFixed(2)}ms. Found ${result.length} shapes.`);
        return result.length;
    } catch (e) {
        console.log(`${name}: FAILED (${e.message})`);
        return -1;
    }
}

// Warmup
// runBenchmark('Warmup (Recursive)', getAllShapes_Recursive); // Skip warmup that might crash

console.log('\n--- Starting Benchmark ---');
const count1 = runBenchmark('Concat (Anti-Pattern)', getAllShapes_Concat);
const count2 = runBenchmark('Recursive (Current)  ', getAllShapes_Recursive);
const count3 = runBenchmark('Iterative (Proposed) ', getAllShapes_Iterative);

if (count1 !== count2 || count2 !== count3) {
    console.error('ERROR: Result counts do not match!');
    console.error(`Concat: ${count1}, Recursive: ${count2}, Iterative: ${count3}`);
    process.exit(1);
} else {
    console.log('All implementations returned correct shape count.');
}
