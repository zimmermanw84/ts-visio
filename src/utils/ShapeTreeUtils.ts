/**
 * Pure utility functions for traversing the raw parsed-XML shape tree that
 * appears inside every `PageContents` document.  All three functions accept
 * the top-level `parsed` object returned by fast-xml-parser so callers do not
 * need to repeat the same boilerplate extraction / normalization code.
 */

function topLevelShapes(parsed: any): any[] {
    const raw = parsed.PageContents?.Shapes?.Shape;
    if (!raw) return [];
    return Array.isArray(raw) ? raw : [raw];
}

function childShapes(shape: any): any[] {
    const raw = shape.Shapes?.Shape;
    if (!raw) return [];
    return Array.isArray(raw) ? raw : [raw];
}

/**
 * Return every shape in the page tree as a flat array of raw parsed objects,
 * including shapes nested inside groups at any depth.
 */
export function gatherAllShapes(parsed: any): any[] {
    const result: any[] = [];
    const recurse = (shapes: any[]) => {
        for (const s of shapes) {
            result.push(s);
            const children = childShapes(s);
            if (children.length) recurse(children);
        }
    };
    recurse(topLevelShapes(parsed));
    return result;
}

/**
 * Find a single raw shape by its `@_ID` attribute anywhere in the page tree.
 * Returns `undefined` if no shape with that ID exists.
 */
export function findShapeById(parsed: any, id: string): any | undefined {
    const recurse = (shapes: any[]): any | undefined => {
        for (const s of shapes) {
            if (s['@_ID'] === id) return s;
            const children = childShapes(s);
            if (children.length) {
                const found = recurse(children);
                if (found) return found;
            }
        }
        return undefined;
    };
    return recurse(topLevelShapes(parsed));
}

/**
 * Build a `Map<id, raw shape>` for every shape in the page tree.
 * Suitable for O(1) repeated lookups by ID.
 */
export function buildShapeMap(parsed: any): Map<string, any> {
    const map = new Map<string, any>();
    const recurse = (shapes: any[]) => {
        for (const s of shapes) {
            map.set(s['@_ID'], s);
            const children = childShapes(s);
            if (children.length) recurse(children);
        }
    };
    recurse(topLevelShapes(parsed));
    return map;
}
