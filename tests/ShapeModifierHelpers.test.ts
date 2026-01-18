import { describe, it, expect, beforeEach, vi } from 'vitest';
import { ShapeModifier } from '../src/ShapeModifier';
import { VisioPackage } from '../src/VisioPackage';

// Mock VisioPackage since we only test internal helpers that don't use it directly
// or use it only for file validtion (which we might skip or mock)
const mockPkg = {
    getFileText: vi.fn(),
    updateFile: vi.fn(),
} as unknown as VisioPackage;

describe('ShapeModifier Helpers', () => {
    let modifier: ShapeModifier;

    beforeEach(() => {
        modifier = new ShapeModifier(mockPkg);
    });

    describe('getEdgePoint', () => {
        it('should return center if start and target are the same', () => {
            const result = (modifier as any).getEdgePoint(10, 10, 4, 2, 10, 10);
            expect(result).toEqual({ x: 10, y: 10 });
        });

        it('should intersect right edge when target is directly right', () => {
            // Box Center (10, 10), Size 4x2.
            // Right edge x = 10 + 2 = 12.
            const result = (modifier as any).getEdgePoint(10, 10, 4, 2, 20, 10);
            expect(result.x).toBeCloseTo(12);
            expect(result.y).toBeCloseTo(10);
        });

        it('should intersect left edge when target is directly left', () => {
            // Left edge x = 10 - 2 = 8.
            const result = (modifier as any).getEdgePoint(10, 10, 4, 2, 0, 10);
            expect(result.x).toBeCloseTo(8);
            expect(result.y).toBeCloseTo(10);
        });

        it('should intersect top edge when target is directly above', () => {
            // Top edge y = 10 + 1 = 11. (Assuming y grows upwards? Visio coords usually valid cartesian or inverted depending on perspective, logic assumes math standard)
            // Visio (0,0) is bottom-left usually. So "up" is +y.
            const result = (modifier as any).getEdgePoint(10, 10, 4, 2, 10, 20);
            expect(result.x).toBeCloseTo(10);
            expect(result.y).toBeCloseTo(11);
        });

        it('should intersect bottom edge when target is directly below', () => {
            // Bottom edge y = 10 - 1 = 9.
            const result = (modifier as any).getEdgePoint(10, 10, 4, 2, 10, 0);
            expect(result.x).toBeCloseTo(10);
            expect(result.y).toBeCloseTo(9);
        });

        it('should intersect corner properly (45 degrees)', () => {
            // Box 2x2 (Square). Half-width = 1.
            // Vector (1, 1) angle.
            // Should hit (11, 11).
            const result = (modifier as any).getEdgePoint(10, 10, 2, 2, 15, 15);
            expect(result.x).toBeCloseTo(11);
            expect(result.y).toBeCloseTo(11);
        });
    });

    describe('getAbsolutePos', () => {
        // Mock Shape Hierarchy
        const mockHierarchy = new Map();

        // Helper to create mock shape entry
        const createEntry = (id: string, pinX: string, pinY: string, parentId?: string, locPinX = '0', locPinY = '0') => {
            return {
                shape: {
                    '@_ID': id,
                    Cell: [
                        { '@_N': 'PinX', '@_V': pinX },
                        { '@_N': 'PinY', '@_V': pinY },
                        { '@_N': 'LocPinX', '@_V': locPinX },
                        { '@_N': 'LocPinY', '@_V': locPinY }
                    ]
                },
                parent: parentId ? mockHierarchy.get(parentId).shape : null
            };
        };

        it('should return local coordinates for top-level shape', () => {
            mockHierarchy.clear();
            mockHierarchy.set('1', createEntry('1', '10', '20'));

            const result = (modifier as any).getAbsolutePos('1', mockHierarchy);
            expect(result).toEqual({ x: 10, y: 20 });
        });

        it('should calculate global coordinates for nested shape (1 level deep)', () => {
            // Parent: Pin(10, 10), LocPin(0,0) -> Origin (10, 10)
            // Child: Pin(2, 3) relative to Parent.
            // Global: (12, 13)
            mockHierarchy.clear();
            mockHierarchy.set('1', createEntry('1', '10', '10')); // Parent
            mockHierarchy.set('2', createEntry('2', '2', '3', '1')); // Child

            const result = (modifier as any).getAbsolutePos('2', mockHierarchy);
            expect(result).toEqual({ x: 12, y: 13 });
        });

        it('should handle complex nesting (2 levels) and LocPin offsets', () => {
            // Grandparent: Pin(100, 100), LocPin(0,0) -> Origin(100, 100)
            // Parent: Pin(10, 10) relative to GP. LocPin(5, 5).
            // Parent Origin in GP space: Pin(10,10) - LocPin(5,5) = (5, 5).
            // Parent Origin Global: GP(100,100) + (5,5) = (105, 105).
            // Child: Pin(1, 1) relative to Parent.
            // Child Global: (106, 106).

            mockHierarchy.clear();
            mockHierarchy.set('1', createEntry('1', '100', '100'));
            mockHierarchy.set('2', createEntry('2', '10', '10', '1', '5', '5'));
            mockHierarchy.set('3', createEntry('3', '1', '1', '2'));

            const result = (modifier as any).getAbsolutePos('3', mockHierarchy);
            expect(result).toEqual({ x: 106, y: 106 });
        });
    });

    describe('buildShapeHierarchy', () => {
        it('should build flat hierarchy correctly', () => {
            const parsed = {
                PageContents: {
                    Shapes: {
                        Shape: [
                            { '@_ID': '1' },
                            { '@_ID': '2' }
                        ]
                    }
                }
            };
            const map = (modifier as any).buildShapeHierarchy(parsed);
            expect(map.size).toBe(2);
            expect(map.get('1').parent).toBeNull();
            expect(map.get('2').parent).toBeNull();
        });

        it('should build nested hierarchy correctly', () => {
            const parsed = {
                PageContents: {
                    Shapes: {
                        Shape: [
                            {
                                '@_ID': '1',
                                Shapes: {
                                    Shape: [
                                        { '@_ID': '2' }
                                    ]
                                }
                            }
                        ]
                    }
                }
            };
            const map = (modifier as any).buildShapeHierarchy(parsed);
            expect(map.size).toBe(2);

            const p1 = map.get('1');
            const p2 = map.get('2');

            expect(p1.parent).toBeNull();
            expect(p2.parent).toBe(p1.shape);
        });
    });
});
