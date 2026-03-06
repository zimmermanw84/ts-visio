
import { describe, it, expect } from 'vitest';
import { VisioDocument } from '../src/VisioDocument';
import { ShapeModifier } from '../src/ShapeModifier';

describe('Swimlane API', () => {
    // regression bug-21: addSwimlaneLane must attach the lane to the pool
    it('should automatically attach lanes to the pool as list members', async () => {
        const doc = await VisioDocument.create();
        const page = doc.pages[0];

        const pool = await page.addSwimlanePool({
            text: 'Pool',
            x: 5, y: 6, width: 8, height: 6,
        });

        const lane1 = await page.addSwimlaneLane(pool, {
            text: 'Lane 1',
            x: 0, y: 0, width: 8, height: 2,
        });

        const lane2 = await page.addSwimlaneLane(pool, {
            text: 'Lane 2',
            x: 0, y: 0, width: 8, height: 2,
        });

        const mod = new ShapeModifier((doc as any).pkg);
        const members = mod.getContainerMembers(page.id, pool.id);

        expect(members).toContain(lane1.id);
        expect(members).toContain(lane2.id);
        expect(members).toHaveLength(2);
    });

    // regression bug-21: lanes without pool attachment produced disconnected shapes
    it('should position lanes relative to pool after attachment', async () => {
        const doc = await VisioDocument.create();
        const page = doc.pages[0];

        const pool = await page.addSwimlanePool({
            text: 'Pool',
            x: 5, y: 6, width: 8, height: 4,
        });

        const lane = await page.addSwimlaneLane(pool, {
            text: 'Lane',
            x: 0, y: 0, width: 8, height: 2,
        });

        const mod = new ShapeModifier((doc as any).pkg);
        const poolGeo = mod.getShapeGeometry(page.id, pool.id);
        const laneGeo = mod.getShapeGeometry(page.id, lane.id);

        // Lane should be positioned within or near the pool's bounds
        const poolLeft   = poolGeo.x - poolGeo.width / 2;
        const poolRight  = poolGeo.x + poolGeo.width / 2;
        const poolBottom = poolGeo.y - poolGeo.height / 2;
        const poolTop    = poolGeo.y + poolGeo.height / 2;

        expect(laneGeo.x).toBeGreaterThanOrEqual(poolLeft - 0.5);
        expect(laneGeo.x).toBeLessThanOrEqual(poolRight + 0.5);
        expect(laneGeo.y).toBeGreaterThanOrEqual(poolBottom - 0.5);
        expect(laneGeo.y).toBeLessThanOrEqual(poolTop + 0.5);
    });
});
