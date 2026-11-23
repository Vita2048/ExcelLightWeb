import { expect, it } from 'vitest';
import { selectAllCornerTriangle } from '../../src/grid/draw-worksheet';
import { gridMetrics } from '../../src/grid/grid-geometry';

it('draws the select-all corner marker as a centered isosceles right triangle', () => {
  const [bottomLeft, bottomRight, topRight] = selectAllCornerTriangle(gridMetrics.rowHeader, gridMetrics.columnHeader);
  const horizontalLeg = bottomRight.x - bottomLeft.x;
  const verticalLeg = bottomRight.y - topRight.y;

  expect(horizontalLeg).toBeGreaterThan(0);
  expect(horizontalLeg).toBeCloseTo(verticalLeg, 8);
  expect(bottomLeft.x).toBeGreaterThanOrEqual(0);
  expect(topRight.y).toBeGreaterThanOrEqual(0);
  expect(bottomRight.x).toBeLessThanOrEqual(gridMetrics.rowHeader);
  expect(bottomLeft.y).toBeLessThanOrEqual(gridMetrics.columnHeader);
});
