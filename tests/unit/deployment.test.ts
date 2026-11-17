import { describe, expect, it } from 'vitest';
import { deploymentBase } from '../../config/base-path';

describe('Pages asset base', () => {
  it('supports relative assets, a repository path, and a domain root', () => {
    expect(deploymentBase()).toBe('./');
    expect(deploymentBase('/ExcelLightWeb/')).toBe('/ExcelLightWeb/');
    expect(deploymentBase('/')).toBe('/');
  });
  it.each(['ExcelLightWeb', '/ExcelLightWeb', 'https://example.com/', '//host/', '/../'])('rejects invalid base %s', value => {
    expect(() => deploymentBase(value)).toThrow();
  });
});
