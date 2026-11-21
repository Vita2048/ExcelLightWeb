import { describe, expect, it } from 'vitest';
import { defaultWorkspacePreferences, LocalWorkspaceError, migrateLegacyPreferences } from '../../src/persistence/local';

describe('local workspace preference migration', () => {
  it('starts in Automatic native mode while keeping Python automatic consent off', () => {
    expect(defaultWorkspacePreferences.nativeCalculation).toBe('automatic');
    expect(defaultWorkspacePreferences.pythonAutomaticConsent).toBe(false);
  });

  it('migrates the desktop settings fields without turning an absent Python choice on', () => {
    const migrated = migrateLegacyPreferences({
      id: 'user', version: 1, manual: true, findHistory: ['SLA'], replacementHistory: ['Open'],
    });
    expect(migrated.nativeCalculation).toBe('manual');
    expect(migrated.pythonAutomaticConsent).toBe(false);
    expect(migrated.searchHistory).toEqual({ find: ['SLA'], replace: ['Open'] });
    expect(migrated.view).toEqual(defaultWorkspacePreferences.view);
  });

  it('preserves only an explicit stored Python Automatic choice', () => {
    expect(migrateLegacyPreferences({ manual: false, pythonAutomatic: true }).pythonAutomaticConsent).toBe(true);
    expect(migrateLegacyPreferences({ manual: false }).pythonAutomaticConsent).toBe(false);
  });

  it.each([
    { manual: 'manual' },
    { manual: false, pythonAutomatic: 'yes' },
    { findHistory: Array.from({ length: 16 }, (_, index) => String(index)) },
    { findHistory: ['same', 'same'] },
  ])('rejects invalid preference data without normalizing it silently: %o', value => {
    expect(() => migrateLegacyPreferences(value)).toThrow(LocalWorkspaceError);
  });
});
