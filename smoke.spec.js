import { test, expect } from '@playwright/test';

test('playwright smoke test runs', () => {
    expect('keepaliverdpteams').toContain('teams');
});
