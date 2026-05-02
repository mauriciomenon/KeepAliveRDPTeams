const { test, expect } = require('@playwright/test');
const packageJson = require('./package.json');

test('package metadata content is valid', () => {
    expect(packageJson.name).toBe('keepaliverdpteams');
    expect(packageJson.devDependencies['@playwright/test']).toBe('1.59.1');
});
