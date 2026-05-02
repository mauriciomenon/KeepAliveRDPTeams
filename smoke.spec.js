const { test, expect } = require('@playwright/test');
const fs = require('node:fs');
const path = require('node:path');

test('desktop app source files are present', () => {
    const requiredFiles = [
        'keep-alive-app.py',
        'teams_checker.py',
        'config.json',
    ];

    for (const file of requiredFiles) {
        const fullPath = path.join(__dirname, file);
        expect(fs.existsSync(fullPath), `${file} should exist`).toBe(true);
        expect(fs.statSync(fullPath).size, `${file} should not be empty`).toBeGreaterThan(0);
    }
});
