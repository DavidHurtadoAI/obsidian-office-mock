import { readFile } from 'node:fs/promises';
import assert from 'node:assert/strict';
import stylelint from 'stylelint';

const result = await stylelint.lint({
  code: await readFile(new URL('../theme.css', import.meta.url), 'utf8'),
  config: {
    plugins: ['stylelint-no-unsupported-browser-features'],
    rules: {
      'plugin/no-unsupported-browser-features': [true, {
        browsers: ['chrome 114'], severity: 'warning',
      }],
    },
  },
});
const warnings = result.results.flatMap(r => r.warnings).map(({ rule, text, line }) => ({ rule, text, line }));
assert.equal(warnings.length, 0, JSON.stringify(warnings, null, 2));
assert.equal(result.errored, false);
console.log('PASS: Chromium 114 compatibility, zero warnings.');
