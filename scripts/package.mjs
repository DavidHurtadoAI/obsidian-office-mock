import { readFile, writeFile, readdir, mkdir, copyFile } from 'node:fs/promises';
import { zipSync, unzipSync } from 'fflate';
import assert from 'node:assert/strict';

const root = new URL('../', import.meta.url);
const manifest = JSON.parse(await readFile(new URL('manifest.json', root), 'utf8'));
const files = {};
for (const name of ['theme.css', 'manifest.json', 'LICENSE', 'THIRD_PARTY_NOTICES.md']) {
  files['Office Mock/' + name] = new Uint8Array(await readFile(new URL(name, root)));
}
for (const name of ['README.md', 'LICENSE', 'THIRD_PARTY_NOTICES.md', 'CHANGELOG.md', 'CONTRIBUTING.md']) {
  files[name] = new Uint8Array(await readFile(new URL(name,root)));
}
for (const folder of ['examples', 'templates', 'docs', 'images']) for (const name of await readdir(new URL(folder + '/', root))) {
  if (/\.(md|png)$/.test(name)) files[folder + '/' + name] = new Uint8Array(await readFile(new URL(folder + '/' + name, root)));
}
const zip = zipSync(files, { level: 9 });
const unpacked = unzipSync(zip);
for (const [name, bytes] of Object.entries(files)) assert.deepEqual(unpacked[name], bytes, name);
assert.deepEqual(Object.keys(unpacked).sort(),Object.keys(files).sort());
assert.ok(Object.keys(unpacked).every(name => !/\.obsidian|\.m?js$|node_modules/.test(name)));
await mkdir(new URL('dist/', root), { recursive: true });
await writeFile(new URL(`dist/Office-Mock-${manifest.version}.zip`, root), zip);
await copyFile(new URL('theme.css', root), new URL('dist/theme.css', root));
await copyFile(new URL('manifest.json', root), new URL('dist/manifest.json', root));
console.log(`Packaged ${manifest.version}; verified all ${Object.keys(unpacked).length} ZIP entries against source bytes.`);
