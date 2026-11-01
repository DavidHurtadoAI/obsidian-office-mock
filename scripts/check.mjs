import assert from 'node:assert/strict';
import { readFile, readdir } from 'node:fs/promises';
import postcss from 'postcss';

const root = new URL('../', import.meta.url);
const read = name => readFile(new URL(name, root), 'utf8');
const manifest = JSON.parse(await read('manifest.json'));
const pkg = JSON.parse(await read('package.json'));
const versions = JSON.parse(await read('versions.json'));
assert.equal(manifest.name, 'Office Mock');
assert.match(manifest.version, /^\d+\.\d+\.\d+$/);
assert.equal(pkg.version, manifest.version);
assert.equal(versions[manifest.version], manifest.minAppVersion);
for (const key of ['name', 'version', 'minAppVersion', 'author']) assert.equal(typeof manifest[key], 'string');
const license = await read('LICENSE');
const copyright = license.match(/^Copyright \(c\) (\d{4}) (.+)$/m);
assert.ok(copyright, 'Project copyright is present');
assert.equal(copyright[2], manifest.author);
const standardMIT = (await read('tests/mit-template.txt')).replace('[year]',copyright[1]).replace('[fullname]',manifest.author);
const normalizeLicense = text => text.replace(/\s+/g,' ').trim();
assert.equal(normalizeLicense(license),normalizeLicense(standardMIT), 'LICENSE must contain only the standard MIT text; third-party notices belong in their own file');
assert.match(await read('THIRD_PARTY_NOTICES.md'), /Lucide[\s\S]*Feather/);

const css = await read('theme.css');
const tree = postcss.parse(css);
assert.ok(!/\bC:\\|@Nexus/.test(css), 'No local paths in the distributed theme');
assert.ok(!/:has\(/.test(css), 'No expensive ancestor matching');
assert.ok(!/@import\b/.test(css), 'No network imports');
assert.ok(!/url\(\s*["']?(?!data:)[^\s"')]/i.test(css), 'Every resource is embedded');
let selectors = 0, resources = 0;
const declarations = [];
tree.walkDecls(decl => declarations.push(decl));
for (const decl of declarations) {
  assert.ok(!decl.important, 'No !important declarations');
  if (decl.prop === 'font-family') assert.ok(!/\b(?:Aptos|Calibri|Segoe|Inter|Arial)\b/i.test(decl.value), 'No imposed font family');
  if (decl.prop === '--font-text-size') assert.fail('Do not override the user font size');
  for (const match of decl.value.matchAll(/url\("data:image\/svg\+xml,([^"]*)"\)/g)) {
    const svg = decodeURIComponent(match[1]);
    assert.ok(svg.startsWith('<svg') && svg.endsWith('</svg>'), 'Valid inline SVG wrapper');
    assert.ok(!/<text\b|<script\b|https?:\/\/(?!www\.w3\.org\/2000\/svg)/i.test(svg), 'SVG contains geometry only');
    if (decl.prop === 'background-image' && decl.parent.selector.startsWith('.theme-dark') && decl.parent.selector.includes('markdown-preview-sizer') && !decl.parent.selector.includes('::')) {
      assert.ok(!/fill=(["'])white\1/.test(svg), 'Dark page backdrop cannot reveal white seams');
    }
    resources++;
  }
}
tree.walkRules(() => selectors++);
assert.ok(css.includes('.theme-light') && css.includes('.theme-dark'));
assert.ok(css.includes('office-plain') && css.includes('is-live-preview'));
assert.ok(resources > 20 && selectors > 200);
const thumbnail = await readFile(new URL('images/screenshot.png',root));
assert.equal(thumbnail.subarray(1,4).toString(),'PNG');
assert.equal(thumbnail.readUInt32BE(16),512);
assert.equal(thumbnail.readUInt32BE(20),288);

let notes = 0;
for (const folder of ['examples', 'templates']) {
  const files = (await readdir(new URL(folder + '/', root))).filter(name => name.endsWith('.md'));
  assert.equal(files.length, 4);
  const types = new Set();
  for (const file of files) {
    const md = await read(folder + '/' + file);
    const type = md.match(/^---\ncssclasses:\n  - office-(email|meeting|document|spreadsheet)\n---\n\n# /);
    assert.ok(type, file + ': ordinary frontmatter and title'); types.add(type[1]);
    const text = md.replace(/<[\w.+-]+@[\w.-]+\.[a-z]+>/gi, '');
    assert.ok(!/<\/?[a-z!]|\bod-|\[![^\]]+\]|aria-hidden|<!--/i.test(text), file + ': no UI markup');
    for (const email of md.matchAll(/[\w.+-]+@([\w.-]+\.[a-z]+)/gi)) assert.ok(email[1].endsWith('.example'));
    notes++;
  }
  assert.equal(types.size, 4);
}
function luminance(hex) {
  const parts = hex.match(/[\da-f]{2}/gi).map(x => parseInt(x,16)/255).map(x => x <= .04045 ? x/12.92 : ((x+.055)/1.055)**2.4);
  return parts[0]*.2126 + parts[1]*.7152 + parts[2]*.0722;
}
const pairs = [
  ['#202d3f','#ffffff'], ['#526176','#ffffff'], ['#e7edf5','#17202d'], ['#b0bfd0','#17202d'],
  ['#ffffff','#1a5fc2'], ['#ffffff','#2f4fae'], ['#ffffff','#1f7a4a'], ['#ffffff','#6b3fc7'],
  ['#8fb9ff','#1c2736'], ['#a6b7ff','#1c2736'], ['#7dd5a0','#1c2736'], ['#c6adff','#1c2736'],
];
for (const [fg,bg] of pairs) {
  const [high,low] = [luminance(fg),luminance(bg)].sort((a,b)=>b-a);
  assert.ok((high+.05)/(low+.05) >= 4.5, `Readable token pair ${fg}/${bg}`);
}
console.log(`PASS: manifest/license, ${selectors} CSS rules, ${resources} offline SVG uses, ${notes} clean notes, ${pairs.length} contrast pairs.`);
