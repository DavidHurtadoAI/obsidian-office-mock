import { readFile, writeFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import postcss from 'postcss';
import stylelint from 'stylelint';
import selectorParser from 'postcss-selector-parser';
import { selectorSpecificity, compare } from '@csstools/selector-specificity';

const root = new URL('../', import.meta.url);
const read = name => readFile(new URL(name, root), 'utf8');
const manifest = JSON.parse(await read('manifest.json'));
const base = postcss.parse(await read('src/base.css'));
const views = postcss.parse(await read('src/office.css'));

// Ordinary notes are documents. An explicit office-plain class opts out.
// No inspection of the native Properties DOM or :has() is needed.
function documentDefaults(selector) {
  let start = 0;
  while ((start = selector.indexOf(':where(', start)) !== -1) {
    let end = start + 7, depth = 1;
    for (; end < selector.length && depth; end++) {
      if (selector[end] === '(') depth++;
      if (selector[end] === ')') depth--;
    }
    const part = selector.slice(start, end);
    if (part.includes(':has(')) {
      const native = part.includes('.markdown-reading-view')
        ? '.markdown-reading-view > .markdown-preview-view'
        : '.markdown-source-view.mod-cm6.is-live-preview';
      const replacement = `:where(${native}:not(.office-demo, .office-email, .office-meeting, .office-document, .office-spreadsheet, .office-plain))`;
      selector = selector.slice(0, start) + replacement + selector.slice(end);
      start += replacement.length;
    } else start = end;
  }
  return selector;
}
views.walkRules(rule => {
  if (rule.selector.includes(':has(> td:first-child')) { rule.remove(); return; }
  rule.selector = documentDefaults(rule.selector)
    .replaceAll(':has(+ .mod-footer)', ':nth-last-child(2)')
    .replaceAll(':has(> td:focus-within)', ':focus-within')
    .replaceAll('.cm-contentContainer', '[class~="cm-contentContainer"]')
    .replaceAll('.HyperMD-header', '[class~="HyperMD-header"]')
    .replaceAll('.HyperMD-quote', '[class~="HyperMD-quote"]')
    .replaceAll('.markdown-preview-view:is(', '.markdown-preview-view:where(')
    .replaceAll('.markdown-source-view.mod-cm6.is-live-preview:is(', '.markdown-source-view.mod-cm6.is-live-preview:where(')
    .replaceAll(':not(.office-demo)', ':where(:not(.office-demo))');
});
views.walkDecls(decl => {
  decl.important = false;
  if (decl.prop === 'color' && decl.value === '#23408e') decl.value = 'var(--oc-accent)';
  if (decl.value.includes('color-mix(')) decl.value = decl.value.replaceAll(', white)', ', #fff)');
  if (decl.value.includes('linear-gradient(105deg, var(--oc-deep), var(--oc-accent))')) {
    decl.value = decl.value.replaceAll('linear-gradient(105deg, var(--oc-deep), var(--oc-accent))', 'linear-gradient(105deg, var(--oc-deep), var(--oc-brand))');
  }
});
views.walkRules(rule => {
  if (rule.nodes?.some(node => node.prop === '--oc-deep')) {
    const accent = rule.nodes.find(node => node.prop === '--oc-accent');
    rule.append({ prop: '--oc-brand', value: accent.value });
  }
  if (rule.selector.includes('blockquote::before') && rule.selector.includes('office-email')) {
    rule.walkDecls('background', decl => { decl.value = decl.value.replace('var(--oc-accent)', 'var(--oc-brand)'); });
  }
});

const darkColors = new Map(Object.entries({
  '#ffffff': '#1c2736', '#f6f7f9': '#243246', '#eceef2': '#141c28',
  '#e0e3e8': '#34465e', '#dde1e6': '#34465e', '#c5cad2': '#566b84',
  '#1d2127': '#e7edf5', '#565e69': '#b0bfd0', '#f1f3f5': '#243246',
  '#f8f9fb': '#243246', '#f4f6f8': '#243246', '#eef1f4': '#29384b',
  '#1a5fc2': '#8fb9ff', '#2f4fae': '#a6b7ff', '#1f7a4a': '#7dd5a0',
  '#6b3fc7': '#c6adff', '#e5eefc': '#253b59', '#e8ecf9': '#2b3553',
  '#e4f2ea': '#203e33', '#eee7fb': '#382e50', '#ffe27a': '#725b1e',
  '#374151': '#d6dfeb', '#b43e4b': '#f29aa4', '#bc3d49': '#f29aa4',
  '#c83f49': '#f29aa4', '#197347': '#7dd5a0', '#996300': '#e6c779',
  '#a77a00': '#e6c779', '#b86524': '#efb47a',
}));
const remap = text => text.replace(/fill=(["'])white\1/g, (_,quote) => `fill=${quote}#1c2736${quote}`).replace(/#[\da-f]{3,8}\b/gi, hex => {
  const lower = hex.toLowerCase();
  const full = lower.length === 4 ? '#' + [...lower.slice(1)].map(c => c + c).join('') : lower;
  const mapped = darkColors.get(full.slice(0, 7));
  return mapped ? mapped + full.slice(7) : hex;
});
function darkValue(decl, selector) {
  if (['--oc-appbar', '--oc-right', '--oc-search', '--oc-deep', '--oc-brand'].includes(decl.prop)) return decl.value;
  if (decl.prop === 'color-scheme') return 'dark';
  if (selector.includes('blockquote::before') && selector.includes('office-email')) return decl.value;
  if (decl.prop === 'color' && /(?:markdown-preview-sizer|cm-sizer)::before/.test(selector)) return decl.value;
  return remap(decl.value.replace(/url\("data:image\/svg\+xml,([^"]*)"\)/g, (_, data) =>
    `url("data:image/svg+xml,${encodeURIComponent(remap(decodeURIComponent(data)))}")`));
}
function darkRules(container) {
  const output = [];
  for (const node of container.nodes ?? []) {
    if (node.type === 'rule') {
      const changes = node.nodes.filter(decl => decl.type === 'decl')
        .map(decl => [decl, darkValue(decl, node.selector)])
        .filter(([decl, value]) => decl.value !== value);
      if (!changes.length) continue;
      const copy = node.clone({ nodes: [] });
      copy.selector = postcss.list.comma(node.selector).map(selector => '.theme-dark ' + selector).join(',\n');
      for (const [decl, value] of changes) copy.append(decl.clone({ value }));
      output.push(copy);
    } else if (node.type === 'atrule' && node.nodes) {
      const children = darkRules(node);
      if (children.length) output.push(node.clone({ nodes: children }));
    }
  }
  return output;
}
const dark = darkRules(views);
// Avoid shipping comments about retired internal packages.
views.walkComments(comment => {
  if (comment.text.includes('Office Mock CSS 0.4.0') || comment.text.includes('Recursos suministrados')) comment.remove();
});
const output = postcss.root();
output.append(postcss.comment({ text: `Office Mock ${manifest.version} | MIT | David Hurtado\nThird-party icon licenses: THIRD_PARTY_NOTICES.md. No remote resources or runtime scripts.` }));
// Community Themes downloads theme.css and manifest.json only. Keep license
// notices inside the CSS too, including when it is downloaded on its own.
output.append(postcss.comment({ text: (await read('LICENSE')).trim() + '\n\n' + (await read('THIRD_PARTY_NOTICES.md')).trim() }));
output.append(...base.nodes, ...views.nodes, postcss.comment({ text: 'Dark document surfaces and offline SVG variants.' }), ...dark);
// Independent native-view selectors have different specificity. Sort their
// rules from low to high; equal-specificity rules retain their source order.
// Keep conditional at-rules in their original positions and sort inside them.
function orderRules(container) {
  const rules = [];
  for (const node of [...container.nodes]) {
    if (node.type === 'atrule' && node.nodes) orderRules(node);
    if (node.type !== 'rule') continue;
    for (const selector of postcss.list.comma(node.selector)) {
      const rule = node.clone({ selector });
      rules.push({ rule, specificity: selectorSpecificity(selectorParser().astSync(selector)) });
    }
  }
  rules.sort((a,b) => compare(a.specificity,b.specificity));
  let next = 0;
  const nodes = [];
  for (const node of container.nodes) {
    if (node.type === 'rule') {
      const count = postcss.list.comma(node.selector).length;
      for (let i=0;i<count;i++) nodes.push(rules[next++].rule);
    } else nodes.push(node);
  }
  container.removeAll();
  container.append(...nodes);
}
orderRules(output);
output.walkRules(rule => {
  rule.raws.before = '\n\n'; rule.raws.between = ' '; rule.raws.after = '\n';
  rule.walkDecls(decl => { decl.raws.before = '\n  '; });
});
const formatted = await stylelint.lint({ code: output.toString(), codeFilename: fileURLToPath(new URL('theme.css', root)), fix: true });
await writeFile(fileURLToPath(new URL('theme.css', root)), formatted.code.trim() + '\n');
console.log(`Built Office Mock ${manifest.version}: native interface, light/dark document views, no !important.`);
