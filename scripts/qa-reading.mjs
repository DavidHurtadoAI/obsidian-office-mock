import assert from 'node:assert/strict';
import { readFile, writeFile, readdir, mkdir } from 'node:fs/promises';
import { createHash } from 'node:crypto';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { connect, devRoot } from './cdp.mjs';

const root=path.resolve(path.dirname(fileURLToPath(import.meta.url)),'..');
const project=devRoot;
const output=path.join(root,'.qa');
await mkdir(output,{recursive:true});
process.env.OBSIDIAN_CDP??='http://127.0.0.1:9235';
const c=await connect();
const checks=[];
const check=(name,ok)=>{assert.ok(ok,name);checks.push(name);console.log('PASS '+name);};
const run=(fn,arg)=>c.evaluate('('+fn.toString()+')('+JSON.stringify(arg)+')');
const pause=()=>new Promise(resolve=>setTimeout(resolve,350));
const hash=bytes=>createHash('sha256').update(bytes).digest('hex');
const notes=[];
const fixtures=[];
const fixtureFolder='Pruebas/Office Mock CSS - Predeterminado';
for(const name of await readdir(path.join(root,'tests/fixtures'))){
 fixtures.push({relative:fixtureFolder+'/'+name,source:await readFile(path.join(root,'tests/fixtures',name),'utf8')});
}
for(const [folder,target]of [['examples','Ejemplos'],['templates','Plantillas']]){
 for(const name of await readdir(path.join(root,folder))){
  const relative='Office Mock CSS/'+target+'/'+name,absolute=path.join(project,'@Nexus-dev',relative);
  const source=await readFile(absolute,'utf8');
  notes.push({relative,absolute,before:hash(source),type:source.match(/  - office-(email|meeting|document|spreadsheet)\n/)[1]});
 }
}
const original=await run(()=>({dark:document.body.classList.contains('theme-dark'),light:document.body.classList.contains('theme-light'),properties:app.vault.getConfig('propertiesInDocument')}));
async function open(file){
 await run(async file=>{
  const leaf=globalThis.officeCleanQaLeaf??app.workspace.getLeaf('tab');globalThis.officeCleanQaLeaf=leaf;
  await leaf.setViewState({type:'markdown',state:{file,mode:'preview'},active:true});await app.workspace.revealLeaf(leaf);app.workspace.setActiveLeaf(leaf,{focus:true});
  const p=leaf.view.previewMode.containerEl.querySelector('.markdown-preview-view');if(p)p.scrollTop=0;
 },file);
 await pause();
 // setViewState resuelve antes de que termine el renderizador de lectura.
 // Esperar a su contenido real evita medir clases de la nota anterior.
 for(let attempt=0;attempt<40;attempt++){
  if(await run(()=>{const v=officeCleanQaLeaf.view,r=v.previewMode.renderer;return r.lastText===v.data&&!r.parsing&&!r.queued&&r.sizerEl.innerText.trim().length>0;}))return;
  await new Promise(resolve=>setTimeout(resolve,50));
 }
 throw Error('Reading renderer did not settle for '+file);
}
async function screenshot(name){
 const dimensions=await run(()=>({width:innerWidth,height:innerHeight,scale:1/devicePixelRatio}));
 const shot=await c.send('Page.captureScreenshot',{format:'png',captureBeyondViewport:false,clip:{x:0,y:0,...dimensions}});
 await writeFile(path.join(output,name+'.png'),Buffer.from(shot.data,'base64'));
}
const styles=()=>{
 const p=officeCleanQaLeaf.view.previewMode.containerEl.querySelector('.markdown-preview-view');
 return JSON.stringify([p,...p.querySelectorAll('h1,h2,p,li,strong,blockquote,table')].slice(0,30).map(el=>{const s=getComputedStyle(el);return [el.tagName,s.color,s.backgroundColor,s.fontFamily,s.fontSize,s.margin,s.padding,s.width];}));
};
try{
 check('the canonical Obsidian workspace has no blocking dialog',await run(()=>!document.querySelector('.modal-container')));
 check('the Obsidian test window is visible for native reading rendering',await run(()=>document.visibilityState==='visible'));
 await run(async({folder,fixtures})=>{
  if(!app.vault.getAbstractFileByPath(folder))await app.vault.createFolder(folder);
  for(const fixture of fixtures){
   const existing=app.vault.getAbstractFileByPath(fixture.relative);
   if(existing){if(await app.vault.read(existing)!==fixture.source)throw Error('Refusing to overwrite a changed QA note: '+fixture.relative);}
   else await app.vault.create(fixture.relative,fixture.source);
  }
 },{folder:fixtureFolder,fixtures});
 await run(async()=>{await app.customCss.readThemes();await app.customCss.setTheme('Office Mock');await app.customCss.loadTheme();});
 await new Promise(resolve=>setTimeout(resolve,600));
 const css=await readFile(path.join(root,'theme.css'),'utf8');
 const cssState=await run(text=>{
  const sheet=new CSSStyleSheet();sheet.replaceSync(text);let count=0;const errors=[];
  const walk=rules=>{for(const rule of rules){if(rule.selectorText){count++;try{document.querySelectorAll(rule.selectorText);}catch{errors.push(rule.selectorText);}}else if(rule.cssRules)walk(rule.cssRules);}};
  walk(sheet.cssRules);return{count,errors};
 },css);
 check('the real renderer accepts every clean CSS rule',cssState.count>80&&!cssState.errors.length);
 const embedded=[...new Set([...css.matchAll(/url\("(data:[^"]+)"\)/g)].map(match=>match[1]))];
 const resources=await run(async urls=>await Promise.all(urls.map(url=>new Promise(resolve=>{const img=new Image();img.onload=()=>resolve(img.naturalWidth>0&&img.naturalHeight>0);img.onerror=()=>resolve(false);img.src=url;}))),embedded);
 check('all embedded SVG application bars and icons load offline',resources.length>15&&resources.every(Boolean));
 await c.send('Emulation.setDeviceMetricsOverride',{width:1920,height:1080,deviceScaleFactor:1,mobile:false});
 for(const dark of [false,true]){
  await run(dark=>{document.body.classList.toggle('theme-dark',dark);document.body.classList.toggle('theme-light',!dark);},dark);
  for(const note of notes){
   await open(note.relative);
   const state=await run(()=>{
    const p=officeCleanQaLeaf.view.previewMode.containerEl.querySelector('.markdown-preview-view'),s=p.querySelector('.markdown-preview-sizer');
    const visible=el=>el&&getComputedStyle(el).display!=='none'&&el.getBoundingClientRect().height>0;
    return{bg:getComputedStyle(p).backgroundColor,accent:getComputedStyle(p).getPropertyValue('--oc-accent').trim(),overflow:p.scrollWidth-p.clientWidth,hidden:[...p.querySelectorAll('.metadata-container,.frontmatter,.inline-title')].every(el=>!visible(el)),before:getComputedStyle(s,'::before').backgroundImage,toolbar:getComputedStyle(s,'::after').content,sourceUI:p.querySelectorAll('header,nav,[class^="od-"]').length,body:p.innerText,h1:p.querySelectorAll('h1').length};
   });
   check(`${dark?'dark':'light'} / ${note.relative}: CSS chrome with clean native content`,state.bg===(dark?'rgb(20, 28, 40)':'rgb(236, 238, 242)')&&state.accent===(dark?{email:'#8fb9ff',meeting:'#c6adff',document:'#a6b7ff',spreadsheet:'#7dd5a0'}:{email:'#1a5fc2',meeting:'#6b3fc7',document:'#2f4fae',spreadsheet:'#1f7a4a'})[note.type]&&state.hidden&&state.overflow<=1&&state.before.includes('data:image/svg+xml')&&state.toolbar!=='none'&&!state.sourceUI&&state.h1===1&&!state.body.includes('Responder a todos'));
   if(!dark&&note.relative.includes('/Ejemplos/'))await screenshot(note.type);
   if(['email','meeting'].includes(note.type))check(note.type+' metadata is an ordinary styled Markdown quote',await run(()=>{
    const el=officeCleanQaLeaf.view.previewMode.containerEl.querySelector('.el-h1 + .el-blockquote blockquote');
    return!!el&&el.innerText.includes(':')&&!el.querySelector('.callout');
   }));
   if(note.type==='spreadsheet')check('spreadsheet UI, letters, numbers and blank grid are generated by CSS',await run(()=>{
    const p=officeCleanQaLeaf.view.previewMode.containerEl.querySelector('.markdown-preview-view'),el=p.querySelector('.el-table'),th=p.querySelector('th'),tr=p.querySelector('tbody tr');
    return getComputedStyle(th).overflow==='visible'&&getComputedStyle(th,'::before').content.includes('oc-col')&&getComputedStyle(tr,'::before').content.includes('oc-row')&&getComputedStyle(el,'::before').backgroundImage.includes('data:image/svg+xml')&&getComputedStyle(el,'::after').backgroundImage.includes('data:image/svg+xml')&&getComputedStyle(p.querySelector('td[align="right"]')).textAlign==='right'&&[...p.querySelectorAll('tbody tr')].every(row=>row.innerText.trim());
   }));
  }
 }
 // La ausencia de propiedad es distinta de una clase ajena o un valor vacío.
 for(const mode of ['visible','hidden','source']){
  await run(mode=>app.vault.setConfig('propertiesInDocument',mode),mode);
  for(const name of ['Sin frontmatter QA.md','Otras propiedades QA.md','Sin titulo QA.md','Clase propia QA.md','Propiedad vacia QA.md','Native QA.md']){
   await open(fixtureFolder+'/'+name);
   const result=await run(()=>{
    const p=officeCleanQaLeaf.view.previewMode.containerEl.querySelector('.markdown-preview-view'),s=p.querySelector('.markdown-preview-sizer');
    const first=[...s.children].find(e=>e.className.startsWith('el-')&&!e.classList.contains('mod-frontmatter'));
    return{document:getComputedStyle(p).getPropertyValue('--oc-app-label').trim()==='"Documentos"'&&getComputedStyle(s,'::after').height==='92px',margin:first&&parseFloat(getComputedStyle(first).paddingTop),overflow:p.scrollWidth-p.clientWidth};
   });
   const expected=name!=='Native QA.md';
   check(`${mode} properties / ${name}: document default or explicit native opt-out`,result.document===expected&&result.overflow<=1&&(!expected||result.margin>=50));
   if(mode==='visible'&&name==='Sin frontmatter QA.md')await screenshot('document-default');
  }
 }
 await run(mode=>app.vault.setConfig('propertiesInDocument',mode),original.properties);
 await open(fixtureFolder+'/Nota incrustada QA.md');
 check('a default document does not add a second toolbar inside a transclusion',await run(()=>{
  const p=officeCleanQaLeaf.view.previewMode.containerEl.querySelector('.markdown-preview-view');
  const embeds=[...p.querySelectorAll('.internal-embed')];
  return embeds.length>0&&getComputedStyle(p).getPropertyValue('--oc-app-label').trim()==='"Documentos"'&&embeds.every(e=>[e,...e.querySelectorAll('.markdown-preview-view,.markdown-preview-sizer')].every(x=>!getComputedStyle(x,'::after').backgroundImage.includes('data:image/svg+xml')));
 }));
 await open('Being more flexible.md');
 check('an existing file without YAML now has document appearance without changing its contents',await run(()=>getComputedStyle(officeCleanQaLeaf.view.previewMode.containerEl.querySelector('.markdown-preview-view')).getPropertyValue('--oc-app-label').trim()==='"Documentos"'));
 await c.send('Emulation.setDeviceMetricsOverride',{width:1000,height:900,deviceScaleFactor:1,mobile:false});
 for(const note of notes.filter(n=>n.relative.includes('/Ejemplos/'))){
  await open(note.relative);
  check(note.type+' remains inside a narrow note pane',await run(()=>{const p=officeCleanQaLeaf.view.previewMode.containerEl.querySelector('.markdown-preview-view');return p.scrollWidth<=p.clientWidth+1;}));
 }
 await screenshot('spreadsheet-narrow');
 check('wide data scrolls horizontally within the table',await run(()=>{const t=officeCleanQaLeaf.view.previewMode.containerEl.querySelector('.el-table');t.scrollLeft=140;return t.scrollWidth>t.clientWidth&&t.scrollLeft>0;}));
 await open('Office Mock CSS/Ejemplos/01 - Propuesta comercial.md');
 await run(async()=>{await officeCleanQaLeaf.setViewState({type:'markdown',state:{file:'Office Mock CSS/Ejemplos/01 - Propuesta comercial.md',mode:'source',source:true},active:true});});
 await pause();
 check('source mode contains only frontmatter and content, with no office UI styling',await run(()=>{const source=officeCleanQaLeaf.view.contentEl.querySelector('.markdown-source-view');return officeCleanQaLeaf.view.getMode()==='source'&&officeCleanQaLeaf.view.editor.getValue().startsWith('---\ncssclasses:')&&!officeCleanQaLeaf.view.editor.getValue().includes('<header')&&getComputedStyle(source).getPropertyValue('--oc-accent').trim()==='';}));
 await open(fixtureFolder+'/Sin frontmatter QA.md');
 await run(async()=>{const state=officeCleanQaLeaf.getViewState();await officeCleanQaLeaf.setViewState({...state,state:{...state.state,mode:'source',source:true}});});
 await pause();
 check('a file without cssclasses keeps its normal source editor',await run(()=>officeCleanQaLeaf.view.getMode()==='source'&&getComputedStyle(officeCleanQaLeaf.view.contentEl.querySelector('.markdown-source-view')).getPropertyValue('--oc-accent').trim()===''));
 check('the standalone theme is selected with both legacy snippets disabled',await run(async()=>{
  const a=JSON.parse(await app.vault.adapter.read('.obsidian/appearance.json'));
  return a.cssTheme==='Office Mock'&&!a.enabledCssSnippets.some(s=>['office-clean','office-demo'].includes(s))&&await app.vault.adapter.exists('.obsidian/themes/Office Mock/theme.css');
 }));
 await open(fixtureFolder+'/Native QA.md');
 check('office-plain keeps a native layout without a generated ribbon',await run(()=>{
  const p=officeCleanQaLeaf.view.previewMode.containerEl.querySelector('.markdown-preview-view'),s=p.querySelector('.markdown-preview-sizer');
  return getComputedStyle(p).getPropertyValue('--oc-accent').trim()===''&&!getComputedStyle(s,'::before').backgroundImage.includes('data:image/svg+xml');
 }));
 await open('Office Mock CSS/Ejemplos/01 - Propuesta comercial.md');
 await run(async()=>{await app.customCss.setTheme('');await app.customCss.loadTheme();});
 check('switching to the default theme removes the bars and keeps the message content',await run(()=>{const p=officeCleanQaLeaf.view.previewMode.containerEl.querySelector('.markdown-preview-view'),s=p.querySelector('.markdown-preview-sizer');return !getComputedStyle(s,'::before').backgroundImage.includes('data:image/svg+xml')&&p.innerText.includes('Hola, Pablo:')&&p.innerText.includes('Laura Martín');}));
 await run(async()=>{await app.customCss.setTheme('Office Mock');await app.customCss.loadTheme();});
 for(const note of notes)check('source bytes preserved: '+note.relative,hash(await readFile(note.absolute))===note.before);
 for(const fixture of fixtures)check('QA source bytes preserved: '+fixture.relative,hash(await readFile(path.join(project,'@Nexus-dev',fixture.relative)))===hash(fixture.source));
}finally{
 await run(async state=>{document.body.classList.toggle('theme-dark',state.dark);document.body.classList.toggle('theme-light',state.light);app.vault.setConfig('propertiesInDocument',state.properties);await app.customCss.setTheme('Office Mock');await app.customCss.loadTheme();},original);
 await c.send('Emulation.setDeviceMetricsOverride',{width:1920,height:1080,deviceScaleFactor:1,mobile:false});
 await open('Office Mock CSS/Ejemplos/01 - Propuesta comercial.md');
 await writeFile(path.join(output,'results.json'),JSON.stringify({version:'1.0.0',date:new Date().toISOString(),checks,viewport:[1920,1080]},null,2)+'\n');
 c.close();
}
console.log(`\n${checks.length} checks passed.`);
