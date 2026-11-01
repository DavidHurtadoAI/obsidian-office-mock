import assert from 'node:assert/strict';
import {readFile,writeFile,readdir,mkdir} from 'node:fs/promises';
import path from 'node:path';
import{fileURLToPath}from'node:url';
import{connect,devRoot}from'./cdp.mjs';

const root=path.resolve(path.dirname(fileURLToPath(import.meta.url)),'..');
const project=devRoot;
const output=path.join(root,'.qa');await mkdir(output,{recursive:true});
process.env.OBSIDIAN_CDP??='http://127.0.0.1:9235';
const c=await connect();
const checks=[];
const check=(name,ok)=>{assert.ok(ok,name);checks.push(name);console.log('PASS '+name);};
const run=(fn,arg)=>c.evaluate('('+fn.toString()+')('+JSON.stringify(arg)+')');
const pause=ms=>new Promise(r=>setTimeout(r,ms??350));
const folder='Pruebas/Office Mock CSS - Edicion funcional';
const fixtures=[];
for(const name of await readdir(path.join(root,'tests/editing-fixtures')))fixtures.push({file:folder+'/'+name,source:await readFile(path.join(root,'tests/editing-fixtures',name),'utf8')});
const original=await run(()=>({dark:document.body.classList.contains('theme-dark'),light:document.body.classList.contains('theme-light'),fontProps:['--font-text-override','--font-interface-override','--font-monospace-override'].map(key=>({key,value:document.body.style.getPropertyValue(key),priority:document.body.style.getPropertyPriority(key)}))}));
async function open(file,mode='source',source=false){
 await run(async({file,mode,source})=>{
  const leaf=globalThis.officeEditingQaLeaf??app.workspace.getLeaf('tab');globalThis.officeEditingQaLeaf=leaf;
  await leaf.setViewState({type:'markdown',state:{file,mode,source},active:true});await app.workspace.revealLeaf(leaf);app.workspace.setActiveLeaf(leaf,{focus:true});
 },{file,mode,source});await pause(450);
}
async function key(key,code,virtual,modifiers=0){
 await c.send('Input.dispatchKeyEvent',{type:'keyDown',key,code,windowsVirtualKeyCode:virtual,modifiers});
 await c.send('Input.dispatchKeyEvent',{type:'keyUp',key,code,windowsVirtualKeyCode:virtual,modifiers:0});
}
async function screenshot(name){
 const dimensions=await run(()=>({width:innerWidth,height:innerHeight,scale:1/devicePixelRatio}));
 const shot=await c.send('Page.captureScreenshot',{format:'png',captureBeyondViewport:false,clip:{x:0,y:0,...dimensions}});
 await writeFile(path.join(output,name+'.png'),Buffer.from(shot.data,'base64'));
}
async function selectCell(row,col){
 const point=await run(({row,col})=>{
  const p=officeEditingQaLeaf.view.contentEl.querySelector('.markdown-source-view');
  const cell=p.querySelector('table.table-editor tbody').rows[row].cells[col].querySelector('.table-cell-wrapper');
  cell.scrollIntoView({block:'center',inline:'nearest'});const r=cell.getBoundingClientRect();
  return{x:r.x+Math.min(30,r.width/2),y:r.y+r.height/2};
 },{row,col});
 await c.send('Input.dispatchMouseEvent',{type:'mousePressed',button:'left',clickCount:1,...point});
 await c.send('Input.dispatchMouseEvent',{type:'mouseReleased',button:'left',clickCount:1,...point});await pause(200);
}
const activeCell=()=>{
 const active=document.activeElement;
 if(!officeEditingQaLeaf.view.contentEl.contains(active))return null;
 const cell=active.closest('td,th');
 return cell?{col:cell.cellIndex,row:cell.parentElement.sectionRowIndex,text:active.textContent}:null;
};
try{
 await run(async({folder,fixtures})=>{
  globalThis.officeEditingPrepared=[];
  if(!app.vault.getAbstractFileByPath(folder))await app.vault.createFolder(folder);
  for(const f of fixtures){const old=app.vault.getAbstractFileByPath(f.file);if(old&&await app.vault.read(old)!==f.source)throw Error('Refusing to overwrite changed QA note: '+f.file);}
  for(const f of fixtures){if(!app.vault.getAbstractFileByPath(f.file))await app.vault.create(f.file,f.source);officeEditingPrepared.push(f.file);}
  await app.customCss.readThemes();await app.customCss.setTheme('Office Mock');await app.customCss.loadTheme();
 },{folder,fixtures});await pause(600);
 await c.send('Emulation.setDeviceMetricsOverride',{width:1920,height:1080,deviceScaleFactor:1,mobile:false});
 const cases=[['email','01 - Propuesta comercial'],['meeting','02 - Seguimiento del proyecto Atlas'],['document','03 - Informe ejecutivo T3'],['spreadsheet','04 - Facturacion T3 2026']];
 for(const dark of[false,true]){
  await run(dark=>{document.body.classList.toggle('theme-dark',dark);document.body.classList.toggle('theme-light',!dark);},dark);
  for(const[type,name]of cases){
   await open('Office Mock CSS/Ejemplos/'+name+'.md');
   const result=await run(()=>{const p=officeEditingQaLeaf.view.contentEl.querySelector('.markdown-source-view'),s=p.querySelector('.cm-sizer');return{live:p.classList.contains('is-live-preview'),accent:getComputedStyle(p).getPropertyValue('--oc-accent').trim(),app:getComputedStyle(s,'::before').content,overflow:p.scrollWidth-p.clientWidth,editor:!!p.querySelector('.cm-content[contenteditable="true"]')};});
   check(`${dark?'dark':'light'} / ${type}: Office UI around the native editable document`,result.live&&result.accent===(dark?{email:'#8fb9ff',meeting:'#c6adff',document:'#a6b7ff',spreadsheet:'#7dd5a0'}:{email:'#1a5fc2',meeting:'#6b3fc7',document:'#2f4fae',spreadsheet:'#1f7a4a'})[type]&&result.app!=='none'&&result.overflow<=1&&result.editor);
   if(!dark)await screenshot(type+'-editing');
  }
 }
 await open('Being more flexible.md');
 check('an existing file without cssclasses has the Word page in Live Preview',await run(()=>getComputedStyle(officeEditingQaLeaf.view.contentEl.querySelector('.markdown-source-view')).getPropertyValue('--oc-app-label').trim()==='"Documentos"'));
 // Simula las tres preferencias de Apariencia sin escribir configuración de usuario.
 await run(()=>{document.body.style.setProperty('--font-text-override','Georgia');document.body.style.setProperty('--font-interface-override','Trebuchet MS');document.body.style.setProperty('--font-monospace-override','Consolas');});
 for(const mode of['preview','source']){
  await open('Office Mock CSS/Ejemplos/03 - Informe ejecutivo T3.md',mode);
  check(mode+' honors the configured text and interface fonts',await run(mode=>{
   const p=mode==='preview'?officeEditingQaLeaf.view.previewMode.containerEl.querySelector('.markdown-preview-view'):officeEditingQaLeaf.view.contentEl.querySelector('.markdown-source-view');
   const s=p.querySelector(mode==='preview'?'.markdown-preview-sizer':'.cm-sizer');
   const text=p.querySelector(mode==='preview'?'p':'.cm-content');
   return getComputedStyle(text).fontFamily.includes('Georgia')&&getComputedStyle(s,'::before').fontFamily.includes('Trebuchet MS');
  },mode));
 }
 await open(folder+'/Texto QA.md');
 await run(()=>{const e=officeEditingQaLeaf.view.editor;e.setCursor(e.offsetToPos(e.getValue().length));e.focus();});
 check('the native document editor receives focus',await run(()=>officeEditingQaLeaf.view.contentEl.contains(document.activeElement)&&document.activeElement.classList.contains('cm-content')));
 await c.send('Input.insertText',{text:'\n\nTexto escrito durante la prueba.'});await pause(200);
 await run(()=>officeEditingQaLeaf.view.save());
 check('typing in the Word page saves ordinary Markdown', (await readFile(path.join(project,'@Nexus-dev',folder,'Texto QA.md'),'utf8')).includes('Texto escrito durante la prueba.'));
 await key('z','KeyZ',90,2);await pause(200);await run(()=>officeEditingQaLeaf.view.save());
 check('native undo restores the original document',await run(original=>officeEditingQaLeaf.view.editor.getValue()===original,fixtures.find(f=>f.file.endsWith('Texto QA.md')).source));
 await open(folder+'/Celdas QA.md');await selectCell(0,0);
 const first=await run(activeCell);check('clicking a spreadsheet cell opens its native cell editor',first?.col===0&&first?.row===0&&first?.text==='Alfa');
 check('the active native cell uses the configured text font',await run(()=>getComputedStyle(document.activeElement).fontFamily.includes('Georgia')));
 await key('a','KeyA',65,2);await c.send('Input.insertText',{text:'Alfa QA'});
 await key('Tab','Tab',9);await pause(200);
 const second=await run(activeCell);check('Tab commits the value and moves to the next real cell',second?.col===1&&second?.row===0);
 await key('a','KeyA',65,2);await c.send('Input.insertText',{text:'125,00'});await key('Tab','Tab',9);await pause(200);
 check('typing a second cell preserves the rest of the table',await run(()=>{const text=officeEditingQaLeaf.view.editor.getValue();return text.includes('Alfa QA')&&text.includes('125,00')&&text.includes('Beta')&&text.includes('200,00')&&text.includes('Pendiente');}));
 await key('Tab','Tab',9,8);await pause(200);const back=await run(activeCell);check('Shift+Tab moves back to the preceding cell',back?.col===1&&back?.row===0);
 await key('Tab','Tab',9);await pause(200);await run(()=>officeEditingQaLeaf.view.save());
 const saved=await readFile(path.join(project,'@Nexus-dev',folder,'Celdas QA.md'),'utf8');
 check('edited cells persist in the Markdown file on disk',saved.includes('Alfa QA')&&saved.includes('125,00')&&saved.startsWith('---\ncssclasses:'));
 await screenshot('spreadsheet-cell-editing');
 await open(folder+'/Celdas QA.md','preview');
 check('reading view displays the persisted cell changes',await run(()=>{const p=officeEditingQaLeaf.view.previewMode.containerEl;return p.innerText.includes('Alfa QA')&&p.innerText.includes('125,00');}));
 await open(folder+'/Celdas QA.md','source',true);
 check('full source mode remains available without decorative UI',await run(()=>{const p=officeEditingQaLeaf.view.contentEl.querySelector('.markdown-source-view');return !p.classList.contains('is-live-preview')&&getComputedStyle(p).getPropertyValue('--oc-accent').trim()==='';}));
 await c.send('Emulation.setDeviceMetricsOverride',{width:1000,height:900,deviceScaleFactor:1,mobile:false});
 await open('Office Mock CSS/Ejemplos/04 - Facturacion T3 2026.md');
 check('wide editable sheets scroll inside the native table wrapper',await run(()=>{const p=officeEditingQaLeaf.view.contentEl.querySelector('.markdown-source-view'),t=p.querySelector('.cm-table-widget > .table-wrapper');t.scrollLeft=120;return t.scrollWidth>t.clientWidth&&t.scrollLeft>0&&p.scrollWidth<=p.clientWidth+1;}));
 await screenshot('spreadsheet-editing-narrow');
}finally{
 // Desmonta los editores de prueba antes de restaurar únicamente sus fixtures.
 const restored=await run(async({fixtures,state})=>{
  const prepared=globalThis.officeEditingPrepared??[];
  if(prepared.length&&globalThis.officeEditingQaLeaf){await officeEditingQaLeaf.view.save?.();await officeEditingQaLeaf.setViewState({type:'empty',state:{}});}
  for(const f of fixtures){const file=app.vault.getAbstractFileByPath(f.file);if(file&&prepared.includes(f.file))await app.vault.modify(file,f.source);}
  document.body.classList.toggle('theme-dark',state.dark);document.body.classList.toggle('theme-light',state.light);
  for(const p of state.fontProps){if(p.value)document.body.style.setProperty(p.key,p.value,p.priority);else document.body.style.removeProperty(p.key);}
  return prepared;
 },{fixtures,state:original});
 for(const f of fixtures.filter(f=>restored.includes(f.file)))assert.equal(await readFile(path.join(project,'@Nexus-dev',f.file),'utf8'),f.source,'QA fixture restored: '+f.file);
 await writeFile(path.join(output,'editing-results.json'),JSON.stringify({version:'1.0.0',date:new Date().toISOString(),checks},null,2)+'\n');
 c.close();
}
console.log(`\n${checks.length} editing checks passed.`);
