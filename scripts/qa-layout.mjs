import assert from 'node:assert/strict';
import { mkdir, writeFile } from 'node:fs/promises';
import { connect } from './cdp.mjs';

const root = new URL('../', import.meta.url);
await mkdir(new URL('.qa/', root), { recursive: true });
const c = await connect();
const run = (fn,arg) => c.evaluate('('+fn.toString()+')('+JSON.stringify(arg)+')');
const pause = ms => new Promise(resolve => setTimeout(resolve,ms));
const original = await run(() => ({
  dark: document.body.classList.contains('theme-dark'), light: document.body.classList.contains('theme-light'),
  left: app.workspace.leftSplit.collapsed, right: app.workspace.rightSplit.collapsed,
  width: innerWidth, height: innerHeight,
}));
const cases = [
  ['document','03 - Informe ejecutivo T3'], ['email','01 - Propuesta comercial'],
  ['meeting','02 - Seguimiento del proyecto Atlas'], ['spreadsheet','04 - Facturacion T3 2026'],
];
const checks = [];
function contrast(fg,bg) {
  const linear=parts=>{const c=parts.map(x=>x/255).map(x=>x<=.04045?x/12.92:((x+.055)/1.055)**2.4);return c[0]*.2126+c[1]*.7152+c[2]*.0722;};
  const a=linear(fg.match(/[\d.]+/g).slice(0,3).map(Number));
  const b=linear(bg.match(/[\da-f]{2}/gi).map(x=>parseInt(x,16)));
  return (Math.max(a,b)+.05)/(Math.min(a,b)+.05);
}
try {
  await run(() => { app.workspace.leftSplit.collapse(); app.workspace.rightSplit.collapse(); });
  for (const dark of [false,true]) {
    await run(dark => { document.body.classList.toggle('theme-dark',dark); document.body.classList.toggle('theme-light',!dark); },dark);
    for (const mode of ['preview','source']) for (const width of [1400,390,320]) {
      await c.send('Emulation.setDeviceMetricsOverride',{width,height:900,deviceScaleFactor:1,mobile:false});
      for (const [type,name] of cases) {
        await run(async({file,mode}) => {
          const leaf=globalThis.officeLayoutLeaf??app.workspace.getLeaf('tab');globalThis.officeLayoutLeaf=leaf;
          await leaf.setViewState({type:'markdown',state:{file,mode,source:false},active:true});
          await app.workspace.revealLeaf(leaf);app.workspace.setActiveLeaf(leaf,{focus:true});
        },{file:'Office Mock CSS/Ejemplos/'+name+'.md',mode});
        await pause(400);
        if (mode==='preview') for(let i=0;i<40;i++) {
          if(await run(()=>{const v=officeLayoutLeaf.view,r=v.previewMode.renderer;return r.lastText===v.data&&!r.queued&&!r.parsing;}))break;
          await pause(50);
        }
        const state=await run(mode=>{
          const v=officeLayoutLeaf.view,p=mode==='preview'?v.previewMode.containerEl.querySelector('.markdown-preview-view'):v.contentEl.querySelector('.markdown-source-view');
          const s=p.querySelector(mode==='preview'?'.markdown-preview-sizer':'.cm-sizer');
          const r=v.containerEl.getBoundingClientRect();p.scrollTop=0;
          return{overflow:p.scrollWidth-p.clientWidth,width:p.clientWidth,bar:getComputedStyle(s,'::before').color,
            ink:getComputedStyle(p).getPropertyValue('--oc-ink').trim(),surface:getComputedStyle(p).getPropertyValue('--oc-surface').trim(),
            headings:[...p.querySelectorAll('h1,h2,h3,.cm-header-1,.cm-header-2,.cm-header-3')].map(e=>getComputedStyle(e).color),
            clip:{x:r.x,y:r.y,width:r.width,height:r.height,scale:1}};
        },mode);
        assert.ok(state.width>200&&state.overflow<=1,JSON.stringify({dark,mode,width,type,...state}));
        assert.equal(state.ink,dark?'#e7edf5':'#1d2127');
        assert.equal(state.surface,dark?'#1c2736':'#fff');
        assert.equal(state.bar,'rgb(255, 255, 255)');
        for(const fg of state.headings)assert.ok(contrast(fg,dark?'#1c2736':'#ffffff')>=4.5,JSON.stringify({dark,mode,width,type,fg}));
        checks.push(`${dark?'dark':'light'} / ${mode} / ${width}px / ${type}: fits, readable surfaces and colored header`);
        if(width===1400){
          const shot=await c.send('Page.captureScreenshot',{format:'png',captureBeyondViewport:false,clip:state.clip});
          await writeFile(new URL(`images/${type}-${dark?'dark':'light'}-${mode}.png`,root),Buffer.from(shot.data,'base64'));
          if(type==='document'&&!dark&&mode==='preview'){
            const clip={...state.clip,height:state.clip.width*9/16,scale:512/state.clip.width};
            const thumb=await c.send('Page.captureScreenshot',{format:'png',captureBeyondViewport:false,clip});
            await writeFile(new URL('images/screenshot.png',root),Buffer.from(thumb.data,'base64'));
          }
        }
      }
    }
  }
} finally {
  await run(state=>{
    document.body.classList.toggle('theme-dark',state.dark);document.body.classList.toggle('theme-light',state.light);
    if(!state.left)app.workspace.leftSplit.expand();if(!state.right)app.workspace.rightSplit.expand();
  },original);
  await c.send('Emulation.setDeviceMetricsOverride',{width:original.width,height:original.height,deviceScaleFactor:1,mobile:false});
  await writeFile(new URL('.qa/layout-results.json',root),JSON.stringify({version:'1.0.0',checks},null,2)+'\n');
  c.close();
}
console.log(`${checks.length} layout checks passed; native screenshots captured.`);
