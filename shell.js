import {brand} from './brand.js';
import {releaseState} from './release.js';
export const base=new URL('.',import.meta.url);
const route=document.body.dataset.route, key='local-suite.accepted-terms.'+base.pathname;
const root=document.querySelector('#root'), local=['localhost','127.0.0.1','[::1]'].includes(location.hostname);
const pagesTest = location.protocol === 'https:' && location.hostname.endsWith('.github.io')
  && new URLSearchParams(location.search).get('preview') === '1';

function routeURL(relative = '.') {
  const url = new URL(relative, base);
  if (pagesTest) url.searchParams.set('preview', '1');
  return url;
}

const esc=s=>String(s).replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const paths={spreadsheets:'M3 4h18v16H3zM3 9h18M3 14h18M9 9v11M15 9v11',documents:'M5 3h10l4 4v14H5zM14 3v5h5M8 12h8M8 16h7',presentations:'M2 3h20M4 4v12h16V4M12 16v4m-4 2 4-3 4 3',lock:'M6 10h12v10H6zM8 10V7a4 4 0 0 1 8 0v3M12 14v3',check:'m4 12 5 5L20 6',info:'M12 2a10 10 0 1 0 0 20 10 10 0 0 0 0-20M12 11v6M12 7v1'};
export function icon(which){return '<svg class="icon" aria-hidden="true" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linejoin="round" stroke-linecap="round"><path d="'+paths[which]+'"/></svg>'}
for(const [k,v] of Object.entries(brand.colors)) document.documentElement.style.setProperty('--'+k,v);
document.title=brand.pageTitle(route);
document.querySelector('meta[name=description]').content=brand.description;
const favicon=document.createElement('link');favicon.rel='icon';favicon.href=new URL(brand.mark,base);document.head.append(favicon);
let terms='',draft=true,accepted=false,stopArt,termsHash='';
const receipt=()=>JSON.stringify({version:brand.termsVersion,hash:termsHash});
export const devNotice=()=>'<aside class="strip">'+icon('info')+'<span><strong>'+esc(brand.stage)+'</strong> '+esc(brand.warning)+'</span></aside>';
const art=()=>'<div class="artcard"><canvas aria-hidden="true"></canvas><span class="artlabel">'+esc(brand.descriptor)+'</span><div><div class="arttitle">'+esc(brand.suite)+'</div><div class="artcaption"><span>WORKSPACE EDITION</span><span>V. '+esc(brand.version)+'</span></div></div></div>';
const termsHTML=()=>terms.trim().split(/\r?\n\s*\r?\n/).map(s=>s.startsWith('## ')?'<h4>'+esc(s.slice(3))+'</h4>':s.startsWith('# ')?'<h3>'+esc(s.slice(2))+'</h3>':'<p>'+esc(s)+'</p>').join('')+'<p>Acceptance applies to this official build or service. It does not limit rights granted by the MIT source-code license or any separate component license.</p>'+(brand.approvedDisclaimers?'<p>'+esc(brand.approvedDisclaimers)+'</p>':'');
function frame(body){
  stopArt?.();
  const mail=(sub)=>'mailto:'+brand.feedback+'?subject='+encodeURIComponent(brand.suite+' '+sub);
  root.innerHTML='<header class="sitehead"><div class="wrap headerinner"><a class="brand" href="'+routeURL().href+'"><img alt="" src="'+new URL(brand.mark,base)+'">'+esc(brand.suite)+'</a><div class="headerend"><button class="plainlink termslink">Terms of Service</button><span class="metatag"><span class="dot"></span>'+esc(brand.stage)+'</span></div></div></header><main class="wrap" tabindex="-1" id="main">'+body+'</main><footer class="wrap foot"><div class="footleft"><span class="footerbrand">'+esc(brand.suite)+'</span><span> / </span><span>Local workspace</span></div><nav aria-label="Support"><button class="plainlink termslink">Terms</button><a href="'+mail('bug report')+'">Report a bug</a><a href="'+mail('feature request')+'">Request a feature</a></nav></footer><dialog aria-labelledby="terms-title" id="termsdialog"><div class="panelhead"><h2 id="terms-title">Terms of Service</h2><span class="draft">'+(draft?'Draft':'Approved')+'</span></div><div class="scrollterms" tabindex="0" aria-label="Agreement text">'+termsHTML()+'</div><div class="actions"><button class="secondary close">Close</button></div></dialog>';
  const dlg=root.querySelector('#termsdialog');root.querySelectorAll('.termslink').forEach(b=>b.onclick=()=>dlg.showModal());dlg.querySelector('.close').onclick=()=>dlg.close();
  const canvas=root.querySelector('canvas');if(canvas)import(new URL(brand.art,base)).then(m=>{if(canvas.isConnected)stopArt=m.illustrate(canvas)});
}
function gate(){
  frame('<div class="intro"><div class="eyebrow">Workspace access</div><h1>Welcome to '+esc(brand.suite)+'.</h1><p class="lead">Review the agreement to enter your local workspace.</p></div><div class="gategrid">'+art()+'<section class="panel"><div class="panelhead"><h2>Terms of Service</h2><span class="draft">'+(draft?'Legal review · Draft':esc(brand.termsVersion))+'</span></div><div class="scrollterms" tabindex="0" role="region" aria-label="Terms of Service text">'+termsHTML()+'</div><button id="agree" class="primary agree">'+icon('check')+'I agree to the Terms of Service</button><button id="decline" class="secondary">Decline</button><p class="hint" id="gatehint" role="status">Acceptance is saved on this device.</p></section></div>'+devNotice());
  root.querySelector('#decline').onclick=()=>{frame('<div class="intro"><h1>Terms declined</h1><p>Your files and saved data have not been changed. No acceptance was saved.</p><button id="reviewterms" class="secondary">Review terms again</button></div>');root.querySelector('#reviewterms').onclick=gate;};
  root.querySelector('#agree').onclick=()=>{try{localStorage.setItem(key,receipt());accepted=true;enter()}catch{root.querySelector('#gatehint').textContent='Browser storage is unavailable. Allow local storage to remember your agreement.'}};
}
function launcher(){
  const cards=Object.entries(brand.apps).map(([k,a])=>'<article class="appcard"><div class="cardtop"><span class="appicon '+k+'">'+icon(k)+'</span><span class="pill '+(a.available?'avail':'')+'">'+(a.available?'Available':'Planned')+'</span></div><h3>'+esc(a.name)+'</h3><p>'+esc(a.subtitle)+'</p><span class="format">'+esc(a.formats)+'</span>'+(a.available?'<a class="appopen" href="'+routeURL(k+'/')+'">Open '+esc(a.name)+'</a>':'')+'</article>').join('');
  frame('<section class="launchhero"><div><div class="eyebrow">Suite launcher</div><h1>'+esc(brand.fullName)+'</h1><p class="lead">Choose an app to view files on this device. This preview does not edit or upload files.</p></div>'+art()+'</section><section><div class="sectionhead"><h2>Applications</h2><span class="session">Local session</span></div><div class="apps">'+cards+'</div></section>'+devNotice());
}
const viewerRoutes = {
  spreadsheets: {load: () => import('./spreadsheets/viewer.js'), stylesheet: 'spreadsheets/viewer.css', entry: 'mount'},
  documents: {load: () => import('./documents/viewer.js'), stylesheet: 'documents/viewer.css', entry: 'mount'},
  presentations: {load: () => import('./presentations/viewer.js'), stylesheet: 'presentations/viewer.css', entry: 'mountViewer'}
};
function loadViewerStyles(path) {
  const href = new URL(path, base).href;
  const existing = [...document.querySelectorAll('link[rel=stylesheet]')].find(link => link.href === href);
  const link = existing || Object.assign(document.createElement('link'), {rel: 'stylesheet', href});
  if (link.sheet) return Promise.resolve();
  return new Promise((resolve, reject) => {
    link.addEventListener('load', resolve, {once:true});
    link.addEventListener('error', () => reject(Error('Viewer styling could not load.')), {once:true});
    if (!existing) document.head.append(link);
  });
}
async function enter(){
  if(!accepted){gate();return}
  if(!route){launcher();return}
  const config=viewerRoutes[route];
  if(!config){launcher();return}
  frame('<div class="intro"><div class="eyebrow">Read-only local viewer</div><h1>'+esc(brand.apps[route].name)+'</h1><p class="lead">Open and view a file from this device. Your original file is never modified.</p></div><div id="viewer"><p class="fail" role="status">Preparing the viewer…</p></div>'+devNotice());
  const container=root.querySelector('#viewer');
  try {
    const [viewer]=await Promise.all([config.load(),loadViewerStyles(config.stylesheet)]);
    await viewer[config.entry](container,{brand,base:routeURL()});
  } catch (error) {
    container.textContent='The viewer could not start. '+error.message;
  }
}
async function start(){
  const res=await fetch(new URL('tos.md',base),{cache:'no-store'});if(!res.ok)throw Error('The local agreement could not be loaded.');
  terms=await res.text();if(!terms.trim())throw Error('The agreement file is empty.');
  const release=await releaseState(brand,terms);
  termsHash=release.digest;
  draft=!release.termsApproved;
  if(!local&&!pagesTest&&!release.ready){frame('<div class="intro"><div class="eyebrow">Release on hold</div><h1>Local preview only.</h1><p class="lead">Public access will open after release review.</p></div>');return}
  try{accepted=localStorage.getItem(key)===receipt()}catch{}
  await enter();
}
addEventListener('pageshow',e=>{if(e.persisted)location.reload()});
addEventListener('storage',e=>{if(e.key===key||e.key===null){if(localStorage.getItem(key)!==receipt())location.reload()}});
start().catch(e=>{root.textContent='Workspace unavailable. '+e.message});
