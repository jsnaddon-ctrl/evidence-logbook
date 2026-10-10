/* Evidence Logbook — hosted version
 * Storage: the apprentice's own OneDrive (Microsoft Graph)
 * AI: Cloudflare Worker -> Anthropic API (key stays on the Worker)
 * The app plans and files evidence. It never writes notes, answers or evidence for the apprentice.
 */
"use strict";
const APP_VERSION = '1.1.0';
const CFG = Object.assign({rootFolder:'Apprenticeship Evidence', redirectUri:''}, window.APP_CONFIG || {});
const DEMO = /[?&]demo=1\b/.test(location.search) || !CFG.clientId || CFG.clientId.startsWith('PASTE');
const ROOT = CFG.rootFolder;
const DATA_PATH = ROOT + '/App data - do not edit/logbook.json';
const GRAPH = 'https://graph.microsoft.com/v1.0';
const SCOPES = ['User.Read', 'Files.ReadWrite'];

const $app = document.getElementById('app');
const STATUS = {new:'Not started', collecting:'Collecting evidence', ready:'Ready for assessor', signed:'Signed off', awarded:'Credits awarded'};
const freshCap = () => ({photos:[], note:'', job:'', unitId:'', itemId:'', sugg:null, busy:false, msg:'', reading:false});
const S = {
  phase:'start', err:'', account:null, units:{}, ev:{}, profile:{},
  view:'units', unitId:null, cap:freshCap(),
  add:{busy:false, msg:'', review:null},
  plan:{q:'', busy:false, res:null, msg:''},
  editing:null, modal:null, toast:'', packBusy:false, saving:false, pf:null
};
let addCtl = null, lastView = '';

if (window.pdfjsLib) pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';

/* ================= helpers ================= */
const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const rid = () => Math.random().toString(36).slice(2,10) + Date.now().toString(36).slice(-4);
const fmtDate = iso => { try { return new Date(iso).toLocaleDateString('en-NZ',{day:'numeric',month:'short',year:'numeric'}); } catch { return ''; } };
const ymd = iso => { const d = new Date(iso); return isNaN(d) ? '' : d.getFullYear() + '-' + String(d.getMonth()+1).padStart(2,'0') + '-' + String(d.getDate()).padStart(2,'0'); };
const unitLabel = u => [u.number, u.title].filter(Boolean).join(' – ') || 'Untitled unit';
const num = v => { const m = String(v||'').match(/\d+(\.\d+)?/); return m ? parseFloat(m[0]) : 0; };
const safeName = s => String(s||'').replace(/[\\/:*?"<>|#%]/g,'-').replace(/\s+/g,' ').trim().replace(/[. ]+$/,'').slice(0,80) || 'Untitled';
const isPdf = f => f.type === 'application/pdf' || /\.pdf$/i.test(f.name||'');
function toast(msg){ S.toast = msg; render(); clearTimeout(toast.t); toast.t = setTimeout(()=>{ S.toast=''; render(); }, 3500); }
function setPath(obj, path, val){ const ks = path.split('.'); let o = obj; for (let i=0;i<ks.length-1;i++){ o = o[/^\d+$/.test(ks[i])?+ks[i]:ks[i]]; if (o == null) return; } const k = ks[ks.length-1]; o[/^\d+$/.test(k)?+k:k] = val; }
function getPath(obj, path){ return path.split('.').reduce((o,k)=> o == null ? o : o[/^\d+$/.test(k)?+k:k], obj); }
function blobToBase64(b){ return new Promise((res,rej)=>{ const r = new FileReader(); r.onload = () => res(String(r.result).split(',')[1]); r.onerror = () => rej(new Error('Could not read file')); r.readAsDataURL(b); }); }
function localDownload(bytes, filename){
  const url = URL.createObjectURL(new Blob([bytes], {type:'application/pdf'}));
  const a = document.createElement('a'); a.href = url; a.download = filename; document.body.appendChild(a); a.click(); a.remove();
  setTimeout(()=>URL.revokeObjectURL(url), 60000);
}

/* ================= Microsoft sign-in ================= */
let msalApp = null;
const redirectUri = () => CFG.redirectUri || (location.origin + location.pathname.replace(/index\.html$/,''));
async function authInit(){
  msalApp = new msal.PublicClientApplication({
    auth:{clientId:CFG.clientId, authority:'https://login.microsoftonline.com/common', redirectUri:redirectUri()},
    cache:{cacheLocation:'localStorage'}
  });
  await msalApp.initialize();
  const res = await msalApp.handleRedirectPromise();
  const acct = (res && res.account) || msalApp.getAllAccounts()[0] || null;
  if (acct) msalApp.setActiveAccount(acct);
  return acct;
}
async function getToken(){
  const account = msalApp.getActiveAccount();
  try { return (await msalApp.acquireTokenSilent({scopes:SCOPES, account})).accessToken; }
  catch (e){
    if (e instanceof msal.InteractionRequiredAuthError){ await msalApp.acquireTokenRedirect({scopes:SCOPES, account}); }
    throw e;
  }
}

/* ================= OneDrive store ================= */
const encPath = p => p.split('/').map(encodeURIComponent).join('/');
async function gfetch(path, {method='GET', headers={}, body}={}){
  const token = await getToken();
  const r = await fetch(path.startsWith('http') ? path : GRAPH + path, {method, headers:{Authorization:'Bearer ' + token, ...headers}, body});
  if (!r.ok){ const err = new Error('OneDrive error ' + r.status); err.status = r.status; try { err.detail = (await r.json()).error; } catch {} throw err; }
  if (r.status === 204) return null;
  const ct = r.headers.get('content-type') || '';
  return ct.includes('json') ? r.json() : r;
}
const emptyData = () => ({version:1, profile:{}, units:{}, ev:{}});

const GraphStore = {
  data: emptyData(), etag: null, thumbs: new Map(),
  async load(){
    try {
      const item = await gfetch(`/me/drive/root:/${encPath(DATA_PATH)}`);
      const txt = await (await fetch(item['@microsoft.graph.downloadUrl'])).text();
      this.data = Object.assign(emptyData(), JSON.parse(txt)); this.etag = item.eTag;
    } catch (e){
      if (e.status === 404){ this.data = emptyData(); this.etag = null; } else throw e;
    }
  },
  async _save(d){
    const q = this.etag ? '' : '?@microsoft.graph.conflictBehavior=fail';
    const headers = {'Content-Type':'application/json'}; if (this.etag) headers['If-Match'] = this.etag;
    const item = await gfetch(`/me/drive/root:/${encPath(DATA_PATH)}:/content${q}`, {method:'PUT', headers, body:JSON.stringify(d)});
    this.etag = item.eTag;
  },
  async mutate(fn){
    for (let attempt=0; attempt<3; attempt++){
      const draft = JSON.parse(JSON.stringify(this.data)); fn(draft);
      try { await this._save(draft); this.data = draft; return; }
      catch (e){ if (e.status === 412 || e.status === 409){ await this.load(); continue; } throw e; }
    }
    throw new Error('Your logbook changed on another device. Try again.');
  },
  async upload(relPath, blob, conflict='rename'){
    const enc = encPath(ROOT + '/' + relPath);
    if (blob.size <= 4*1024*1024){
      const it = await gfetch(`/me/drive/root:/${enc}:/content?@microsoft.graph.conflictBehavior=${conflict}`, {method:'PUT', headers:{'Content-Type':blob.type||'application/octet-stream'}, body:blob});
      return {id:it.id, name:it.name};
    }
    const sess = await gfetch(`/me/drive/root:/${enc}:/createUploadSession`, {method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({item:{'@microsoft.graph.conflictBehavior':conflict}})});
    const CH = 320*1024*10; let start = 0, last = null;
    while (start < blob.size){
      const end = Math.min(start + CH, blob.size);
      const r = await fetch(sess.uploadUrl, {method:'PUT', headers:{'Content-Range':`bytes ${start}-${end-1}/${blob.size}`}, body:blob.slice(start, end)});
      if (!r.ok){ const err = new Error('Upload failed ' + r.status); err.status = r.status; throw err; }
      last = r; start = end;
    }
    const it = await last.json(); return {id:it.id, name:it.name};
  },
  async ensureFolderAbs(full){
    try { return await gfetch(`/me/drive/root:/${encPath(full)}`); } catch (e){ if (e.status !== 404) throw e; }
    const parts = full.split('/'); const name = parts.pop(); const parent = parts.join('/');
    if (parent) await this.ensureFolderAbs(parent);
    const url = parent ? `/me/drive/root:/${encPath(parent)}:/children` : '/me/drive/root/children';
    try { return await gfetch(url, {method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({name, folder:{}, '@microsoft.graph.conflictBehavior':'fail'})}); }
    catch (e){ if (e.status === 409) return await gfetch(`/me/drive/root:/${encPath(full)}`); throw e; }
  },
  async move(id, folderRel){
    const f = await this.ensureFolderAbs(ROOT + '/' + folderRel);
    await gfetch(`/me/drive/items/${id}?@microsoft.graph.conflictBehavior=rename`, {method:'PATCH', headers:{'Content-Type':'application/json'}, body:JSON.stringify({parentReference:{id:f.id}})});
  },
  async thumbUrl(id){
    if (!this.thumbs.has(id)) this.thumbs.set(id, (async () => {
      try { const r = await gfetch(`/me/drive/items/${id}/thumbnails`); const t = r.value && r.value[0]; if (t && t.medium) return t.medium.url; } catch {}
      const it = await gfetch(`/me/drive/items/${id}`); return it['@microsoft.graph.downloadUrl'];
    })().catch(e => { this.thumbs.delete(id); throw e; }));
    return this.thumbs.get(id);
  },
  async bytes(id){ const it = await gfetch(`/me/drive/items/${id}`); return await (await fetch(it['@microsoft.graph.downloadUrl'])).arrayBuffer(); },
  async webUrl(id){ return (await gfetch(`/me/drive/items/${id}?select=webUrl`)).webUrl; },
  async rootUrl(){ return (await this.ensureFolderAbs(ROOT)).webUrl; },
  async remove(id){ try { await gfetch(`/me/drive/items/${id}`, {method:'DELETE'}); } catch (e){ if (e.status !== 404) throw e; } },
  async removePath(rel){ try { const it = await gfetch(`/me/drive/root:/${encPath(ROOT + '/' + rel)}`); await this.remove(it.id); } catch (e){ if (e.status !== 404) throw e; } }
};

/* Demo store: everything in memory, for trying the app before setup */
const DemoStore = {
  data: emptyData(), files: new Map(),
  async load(){}, async mutate(fn){ const d = JSON.parse(JSON.stringify(this.data)); fn(d); this.data = d; },
  async upload(relPath, blob){ const id = 'f' + rid(); const name = relPath.split('/').pop(); this.files.set(id, {blob, name, path:relPath, url:URL.createObjectURL(blob)}); return {id, name}; },
  async move(id, folderRel){ const f = this.files.get(id); if (f) f.path = folderRel + '/' + f.name; },
  async thumbUrl(id){ const f = this.files.get(id); if (!f) throw new Error('missing'); return f.url; },
  async bytes(id){ return await this.files.get(id).blob.arrayBuffer(); },
  async webUrl(id){ return this.files.get(id).url; },
  async rootUrl(){ return null; },
  async remove(id){ this.files.delete(id); }, async removePath(){}
};
const store = DEMO ? DemoStore : GraphStore;
function sync(){ S.units = store.data.units || {}; S.ev = store.data.ev || {}; S.profile = store.data.profile || {}; }
async function mutate(fn, okMsg){
  S.saving = true; render();
  try { await store.mutate(fn); sync(); if (okMsg) toast(okMsg); }
  catch (e){ toast(storeMsg(e)); throw e; }
  finally { S.saving = false; render(); }
}
function storeMsg(e){
  if (!e) return 'Something went wrong. Try again.';
  if (e.status === 507) return 'Your OneDrive is full. Free up space, then try again.';
  if (e.status === 401 || e.status === 403) return 'OneDrive access was refused. Sign out and back in, then try again.';
  if (e.status) return `OneDrive didn’t accept that (${e.status}). Check your connection and try again.`;
  return e.message || 'Something went wrong. Try again.';
}

/* ================= AI (via Worker) ================= */
const accessCode = () => { try { return localStorage.getItem('logbook-access-code') || ''; } catch { return ''; } };
async function ai(kind, prompt, images=[], maxTokens=2000){
  if (DEMO) return demoAI(kind, prompt);
  if (!CFG.workerUrl || CFG.workerUrl.startsWith('PASTE')) throw Object.assign(new Error('The AI helper isn’t set up yet.'), {code:'no_worker'});
  const imgs = [];
  for (const b of images) imgs.push({media_type:'image/jpeg', data: await blobToBase64(b)});
  const r = await fetch(CFG.workerUrl, {method:'POST', headers:{'Content-Type':'application/json', 'X-Access-Code':accessCode()}, body:JSON.stringify({prompt, images:imgs, max_tokens:maxTokens}), signal: ai.signal});
  const data = await r.json().catch(()=>({}));
  if (!r.ok) throw Object.assign(new Error(data.error || 'AI error'), {code: data.error || 'ai_error', status: data.status || r.status});
  const txt = String(data.text || '').replace(/```json|```/g, '').trim();
  const a = txt.indexOf('{'), b = txt.lastIndexOf('}');
  try { return JSON.parse(txt.slice(a, b + 1)); } catch { throw Object.assign(new Error('bad json'), {code:'invalid_json'}); }
}
function aiMsg(e){
  if (e && e.name === 'AbortError') return 'Stopped.';
  switch (e && e.code){
    case 'bad_code': return 'The AI helper needs the access code from your trainer. Enter it in Settings.';
    case 'no_worker': return 'The AI helper isn’t set up yet, so pick items yourself for now.';
    case 'origin_not_allowed': return 'The AI helper doesn’t recognise this web address. Check ALLOWED_ORIGINS on the Worker.';
    case 'too_large': return 'That file is too big to read in one go. Upload the pages for one unit at a time.';
    case 'invalid_json': return 'The AI’s answer couldn’t be used. Try again.';
    case 'upstream': return (e.status === 429 || e.status === 529) ? 'The AI is busy right now. Try again in a minute.' : 'The AI helper had a problem (' + e.status + '). Try again.';
    default: return 'The AI helper couldn’t be reached. Check your connection and try again.';
  }
}
function demoAI(kind){
  return new Promise(res => setTimeout(() => {
    const open = openItems();
    if (kind === 'checklist') res({units:[{number:'DEMO 101', title:'Demo unit – practise with this', credits:'4', pages:null, items:[
      {kind:'photo', text:'Photo of work area set up safely before starting'}, {kind:'photo', text:'Photo of correct tool selected for the task'},
      {kind:'photo', text:'Photo of finished job showing quality of work'}, {kind:'book', text:'Complete written questions 1-5 in the unit book'}]}]});
    else if (kind === 'match') res({matches: open.slice(0,3).map(x=>({unitId:x.u.id, itemId:x.it.id, why:'Demo suggestion'}))});
    else res({guide:{title:'Demo job', overview:'Demo guide – the real one is written for your job.', safety:['Demo safety point'], tools:['Demo tool'], checks:['Demo check 1','Demo check 2'], learn:['Demo learning point'], standard:''}, items: open.slice(0,4).map(x=>({unitId:x.u.id, itemId:x.it.id, shots:'Demo: photograph the key step'}))});
  }, 700));
}

/* ================= data helpers ================= */
function sortedUnits(){ return Object.values(S.units).sort((a,b)=> String(a.number||a.title).localeCompare(String(b.number||b.title), undefined, {numeric:true})); }
function evFor(uid, iid){ return Object.values(S.ev).filter(e=>e.unitId===uid && e.itemId===iid).sort((a,b)=>String(a.takenAt).localeCompare(String(b.takenAt))); }
function itemOk(u, it, s){ return it.kind === 'book' ? !!(u.bookDone && u.bookDone[it.id]) : (s.per[it.id]||[]).length > 0; }
function unitStats(u){
  const items = u.items || []; const per = {}; let done = 0;
  for (const it of items){ per[it.id] = evFor(u.id, it.id); }
  const s = {per}; for (const it of items) if (itemOk(u, it, s)) done++;
  const total = items.length;
  const status = u.signed ? (u.creditsConfirmed ? 'awarded' : 'signed') : (total && done === total) ? 'ready' : done > 0 ? 'collecting' : 'new';
  return {done, total, status, per};
}
function openItems(){
  const out = [];
  for (const u of sortedUnits()){ if (u.signed) continue; for (const it of (u.items||[])) if (it.kind !== 'book') out.push({u, it, n: evFor(u.id, it.id).length}); }
  return out;
}
function creditSummary(){
  const c = {awarded:0, signed:0, ready:0, progress:0, total:0, unknown:0};
  for (const u of Object.values(S.units)){
    const cr = num(u.credits); if (!cr) c.unknown++; c.total += cr;
    const st = unitStats(u).status;
    if (st === 'awarded') c.awarded += cr; else if (st === 'signed') c.signed += cr; else if (st === 'ready') c.ready += cr; else if (st === 'collecting') c.progress += cr;
  }
  return c;
}

/* ================= start-up ================= */
async function start(){
  render();
  try {
    if (DEMO){ S.account = {name:'Demo mode'}; sync(); S.phase = 'ready'; render(); return; }
    if (!window.msal){ S.phase = 'error'; S.err = 'The Microsoft sign-in library didn’t load. Check your connection and reload.'; render(); return; }
    const acct = await authInit();
    if (!acct){ S.phase = 'signin'; render(); return; }
    S.account = acct; S.phase = 'loading'; render();
    await store.load(); sync(); S.phase = 'ready'; render();
    document.addEventListener('visibilitychange', async () => {
      if (document.visibilityState !== 'visible' || S.phase !== 'ready' || S.saving) return;
      try { await store.load(); sync(); render(); } catch {}
    });
  } catch (e){
    S.phase = 'error'; S.err = (e && e.errorMessage) || storeMsg(e); render();
  }
}

/* ================= rendering ================= */
function render(){
  const a = document.activeElement; const fid = a && a.id; let ss = null, se = null;
  try { ss = a.selectionStart; se = a.selectionEnd; } catch {}
  const key = S.view + '|' + (S.unitId||'') + '|' + (S.editing?1:0) + '|' + (S.add.review?1:0);
  const sy = window.scrollY;
  $app.innerHTML = vMain() + vNav() + vModal() + (S.toast ? `<div class="toast" role="status">${esc(S.toast)}</div>` : '');
  if (key === lastView) window.scrollTo(0, sy); else window.scrollTo(0, 0);
  lastView = key;
  if (fid){ const el = document.getElementById(fid); if (el && el.type !== 'file'){ el.focus({preventScroll:true}); try { if (ss != null) el.setSelectionRange(ss, se); } catch {} } }
  fillThumbs();
}
function fillThumbs(){
  document.querySelectorAll('img[data-thumb]:not([src])').forEach(img => {
    store.thumbUrl(img.dataset.thumb).then(u => { img.src = u; }).catch(()=>{ img.alt = 'Photo unavailable'; });
  });
}
const thumb = (id, alt='') => `<img data-thumb="${esc(id)}" alt="${esc(alt)}" loading="lazy">`;

function vTop(title, sub){
  return `<div class="acct"><span>${esc(S.account && (S.account.name || S.account.username) || '')}</span>${S.saving?'<span class="saving">Saving…</span>':''}</div>
  ${DEMO?'<div class="demo">Demo mode – nothing is saved</div>':''}
  <header class="top"><h1>${title}</h1>${sub?`<p class="sub">${sub}</p>`:''}</header>`;
}

function vMain(){
  if (S.phase === 'start' || S.phase === 'loading') return `<div class="wall"><h1>Evidence logbook</h1><p class="status busy">${S.phase==='loading'?'Opening your logbook in OneDrive…':'Starting…'}</p></div>`;
  if (S.phase === 'error') return `<div class="wall"><h1>Evidence logbook</h1><p class="status err">${esc(S.err)}</p><button class="btn" data-act="reload">Reload</button> <button class="btn ghost" data-act="signout">Sign out</button></div>`;
  if (S.phase === 'signin') return `<div class="wall"><h1>Evidence logbook</h1>
    <p class="sub">Plan your unit evidence, file photos against your unit checklists, and print a pack for your trainer to sign.</p>
    <p>Sign in with your Microsoft account. Your photos and packs are saved in your own OneDrive, in a folder called “${esc(ROOT)}”.</p>
    <button class="btn primary block big" data-act="signin">Sign in with Microsoft</button>
    <p class="meta" style="margin-top:14px">Works with Outlook, Hotmail, Live and work or school Microsoft accounts.</p></div>`;
  if (S.view === 'unit') return vUnit();
  if (S.view === 'capture') return vCapture();
  if (S.view === 'plan') return vPlan();
  if (S.view === 'settings') return vSettings();
  return vUnits();
}

function punch(u, s){ return `<span class="punch" aria-hidden="true">${(u.items||[]).map(it => `<i class="${it.kind==='book'?'bk ':''}${itemOk(u,it,s)?'ok':''}"></i>`).join('')}</span>`; }

function vCredits(){
  const c = creditSummary(); const goal = num(S.profile.goal); const den = Math.max(goal || c.total, 1);
  const pct = v => Math.min(100, v/den*100).toFixed(2) + '%'; const f = v => Number.isInteger(v) ? v : v.toFixed(1);
  return `<div class="panel credits"><div class="big">${f(c.awarded)}<span>of ${f(goal || c.total)} credits awarded</span></div>
    <div class="cbar" role="img" aria-label="Credit progress"><i class="c-aw" style="width:${pct(c.awarded)}"></i><i class="c-sg" style="width:${pct(c.signed)}"></i><i class="c-rd" style="width:${pct(c.ready)}"></i><i class="c-ip" style="width:${pct(c.progress)}"></i></div>
    <div class="legend"><span><i class="c-aw"></i>Awarded <b>${f(c.awarded)}</b></span><span><i class="c-sg"></i>Signed off, waiting <b>${f(c.signed)}</b></span>
    <span><i class="c-rd"></i>Ready for assessor <b>${f(c.ready)}</b></span><span><i class="c-ip"></i>In progress <b>${f(c.progress)}</b></span></div>
    ${goal ? `<p class="meta">${f(c.total)} credits loaded so far${goal > c.total ? ` · ${f(goal - c.total)} still to load` : ''}</p>` : `<p class="meta">Set your qualification total in Settings to track against it.</p>`}
    ${c.unknown ? `<p class="meta">${c.unknown} unit${c.unknown===1?' has':'s have'} no credit value. Add it with Edit checklist.</p>` : ''}</div>`;
}

function vUnits(){
  let h = vTop('Evidence logbook', esc(S.profile.programme || 'Your apprenticeship units'));
  if (S.add.review) return h + vReview();
  const us = sortedUnits();
  if (us.length) h += vCredits();
  h += `<label class="btn primary block" for="f-book" role="button" tabindex="0">Add unit book (PDF or page photos)</label>`;
  if (S.add.msg) h += `<p class="status ${S.add.busy?'busy':''}">${esc(S.add.msg)}</p>`;
  if (S.add.busy) h += `<button class="btn ghost" data-act="add-stop">Stop</button>`;
  if (!us.length){
    h += `<div class="empty"><h2>No units yet</h2><p>Add your unit book and it becomes a checklist of the photos and book work you need for each unit. You check the list before it’s saved.</p></div>`;
  } else {
    h += `<ul class="units">${us.map(u => { const s = unitStats(u); return `<li><button class="unitcard" data-act="open-unit" data-id="${u.id}">
      ${u.number?`<span class="un">${esc(u.number)}${u.credits?` · ${esc(u.credits)} credits`:''}</span>`:''}
      <span class="ut">${esc(u.title||'Untitled unit')}</span>${punch(u,s)}
      <span class="chip ${s.status}">${STATUS[s.status]}</span><span class="cnt">${s.done} of ${s.total}</span></button></li>`; }).join('')}</ul>`;
  }
  return h;
}

function vEditorUnit(u, base, opts={}){
  return `<div class="eunit">
    <div class="grid3">
      <label class="field"><span>Unit no.</span><input type="text" id="${base}-num" data-bind="${base}.number" value="${esc(u.number)}"></label>
      <label class="field wide"><span>Title</span><input type="text" id="${base}-title" data-bind="${base}.title" value="${esc(u.title)}"></label>
      <label class="field"><span>Credits</span><input type="text" id="${base}-cr" data-bind="${base}.credits" value="${esc(u.credits)}"></label>
    </div>
    <h3>Evidence checklist</h3>
    ${(u.items||[]).map((it,i)=>`<div class="eitem">
      <select id="${base}-k${i}" data-bind="${base}.items.${i}.kind" aria-label="Item type">
        <option value="photo" ${it.kind!=='book'?'selected':''}>Photo evidence</option>
        <option value="book" ${it.kind==='book'?'selected':''}>Book work (done in unit book)</option>
      </select>
      <button class="btn danger small" data-act="item-del" data-base="${base}" data-i="${i}">Remove</button>
      <textarea id="${base}-t${i}" data-bind="${base}.items.${i}.text" aria-label="Checklist item">${esc(it.text)}</textarea>
    </div>`).join('')}
    <div class="row" style="margin-top:10px">
      <button class="btn small" data-act="item-add" data-base="${base}">Add item</button>
      ${opts.removable?`<button class="btn danger small" data-act="rv-unit-del" data-i="${opts.index}">Remove this unit</button>`:''}
    </div>
  </div>`;
}

function vReview(){
  const r = S.add.review;
  return `<section class="editor">
    <h2>Check the checklist</h2>
    <p>Here’s what the unit book asks you to show. Fix anything that’s wrong, then save. Your trainer’s and assessor’s requirements always come first.</p>
    ${r.note?`<p class="status">${esc(r.note)}</p>`:''}
    ${r.units.map((u,i)=>vEditorUnit(u,'add.review.units.'+i,{removable:true,index:i})).join('')}
    <div class="row"><button class="btn go" data-act="rv-save" ${S.saving?'disabled':''}>Save ${r.units.length===1?'unit':r.units.length+' units'}</button><button class="btn ghost" data-act="rv-cancel">Cancel</button></div>
  </section>`;
}

function vUnit(){
  const u = S.units[S.unitId];
  if (!u){ S.view = 'units'; return vUnits(); }
  if (S.editing && S.editing.unitId === u.id){
    return `<button class="back" data-act="edit-cancel">‹ Cancel</button><section class="editor"><h2>Edit checklist</h2>${vEditorUnit(S.editing.unit,'editing.unit')}
      <div class="row"><button class="btn go" data-act="edit-save" ${S.saving?'disabled':''}>Save changes</button><button class="btn ghost" data-act="edit-cancel">Cancel</button></div></section>`;
  }
  const s = unitStats(u); const missing = s.total - s.done;
  let h = `<button class="back" data-act="nav" data-v="units">‹ All units</button>
  <header class="unithead">${u.number?`<div class="un">${esc(u.number)}${u.credits?` · ${esc(u.credits)} credits`:''}</div>`:''}
    <h1>${esc(u.title||'Untitled unit')}</h1>${punch(u,s)}
    <span class="chip ${s.status}">${STATUS[s.status]}</span> <span class="meta">${s.done} of ${s.total} items done</span></header>`;
  h += `<div class="panel">`;
  if (s.status === 'signed' || s.status === 'awarded'){
    h += `<h3>Signed off ${u.signed.at?fmtDate(u.signed.at):''}</h3><p><button class="linkbtn" data-act="open-file" data-id="${esc(u.signed.id)}">Open the signed copy</button></p>
      <label class="bookcheck" style="margin-left:0"><input type="checkbox" data-act="confirm-credits" ${u.creditsConfirmed?'checked':''}> Credits confirmed by Competenz / MITO${u.credits?` (${esc(u.credits)})`:''}</label>
      <div class="row"><button class="btn small" data-act="pack-open" ${S.packBusy?'disabled':''}>Print evidence pack again</button><button class="btn danger small" data-act="unsign">Remove signed copy</button></div>`;
  } else if (s.status === 'ready'){
    h += `<h3>Ready for your trainer</h3><p>Every item has evidence. Print the pack, hand it to your trainer or assessor to sign, then upload the signed copy.</p>
      <div class="stack"><button class="btn primary block" data-act="pack-open" ${S.packBusy?'disabled':''}>${S.packBusy?'Building pack…':'Print evidence pack'}</button>
      <label class="btn block" for="f-signed" role="button" tabindex="0">Upload signed copy</label></div>`;
  } else {
    h += `<h3>${missing} item${missing===1?'':'s'} still to do</h3><p class="meta">The final pack unlocks once every item has a photo or is ticked off in your unit book.</p>
      <button class="btn block" data-act="pack-open" data-draft="1" ${S.packBusy?'disabled':''}>${S.packBusy?'Building…':'Preview draft pack'}</button>`;
  }
  h += `</div>`;
  h += `<div class="panel"><h3>Checklist</h3><ul class="items">${(u.items||[]).map((it,i)=>{
    const ok = itemOk(u,it,s); let body = '';
    if (it.kind === 'book'){
      body = `<label class="bookcheck"><input type="checkbox" data-act="book-toggle" data-item="${it.id}" ${ok?'checked':''} ${u.signed?'disabled':''}> Done in my unit book</label>`;
    } else {
      const evs = s.per[it.id] || [];
      body = `<div class="thumbs">${evs.filter(e=>e.photos&&e.photos.length).map(e => `<button data-act="ev-open" data-id="${e.id}" aria-label="Open evidence from ${esc(fmtDate(e.takenAt))}">${thumb(e.photos[0].id)}</button>`).join('')}
        ${u.signed?'':`<button class="addph" data-act="cap-for" data-unit="${u.id}" data-item="${it.id}">Add photo</button>`}</div>`;
    }
    return `<li class="item ${ok?'ok':''}"><div class="ihead"><span class="tick">${ok?'✓':i+1}</span><p><span class="kind">${it.kind==='book'?'Book work':'Photo evidence'}</span>${esc(it.text)}</p></div>${body}</li>`;
  }).join('')}</ul></div>`;
  h += `<div class="panel"><h3>Unit book</h3>${u.book?`<p><button class="linkbtn" data-act="open-file" data-id="${esc(u.book.id)}">Open unit book${u.pages?` (pages ${u.pages[0]}–${u.pages[1]})`:''}</button></p>`:'<p class="meta">No unit book attached.</p>'}
    <p class="meta">Filled in your book on paper or in the PDF? Upload the completed pages for this unit so they go in the evidence pack.</p>
    <label class="btn small" for="f-replbook" role="button" tabindex="0">Upload completed book pages</label></div>`;
  h += `<p class="meta">OneDrive folder: ${esc(ROOT)} › ${esc(u.folder)}</p>`;
  h += `<div class="row" style="margin-top:16px">${u.signed?'':`<button class="btn small" data-act="edit-open">Edit checklist</button>`}<button class="btn danger small" data-act="unit-del">Delete unit</button></div>`;
  return h;
}

function vCapture(){
  const c = S.cap; const items = openItems(); const units = [...new Set(items.map(x=>x.u))];
  let h = vTop('Capture evidence', 'Photo, your note, then pick where it goes.');
  if (!units.length) return h + `<div class="empty"><h2>Add a unit book first</h2><p>Photos are filed against your unit checklists.</p><button class="btn primary" data-act="nav" data-v="units">Go to units</button></div>`;
  h += `<label class="btn primary block big" for="f-cam" role="button" tabindex="0">Take photo</label>
    <label class="btn ghost block" style="margin-top:10px" for="f-gal" role="button" tabindex="0">Choose from gallery</label>
    ${c.reading?'<p class="status busy">Preparing photo…</p>':''}`;
  if (c.photos.length) h += `<div class="capthumbs">${c.photos.map((p,i)=>`<div><img src="${p.url}" alt="Photo ${i+1}"><button data-act="ph-del" data-i="${i}" aria-label="Remove photo ${i+1}">×</button></div>`).join('')}</div>`;
  h += `<label class="field"><span>Your note</span><small>What did you do? Your own words — tap the mic on your keyboard to talk instead of typing.</small><textarea id="cap-note" data-bind="cap.note">${esc(c.note)}</textarea></label>
    <label class="field"><span>Job or vehicle</span><input type="text" id="cap-job" data-bind="cap.job" value="${esc(c.job)}" placeholder="e.g. Hino 500 – 250-hour service"></label>
    <h2 style="margin-top:22px">Where does it go?</h2>
    <button class="btn block" data-act="suggest" style="margin-top:10px" ${c.busy?'disabled':''}>Suggest the checklist item</button>`;
  if (c.msg) h += `<p class="status ${c.busy?'busy':''}">${esc(c.msg)}</p>`;
  if (c.sugg && c.sugg.length){
    h += `<ul class="sugg">${c.sugg.map(m=>{ const u = S.units[m.unitId]; const it = (u.items||[]).find(x=>x.id===m.itemId); const sel = c.unitId===m.unitId && c.itemId===m.itemId;
      return `<li><button class="${sel?'sel':''}" data-act="pick" data-unit="${m.unitId}" data-item="${m.itemId}" aria-pressed="${sel}"><span class="su">${esc(unitLabel(u))}</span><span class="si">${esc(it.text)}</span>${m.why?`<span class="sw">${esc(m.why)}</span>`:''}</button></li>`; }).join('')}</ul>`;
  }
  const selUnit = S.units[c.unitId];
  h += `<label class="field"><span>Unit</span><select id="cap-unit"><option value="">Choose a unit…</option>${units.map(u=>`<option value="${u.id}" ${c.unitId===u.id?'selected':''}>${esc(unitLabel(u))}</option>`).join('')}</select></label>`;
  if (selUnit) h += `<label class="field"><span>Checklist item</span><select id="cap-item"><option value="">Choose an item…</option>${(selUnit.items||[]).filter(it=>it.kind!=='book').map(it=>{ const n = evFor(selUnit.id,it.id).length; return `<option value="${it.id}" ${c.itemId===it.id?'selected':''}>${esc(it.text)}${n?` (${n} saved)`:''}</option>`; }).join('')}</select></label>`;
  const can = c.photos.length && c.unitId && c.itemId && !c.busy;
  h += `<button class="btn go block big" data-act="save-ev" ${can?'':'disabled'} style="margin-top:8px">Save evidence</button>`;
  return h;
}

function vPlan(){
  const p = S.plan;
  let h = vTop('Plan a job', 'Tell it what’s coming up at work. You’ll get a job guide and see which checklist items the job could tick off.');
  h += `<label class="field"><span>Upcoming job</span><input type="text" id="plan-q" data-bind="plan.q" value="${esc(p.q)}" placeholder="e.g. Oil and filter service on a Hino 500"></label>
    <button class="btn primary block" data-act="plan" ${p.busy?'disabled':''}>What can this job cover?</button>`;
  if (p.msg) h += `<p class="status ${p.busy?'busy':''}">${esc(p.msg)}</p>`;
  if (p.guide){
    const g = p.guide; const list = (title, arr) => arr && arr.length ? `<h3 style="margin-top:14px">${title}</h3><ul class="guide">${arr.map(x=>`<li>${esc(x)}</li>`).join('')}</ul>` : '';
    h += `<div class="panel"><h3>Job guide${g.title?': '+esc(g.title):''}</h3>
      ${g.overview?`<p>${esc(g.overview)}</p>`:''}
      ${list('Safety first', g.safety)}${list('Tools and gear', g.tools)}${list('What to check', g.checks)}${list('Learning points', g.learn)}
      <p class="rule">General guidance only. Always follow your workshop’s procedures, the manufacturer’s manual and specs${g.standard?', and '+esc(g.standard):''}, and check with your supervisor.</p></div>`;
  }
  if (p.res && p.res.length){
    h += `<div class="panel"><h3>This job could cover</h3><ul class="items">${p.res.map(r=>{ const u = S.units[r.unitId]; if (!u) return ''; const it = (u.items||[]).find(x=>x.id===r.itemId); if (!it) return ''; const n = evFor(u.id,it.id).length;
      return `<li class="item ${n?'ok':''}"><div class="ihead"><span class="tick">${n?'✓':'•'}</span><p><span class="kind">${esc(unitLabel(u))}</span>${esc(it.text)}${r.shots?`<span class="meta" style="display:block">Photograph: ${esc(r.shots)}</span>`:''}</p></div>
      <div class="thumbs"><button class="btn small" data-act="cap-for" data-unit="${u.id}" data-item="${it.id}" data-job="1">Capture for this</button></div></li>`; }).join('')}</ul></div>`;
  }
  return h;
}

function vSettings(){
  if (!S.pf) S.pf = {name:S.profile.name||'', trainer:S.profile.trainer||'', workplace:S.profile.workplace||'', programme:S.profile.programme||'', goal:S.profile.goal||'', code:accessCode()};
  const f = S.pf;
  return vTop('Settings', 'These details go on the cover of each evidence pack.') + `
  <label class="field"><span>Apprentice name</span><input type="text" id="pf-name" data-bind="pf.name" value="${esc(f.name)}"></label>
  <label class="field"><span>Workplace trainer</span><input type="text" id="pf-trainer" data-bind="pf.trainer" value="${esc(f.trainer)}"></label>
  <label class="field"><span>Workplace</span><input type="text" id="pf-wp" data-bind="pf.workplace" value="${esc(f.workplace)}"></label>
  <label class="field"><span>Programme</span><input type="text" id="pf-prog" data-bind="pf.programme" value="${esc(f.programme)}" placeholder="e.g. Competenz Light Fabrication L3"></label>
  <label class="field"><span>Credits needed for the qualification</span><small>From your training plan, e.g. 120 for a Level 3 certificate.</small><input type="text" inputmode="numeric" id="pf-goal" data-bind="pf.goal" value="${esc(f.goal)}"></label>
  <label class="field"><span>AI access code</span><small>From your trainer. Only stored on this device.</small><input type="text" id="pf-code" data-bind="pf.code" value="${esc(f.code)}" autocomplete="off"></label>
  <button class="btn go" data-act="pf-save" ${S.saving?'disabled':''}>Save details</button>
  <div class="panel"><h3>Your OneDrive folder</h3><p>Everything is saved in your OneDrive under “${esc(ROOT)}”, one folder per unit.</p>
    ${DEMO?'':'<button class="btn small" data-act="open-root">Open folder in OneDrive</button>'}</div>
  <div class="panel"><h3>How this app works</h3>
    <p>It plans your evidence, files your photos against the right checklist item, and tells you what’s missing.</p>
    <p>It never writes notes, answers or evidence for you. Every note is yours, and only your trainer or assessor can sign a unit off.</p></div>
  <div class="row">${DEMO?'':'<button class="btn ghost small" data-act="signout">Sign out</button>'}</div>
  <p class="meta">Version ${APP_VERSION}</p>`;
}

function vNav(){
  if (S.phase !== 'ready') return '';
  const b = (v,label,cls='') => `<button class="${cls} ${S.view===v||(v==='units'&&S.view==='unit')?'on':''}" data-act="nav" data-v="${v}">${label}</button>`;
  return `<nav class="bar"><div class="in">${b('units','Units')}${b('plan','Plan')}${b('capture','Capture','cap')}${b('settings','Settings')}</div></nav>`;
}

function vModal(){
  const m = S.modal; if (!m) return '';
  if (m.type === 'ev'){
    const e = S.ev[m.id]; if (!e) return ''; const u = S.units[e.unitId];
    const opts = sortedUnits().filter(x=>!x.signed).flatMap(x=>(x.items||[]).filter(it=>it.kind!=='book').map(it=>`<option value="${x.id}|${it.id}" ${x.id===e.unitId&&it.id===e.itemId?'selected':''}>${esc((x.number||x.title)+': '+it.text)}</option>`)).join('');
    return `<div class="modal" data-act="modal-bg"><div class="sheet" role="dialog" aria-label="Evidence">
      <div class="row" style="justify-content:space-between"><h2>Evidence</h2><button class="btn small" data-act="modal-close">Close</button></div>
      <p class="meta">${esc(fmtDate(e.takenAt))}${e.job?' · '+esc(e.job):''}</p>
      ${(e.photos||[]).map(p=>thumb(p.id,'Evidence photo')).join('')}
      <h3>Your note</h3><div class="note">${e.note?esc(e.note):'<span class="meta">No note</span>'}</div>
      ${u && !u.signed ? `<label class="field"><span>Move to another item</span><select id="mv-target">${opts}</select></label>
      <button class="btn danger" data-act="ev-del">Delete this evidence</button>` : ''}
    </div></div>`;
  }
  if (m.type === 'pack'){
    const u = S.units[S.unitId]; if (!u) return '';
    return `<div class="modal" data-act="modal-bg"><div class="sheet" role="dialog" aria-label="Evidence pack">
      <h2>${m.draft?'Preview draft pack':'Print evidence pack'}</h2>
      ${m.draft?'<p class="status">This draft is marked “DRAFT – not complete” on every page and lists what’s still missing. It isn’t for sign-off.</p>':''}
      <p>One PDF with a cover page${u.book?', your unit book pages':''}, every photo with your notes under its checklist item, and a sign-off page for your trainer. It’s saved in the unit’s OneDrive folder and downloaded to this device.</p>
      ${u.book?`<label class="bookcheck" style="margin-left:0"><input type="checkbox" id="pk-book" ${m.book!==false?'checked':''}> Include unit book pages</label>`:''}
      <div class="row" style="margin-top:12px"><button class="btn primary" data-act="pack-build" ${S.packBusy?'disabled':''}>${S.packBusy?'Building…':'Build PDF'}</button><button class="btn ghost" data-act="modal-close">Cancel</button></div>
    </div></div>`;
  }
  return '';
}

/* ================= images & PDFs ================= */
function loadImg(file){ return new Promise((res,rej)=>{ const i = new Image(); i.onload = () => res(i); i.onerror = () => rej(new Error('That photo format could not be opened. Try a JPEG or PNG.')); i.src = URL.createObjectURL(file); }); }
async function compressImage(file, max=1600, q=0.82){
  let src; try { src = await createImageBitmap(file, {imageOrientation:'from-image'}); } catch { src = await loadImg(file); }
  const w0 = src.width || src.naturalWidth, h0 = src.height || src.naturalHeight;
  const sc = Math.min(1, max/Math.max(w0,h0)); const w = Math.round(w0*sc), h = Math.round(h0*sc);
  const cv = document.createElement('canvas'); cv.width = w; cv.height = h;
  const ctx = cv.getContext('2d'); ctx.fillStyle = '#fff'; ctx.fillRect(0,0,w,h); ctx.drawImage(src,0,0,w,h);
  return await new Promise((res,rej)=>cv.toBlob(b=>b?res(b):rej(new Error('Could not read that photo.')),'image/jpeg',q));
}
async function renderPage(pdf, p){
  const page = await pdf.getPage(p); const v1 = page.getViewport({scale:1});
  const vp = page.getViewport({scale: Math.min(2.2, 1500/v1.width)});
  const cv = document.createElement('canvas'); cv.width = vp.width; cv.height = vp.height;
  const ctx = cv.getContext('2d'); ctx.fillStyle = '#fff'; ctx.fillRect(0,0,cv.width,cv.height);
  await page.render({canvasContext:ctx, viewport:vp}).promise;
  return await new Promise(res=>cv.toBlob(res,'image/jpeg',0.8));
}
async function imagesToPdf(blobs){
  const {PDFDocument} = PDFLib; const doc = await PDFDocument.create();
  for (const b of blobs){
    const img = await doc.embedJpg(new Uint8Array(await b.arrayBuffer()));
    const W = 595.28, H = 841.89, M = 18; const sc = Math.min((W-2*M)/img.width, (H-2*M)/img.height);
    const pg = doc.addPage([W,H]); pg.drawImage(img,{x:(W-img.width*sc)/2, y:(H-img.height*sc)/2, width:img.width*sc, height:img.height*sc});
  }
  return await doc.save();
}
async function filesToPdfBlob(files){
  const pdfFile = files.find(isPdf);
  if (pdfFile) return pdfFile;
  const blobs = []; for (const f of files) blobs.push(await compressImage(f, 1800, 0.85));
  return new Blob([await imagesToPdf(blobs)], {type:'application/pdf'});
}

/* ================= AI prompts ================= */
function checklistPrompt(text, numPages, withImages){
  return `You are helping a trade apprentice in New Zealand organise the evidence they must collect for their unit book. ${withImages ? 'The unit book pages are attached as images, in order (the first image is page 1).' : 'The text of their unit book is below, with [Page n] markers.'}

Identify each unit standard (or assessment unit) in the book. For each one, list the evidence the apprentice must collect, based ONLY on what the book's performance criteria and assessment instructions ask for.

Rules:
- Each checklist item describes evidence to collect, starting with what to capture, e.g. "Photo of new fuel filter fitted and system bled" or "Photo of torque wrench set to spec on wheel nuts".
- Use kind "photo" for anything shown with photos or other workplace evidence. Use kind "book" for written questions, worksheets or sections the apprentice completes in the unit book themselves; describe only what and where (e.g. "Complete written questions 1-12, pages 8-10"), never the answers.
- Never answer questions. Never write descriptions of work, notes or evidence statements for the apprentice.
- Keep each item under 20 words. Merge duplicates. Usually 4-15 items per unit.
- "pages": the first and last page numbers of that unit in the book if you can tell, otherwise null. The book has ${numPages || 'an unknown number of'} pages.
- If the file is not a unit book, return {"units":[]}.

Reply with only JSON in this shape:
{"units":[{"number":"US 12345","title":"Unit title","level":"3","credits":"4","pages":[1,12],"items":[{"kind":"photo","text":"..."}]}]}
${withImages ? '' : '\nUNIT BOOK:\n' + text.slice(0, 180000)}`;
}
function normUnit(u, numPages){
  let pages = null;
  if (Array.isArray(u.pages) && u.pages.length === 2){ const a = parseInt(u.pages[0]), b = parseInt(u.pages[1]); if (a >= 1 && b >= a && (!numPages || b <= numPages)) pages = [a,b]; }
  return { number:String(u.number||'').trim(), title:String(u.title||'').trim(), credits:String(u.credits||'').trim(), level:String(u.level||'').trim(), pages,
    items:(Array.isArray(u.items)?u.items:[]).map(it=>({id:rid(), kind: it && it.kind==='book' ? 'book' : 'photo', text:String(it && it.text || '').trim()})).filter(it=>it.text) };
}

/* ================= actions ================= */
const MAX_AI_IMAGES = 20;
async function addBook(files){
  if (!files.length || S.add.busy) return;
  S.add.busy = true; S.add.msg = 'Opening the unit book…'; render();
  addCtl = new AbortController(); ai.signal = addCtl.signal;
  try {
    let text = '', images = [], numPages = 0, note = '';
    const pdfFile = files.find(isPdf);
    if (pdfFile){
      if (!window.pdfjsLib) throw new Error('The PDF reader didn’t load. Reload the page and try again.');
      const bytes = new Uint8Array(await pdfFile.arrayBuffer());
      const pdf = await pdfjsLib.getDocument({data: bytes.slice()}).promise; numPages = pdf.numPages;
      for (let p=1; p<=numPages; p++){ S.add.msg = `Reading page ${p} of ${numPages}…`; render(); const tc = await (await pdf.getPage(p)).getTextContent(); text += `\n[Page ${p}]\n` + tc.items.map(i=>i.str).join(' '); }
      if (text.replace(/\[Page \d+\]|\s/g,'').length < 150 * Math.min(numPages, 3)){
        const n = Math.min(numPages, MAX_AI_IMAGES);
        for (let p=1; p<=n; p++){ S.add.msg = `Preparing scanned page ${p} of ${n}…`; render(); images.push(await renderPage(pdf, p)); }
        if (numPages > n) note = `Only the first ${n} pages of this scanned book could be read. Upload the rest as a separate file.`;
        text = '';
      }
    } else {
      for (const f of files.slice(0, MAX_AI_IMAGES)) images.push(await compressImage(f, 1800, 0.85));
      numPages = images.length;
      if (files.length > MAX_AI_IMAGES) note = `Only the first ${MAX_AI_IMAGES} page photos could be read. Add the rest as a separate upload.`;
    }
    S.add.msg = 'Building your evidence checklist from the unit book. This can take a minute or two…'; render();
    const res = await ai('checklist', checklistPrompt(text, numPages, images.length > 0), images, 8000);
    const units = (res && Array.isArray(res.units) ? res.units : []).map(u=>normUnit(u, numPages)).filter(u=>u.items.length);
    if (!units.length) throw new Error('No units or evidence requirements were found in that file. Check it’s the unit book and try again.');
    S.add.msg = 'Saving the unit book to OneDrive…'; render();
    const bookBlob = pdfFile || new Blob([await imagesToPdf(images)], {type:'application/pdf'});
    const bookName = safeName(pdfFile ? pdfFile.name.replace(/\.pdf$/i,'') : 'Unit book ' + ymd(new Date().toISOString())) + '.pdf';
    const book = await store.upload('Unit books/' + bookName, bookBlob, 'rename');
    S.add.review = {book, units, note}; S.add.msg = '';
  } catch (err){
    S.add.msg = (err && err.code) ? aiMsg(err) : (err && err.name === 'AbortError') ? 'Stopped.' : (err && err.status) ? storeMsg(err) : ((err && err.message) || 'Something went wrong reading that file.');
  } finally { S.add.busy = false; addCtl = null; ai.signal = undefined; render(); }
}

async function saveReview(){
  const r = S.add.review; const now = new Date().toISOString();
  const used = new Set(Object.values(S.units).map(u=>u.folder));
  const toAdd = r.units.map(u => {
    const items = u.items.filter(i=>String(i.text).trim()).map(i=>({id:i.id, kind:i.kind==='book'?'book':'photo', text:String(i.text).trim()}));
    if (!items.length) return null;
    let folder = safeName([u.number, u.title].filter(Boolean).join(' - ') || 'Unit'), k = 2; while (used.has(folder)) folder = safeName(folder) + ' (' + (k++) + ')'; used.add(folder);
    return {id:'u_'+rid(), number:String(u.number).trim(), title:String(u.title).trim()||'Untitled unit', credits:String(u.credits).trim(), level:u.level||'', pages:u.pages||null, book:r.book, folder, items, bookDone:{}, signed:null, creditsConfirmed:false, createdAt:now};
  }).filter(Boolean);
  try { await mutate(d => { for (const u of toAdd) d.units[u.id] = u; }, toAdd.length === 1 ? 'Unit saved' : toAdd.length + ' units saved'); S.add.review = null; render(); } catch {}
}

async function addPhotos(files){
  for (const f of files){
    let blob = null;
    try { blob = await compressImage(f); } catch (err){ if (/^image\/(jpeg|jpg)$/i.test(f.type)) blob = f; else toast(err.message || 'That photo could not be opened.'); }
    if (blob) S.cap.photos.push({blob, url:URL.createObjectURL(blob), takenAt:new Date(f.lastModified || Date.now()).toISOString()});
  }
  S.cap.reading = false; S.cap.sugg = null; S.cap.msg = ''; render();
}

async function suggest(){
  const c = S.cap;
  if (!c.photos.length && !c.note.trim()){ toast('Take a photo or write your note first.'); return; }
  const items = openItems(); if (!items.length) return;
  const sendImgs = c.photos.slice(0, 4).map(p=>p.blob);
  c.busy = true; c.sugg = null; c.msg = 'Matching to your checklist…'; render();
  const list = items.map(x=>({unitId:x.u.id, itemId:x.it.id, unit:unitLabel(x.u), item:x.it.text, have:x.n}));
  const prompt = `An apprentice has collected workplace evidence: ${c.photos.length} photo(s)${sendImgs.length?' (attached)':''}.
Their note: ${JSON.stringify(c.note.trim() || '(none)')}
Job or vehicle: ${JSON.stringify(c.job.trim() || '(none)')}

Their evidence checklist as JSON (have = photos already saved for that item):
${JSON.stringify(list)}

Pick up to 3 checklist items this evidence most likely belongs to, best first. Use only unitId and itemId values from the list. Prefer items with fewer saved photos when two fit equally. If nothing fits, return an empty list.
"why": at most 12 words saying what in the photo or note matches the item. Do not describe, judge or assess the work.
Reply with only JSON: {"matches":[{"unitId":"...","itemId":"...","why":"..."}]}`;
  try {
    const r = await ai('match', prompt, sendImgs, 600);
    const ms = (r && Array.isArray(r.matches) ? r.matches : []).filter(m => S.units[m.unitId] && (S.units[m.unitId].items||[]).some(it=>it.id===m.itemId && it.kind!=='book')).slice(0,3).map(m=>({unitId:m.unitId, itemId:m.itemId, why:String(m.why||'').slice(0,120)}));
    c.sugg = ms;
    if (ms.length){ c.unitId = ms[0].unitId; c.itemId = ms[0].itemId; c.msg = 'Best match is selected. Tap another if it fits better.'; }
    else c.msg = 'No clear match. Pick the unit and item yourself below.';
  } catch (e){ c.msg = aiMsg(e); }
  finally { c.busy = false; render(); }
}

async function saveEvidence(){
  const c = S.cap; const u = S.units[c.unitId]; if (!u) return;
  const idx = (u.items||[]).findIndex(it=>it.id===c.itemId); const it = u.items[idx];
  c.busy = true; render();
  try {
    const photos = [];
    for (let i=0; i<c.photos.length; i++){
      c.msg = `Uploading photo ${i+1} of ${c.photos.length}…`; render();
      const name = safeName(`${ymd(c.photos[i].takenAt)} Item ${idx+1} - ${it.text.slice(0,40)}${c.photos.length>1?' ('+(i+1)+')':''}`) + '.jpg';
      photos.push(await store.upload(`${u.folder}/Photos/${name}`, c.photos[i].blob, 'rename'));
    }
    const ev = {id:'e_'+rid(), unitId:c.unitId, itemId:c.itemId, photos, note:c.note.trim(), job:c.job.trim(), takenAt:c.photos[0].takenAt, createdAt:new Date().toISOString()};
    await mutate(d => { d.ev[ev.id] = ev; });
    c.photos.forEach(p=>URL.revokeObjectURL(p.url));
    const job = c.job; S.cap = freshCap(); S.cap.job = job;
    toast('Saved to ' + (u.number || u.title));
  } catch (e){ c.busy = false; c.msg = storeMsg(e); render(); }
}

async function runPlan(){
  const p = S.plan; const q = p.q.trim(); if (!q){ toast('Describe the job first.'); return; }
  const items = openItems();
  p.busy = true; p.res = null; p.guide = null; p.msg = 'Putting together a job guide…'; render();
  const list = items.map(x=>({unitId:x.u.id, itemId:x.it.id, unit:unitLabel(x.u), item:x.it.text, have:x.n}));
  const prompt = `An apprentice in New Zealand has this job coming up at work: ${JSON.stringify(q)}.
Their programme: ${JSON.stringify(S.profile.programme || 'a trade apprenticeship')}.

Part 1 - "guide": a short practical briefing to help them prepare, like an experienced tradesperson would give before the job.
- "title": short job name. "overview": 1-2 sentences on what the job involves.
- "safety": up to 5 key hazards or precautions. "tools": up to 6 tools, gear or documents needed.
- "checks": up to 12 things to inspect, check or do, in a sensible order, each under 18 words.
- "learn": up to 3 things worth understanding or asking their supervisor about.
- "standard": the main NZ standard, manual or rule that applies if there is a well-known one (e.g. the NZTA Vehicle Inspection Requirements Manual for a CoF), otherwise "".
- Do not give specific numbers such as torque settings, pressures, clearances or limits; say to check the manufacturer's spec or the relevant manual instead.
- Do not write notes, answers to unit book questions, or evidence statements for the apprentice.

Part 2 - "items": from their evidence checklist below (have = photos already saved), list the items this job could realistically provide evidence for, most useful first, preferring items where have is 0. Use only unitId and itemId values from the list. At most 10. If the list is empty, return [].
"shots": what to photograph during the job, at most 15 words.
Checklist: ${JSON.stringify(list)}

Reply with only JSON: {"guide":{"title":"","overview":"","safety":[],"tools":[],"checks":[],"learn":[],"standard":""},"items":[{"unitId":"...","itemId":"...","shots":"..."}]}`;
  try {
    const r = await ai('plan', prompt, [], 2500);
    const g = r && r.guide; const arr = (a, n) => (Array.isArray(a) ? a : []).map(x=>String(x||'').trim()).filter(Boolean).slice(0, n).map(x=>x.slice(0, 220));
    p.guide = g ? {title:String(g.title||'').slice(0,80), overview:String(g.overview||'').slice(0,400), safety:arr(g.safety,5), tools:arr(g.tools,6), checks:arr(g.checks,12), learn:arr(g.learn,3), standard:String(g.standard||'').slice(0,120)} : null;
    p.res = (r && Array.isArray(r.items) ? r.items : []).filter(x => S.units[x.unitId] && (S.units[x.unitId].items||[]).some(it=>it.id===x.itemId)).slice(0,10).map(x=>({unitId:x.unitId, itemId:x.itemId, shots:String(x.shots||'').slice(0,160)}));
    p.msg = p.res.length ? '' : (items.length ? 'This job doesn’t match anything still open on your checklists.' : 'Add a unit book to see which checklist items a job can cover.');
  } catch (e){ p.msg = aiMsg(e); }
  finally { p.busy = false; render(); }
}

async function uploadSigned(file){
  const u = S.units[S.unitId]; if (!file || !u) return;
  if (!confirm('Upload the copy your trainer or assessor has signed? This marks the unit as signed off.')) return;
  try {
    toast('Uploading signed copy…');
    const pdf = isPdf(file);
    const r = await store.upload(`${u.folder}/Signed copy.${pdf?'pdf':'jpg'}`, pdf ? file : await compressImage(file, 2200, 0.85), 'replace');
    await mutate(d => { d.units[u.id].signed = {id:r.id, name:r.name, at:new Date().toISOString()}; }, 'Unit marked as signed off');
  } catch (e){ if (e && e.status) toast(storeMsg(e)); }
}

async function replaceBook(files){
  const u = S.units[S.unitId]; if (!files.length || !u) return;
  try {
    toast('Uploading completed pages…');
    const blob = await filesToPdfBlob(files);
    const r = await store.upload(`${u.folder}/Completed unit book.pdf`, blob, 'replace');
    await mutate(d => { d.units[u.id].book = r; d.units[u.id].pages = null; }, 'Completed book pages saved');
  } catch (e){ if (e && e.status) toast(storeMsg(e)); else if (e && e.message) toast(e.message); }
}

async function deleteEvidence(id){
  const e = S.ev[id]; if (!e) return;
  if (!confirm('Delete this evidence? The photos go to your OneDrive recycle bin.')) return;
  try {
    await mutate(d => { delete d.ev[id]; });
    for (const p of (e.photos||[])) await store.remove(p.id).catch(()=>{});
    S.modal = null; toast('Evidence deleted');
  } catch {}
}

async function moveEvidence(id, unitId, itemId){
  const e = S.ev[id]; const to = S.units[unitId]; if (!e || !to) return;
  try {
    await mutate(d => { d.ev[id].unitId = unitId; d.ev[id].itemId = itemId; }, 'Evidence moved');
    if (e.unitId !== unitId) for (const p of (e.photos||[])) await store.move(p.id, `${to.folder}/Photos`).catch(()=>{});
  } catch {}
}

async function deleteUnit(){
  const u = S.units[S.unitId]; if (!u) return;
  const evs = Object.values(S.ev).filter(e=>e.unitId===u.id);
  if (!confirm(`Delete “${unitLabel(u)}” and its ${evs.length} evidence record${evs.length===1?'':'s'}? Its OneDrive folder goes to your recycle bin.`)) return;
  try {
    await mutate(d => { delete d.units[u.id]; for (const e of evs) delete d.ev[e.id]; }, 'Unit deleted');
    await store.removePath(u.folder).catch(()=>{});
    S.view = 'units'; S.unitId = null; render();
  } catch {}
}

async function openFile(id){
  const w = window.open('', '_blank');
  try { const url = await store.webUrl(id); if (w) w.location = url; else location.href = url; }
  catch (e){ if (w) w.close(); toast(storeMsg(e)); }
}

/* ================= evidence pack ================= */
function clean(s){
  return String(s||'').normalize('NFD').replace(/[\u0300-\u036f]/g,'')
    .replace(/[\u2018\u2019]/g,"'").replace(/[\u201C\u201D]/g,'"').replace(/[\u2013\u2014]/g,'-').replace(/\u2026/g,'...').replace(/[^\x20-\x7E\n]/g,'');
}
function wrap(font, size, text, maxW){
  const out = [];
  for (const para of clean(text).split('\n')){
    let line = '';
    for (const w of para.split(/\s+/)){
      if (!w) continue; const t = line ? line + ' ' + w : w;
      if (font.widthOfTextAtSize(t, size) <= maxW){ line = t; continue; }
      if (line) out.push(line);
      let ww = w;
      while (font.widthOfTextAtSize(ww, size) > maxW){ let k = ww.length; while (k > 1 && font.widthOfTextAtSize(ww.slice(0,k), size) > maxW) k--; out.push(ww.slice(0,k)); ww = ww.slice(k); }
      line = ww;
    }
    out.push(line);
  }
  return out;
}

async function buildPack(includeBook, draft){
  const u = S.units[S.unitId]; if (!u) return;
  const s = unitStats(u);
  S.packBusy = true; render();
  try {
    const {PDFDocument, StandardFonts, rgb, degrees} = PDFLib;
    const doc = await PDFDocument.create();
    const F = await doc.embedFont(StandardFonts.Helvetica), FB = await doc.embedFont(StandardFonts.HelveticaBold);
    const W = 595.28, H = 841.89, M = 48, CW = W - 2*M;
    const ink = rgb(.09,.13,.17), mute = rgb(.36,.41,.47), rule = rgb(.78,.81,.85), red = rgb(.73,.25,.17);
    const ours = []; const newPage = () => { const p = doc.addPage([W,H]); ours.push(p); return p; };
    const lines = (pg, txt, x, y, font, size, maxW, color=ink, lh=size*1.3, maxLines=999) => { for (const l of wrap(font, size, txt, maxW).slice(0, maxLines)){ pg.drawText(l, {x, y:y-size, size, font, color}); y -= lh; } return y; };
    const P = S.profile; const label = unitLabel(u);

    // Cover
    let pg = newPage(); let y = H - M;
    pg.drawRectangle({x:0, y:H-10, width:W, height:10, color:rgb(1,.78,.16)});
    y = lines(pg, draft ? 'Evidence pack - DRAFT' : 'Evidence pack', M, y - 10, FB, 28, CW, ink, 34);
    y = lines(pg, label, M, y - 4, FB, 16, CW, ink, 21); y -= 18;
    const dates = Object.values(s.per).flat().map(e=>e.takenAt).filter(Boolean).sort();
    const rows = [['Apprentice', P.name], ['Workplace trainer', P.trainer], ['Workplace', P.workplace], ['Programme', P.programme], ['Unit standard', u.number], ['Credits', u.credits],
      ['Evidence dates', dates.length ? fmtDate(dates[0]) + (dates.length>1 ? ' to ' + fmtDate(dates[dates.length-1]) : '') : ''], ['Checklist', `${s.done} of ${s.total} items complete`], ['Pack printed', fmtDate(new Date().toISOString())]];
    for (const [k,v] of rows){
      pg.drawText(clean(k), {x:M, y:y-11, size:11, font:FB, color:mute});
      y = Math.min(y - 18, lines(pg, v || '-', M + 140, y, F, 11, CW - 140, ink, 15));
      pg.drawLine({start:{x:M,y:y+4}, end:{x:W-M,y:y+4}, thickness:.5, color:rule}); y -= 6;
    }
    y -= 16; y = lines(pg, 'Contents', M, y, FB, 13, CW);
    [draft ? 'Still to do' : null, (includeBook && u.book) ? 'Unit book pages' : null, 'Photo evidence, grouped by checklist item, with the apprentice\'s own notes', 'Workplace trainer / assessor sign-off']
      .filter(Boolean).forEach((c,i)=>{ y = lines(pg, `${i+1}. ${c}`, M, y - 2, F, 11, CW, ink, 15); });

    // Draft: what's missing
    if (draft){
      const miss = (u.items||[]).map((it,i)=>({it,i})).filter(x => !itemOk(u, x.it, s));
      pg = newPage(); y = H - M;
      y = lines(pg, 'Still to do', M, y, FB, 20, CW, ink, 26);
      y = lines(pg, miss.length ? `${miss.length} of ${s.total} checklist items have no evidence yet.` : 'Every checklist item has evidence. Print the final pack from the app.', M, y - 2, F, 11, CW, mute, 15); y -= 10;
      for (const {it,i} of miss){
        if (y < M + 40){ pg = newPage(); y = H - M; }
        pg.drawRectangle({x:M, y:y-13, width:12, height:12, borderColor:red, borderWidth:1.2});
        const yy = lines(pg, `${i+1}. ${it.text}`, M + 22, y, F, 10.5, CW - 130, ink, 14);
        pg.drawText(it.kind === 'book' ? 'Unit book not ticked' : 'Photo missing', {x:W-M-100, y:y-11, size:9, font:FB, color:red});
        y = yy - 8;
      }
    }

    // Unit book
    if (includeBook && u.book){
      try {
        const src = await PDFDocument.load(await store.bytes(u.book.id), {ignoreEncryption:true});
        let idx = src.getPageIndices(); if (u.pages) idx = idx.filter(i => i+1 >= u.pages[0] && i+1 <= u.pages[1]);
        (await doc.copyPages(src, idx)).forEach(p => doc.addPage(p));
      } catch { pg = newPage(); lines(pg, '[The unit book could not be added. Print it separately.]', M, H - M, FB, 12, CW, red); }
    }

    // Evidence
    const items = u.items || [];
    for (let i=0; i<items.length; i++){
      const it = items[i]; if (it.kind === 'book') continue;
      const blocks = (s.per[it.id]||[]).flatMap(e => (e.photos||[]).map((p,k)=>({e, pid:p.id, first:k===0, n:(e.photos||[]).length, k})));
      for (let b=0; b<blocks.length; b+=2){
        pg = newPage(); y = H - M;
        y = lines(pg, `Item ${i+1}${b?' (continued)':''}`, M, y, FB, 10, CW, mute, 13);
        y = lines(pg, it.text, M, y, FB, 13, CW, ink, 17, 3); y -= 8;
        const boxH = (y - M) / 2;
        for (const blk of blocks.slice(b, b+2)){
          const top = y; const imgH = boxH - 76;
          try {
            const ib = await store.bytes(blk.pid);
            let img; try { img = await doc.embedJpg(ib); } catch { img = await doc.embedPng(ib); }
            const sc = Math.min(CW/img.width, imgH/img.height);
            pg.drawImage(img, {x:M, y:top - img.height*sc, width:img.width*sc, height:img.height*sc});
            y = top - img.height*sc - 8;
          } catch { y = lines(pg, '[Photo could not be loaded]', M, top, F, 10, CW, mute); }
          y = lines(pg, [fmtDate(blk.e.takenAt), blk.e.job, blk.n>1 ? `photo ${blk.k+1} of ${blk.n}` : ''].filter(Boolean).join('  |  '), M, y, FB, 9.5, CW, mute, 13);
          if (blk.first) lines(pg, 'Apprentice\'s note: ' + (blk.e.note || '(no note)'), M, y, F, 10, CW, ink, 13, 3);
          y = top - boxH;
        }
      }
    }

    // Sign-off
    pg = newPage(); y = H - M;
    y = lines(pg, 'Workplace trainer / assessor sign-off', M, y, FB, 18, CW, ink, 23);
    y = lines(pg, label, M, y - 2, F, 11, CW, mute, 15); y -= 10;
    pg.drawText('Checked', {x:W-M-48, y:y-10, size:9, font:FB, color:mute}); y -= 16;
    for (let i=0; i<items.length; i++){
      const it = items[i]; const ls = wrap(F, 10.5, it.text, CW - 130);
      if (y - (ls.length*14 + 12) < M + 150){ pg = newPage(); y = H - M; }
      const evN = it.kind==='book' ? 'Unit book' : `${(s.per[it.id]||[]).reduce((a,e)=>a+(e.photos||[]).length,0)} photo(s)`;
      pg.drawText(`${i+1}.`, {x:M, y:y-11, size:10.5, font:FB, color:ink});
      let yy = y; for (const l of ls){ pg.drawText(l, {x:M+22, y:yy-11, size:10.5, font:F, color:ink}); yy -= 14; }
      pg.drawText(clean(evN), {x:W-M-118, y:y-11, size:9, font:F, color:mute});
      pg.drawRectangle({x:W-M-36, y:y-14, width:14, height:14, borderColor:ink, borderWidth:1});
      y = yy - 6; pg.drawLine({start:{x:M,y:y+2}, end:{x:W-M,y:y+2}, thickness:.5, color:rule}); y -= 6;
    }
    if (y < M + 210){ pg = newPage(); y = H - M; }
    y -= 18;
    for (const f of ['Trainer / assessor name', 'Signature', 'Date']){
      pg.drawText(f, {x:M, y:y-11, size:10.5, font:FB, color:ink});
      pg.drawLine({start:{x:M+150,y:y-13}, end:{x:W-M,y:y-13}, thickness:.8, color:ink}); y -= 34;
    }
    pg.drawText('Comments', {x:M, y:y-11, size:10.5, font:FB, color:ink}); y -= 30;
    for (let k=0; k<4; k++){ pg.drawLine({start:{x:M,y}, end:{x:W-M,y}, thickness:.6, color:rule}); y -= 26; }

    if (draft) ours.forEach(p => p.drawText('DRAFT - NOT COMPLETE', {x:110, y:250, size:52, font:FB, color:rgb(.75,.2,.15), opacity:.13, rotate:degrees(45)}));
    ours.forEach((p,i)=>{ p.drawText(clean(`${label}  -  ${P.name||'Apprentice'}  -  page ${i+1} of ${ours.length}`).slice(0,110), {x:M, y:24, size:8, font:F, color:mute}); });

    const out = await doc.save();
    const fname = safeName(`${u.number || u.title || 'Unit'} ${draft ? 'DRAFT ' : ''}evidence pack`) + '.pdf';
    localDownload(out, fname);
    try { await store.upload(`${u.folder}/${draft ? 'DRAFT evidence pack' : 'Evidence pack'}.pdf`, new Blob([out], {type:'application/pdf'}), 'replace'); toast((draft ? 'Draft' : 'Evidence') + ' pack downloaded and saved to OneDrive'); }
    catch { toast('Pack downloaded, but it couldn’t be saved to OneDrive.'); }
    S.modal = null;
  } catch (e){ toast('Couldn’t build the pack: ' + ((e && e.message) || 'unknown error')); }
  finally { S.packBusy = false; render(); }
}

/* ================= events ================= */
document.addEventListener('input', e => { const t = e.target; if (t.dataset && t.dataset.bind && t.tagName !== 'SELECT') setPath(S, t.dataset.bind, t.value); });

document.addEventListener('change', e => {
  const t = e.target;
  if (t.type === 'file'){
    const files = [...(t.files||[])]; t.value = '';
    if (!files.length || S.phase !== 'ready') return;
    if (t.id === 'f-book') addBook(files);
    else if ((t.id === 'f-cam' || t.id === 'f-gal') && S.cap.busy) return;
    else if (t.id === 'f-cam' || t.id === 'f-gal'){ S.cap.reading = true; render(); addPhotos(files); }
    else if (t.id === 'f-signed') uploadSigned(files[0]);
    else if (t.id === 'f-replbook') replaceBook(files);
    return;
  }
  if (t.id === 'cap-unit'){ S.cap.unitId = t.value; S.cap.itemId = ''; render(); return; }
  if (t.id === 'cap-item'){ S.cap.itemId = t.value; render(); return; }
  if (t.id === 'pk-book'){ if (S.modal) S.modal.book = t.checked; return; }
  if (t.id === 'mv-target'){ const [unitId, itemId] = t.value.split('|'); if (S.modal && unitId && itemId) moveEvidence(S.modal.id, unitId, itemId); return; }
  if (t.dataset && t.dataset.act === 'confirm-credits'){ const id = S.unitId, v = t.checked; mutate(d => { d.units[id].creditsConfirmed = v; d.units[id].confirmedAt = v ? new Date().toISOString() : null; }).catch(()=>{}); return; }
  if (t.dataset && t.dataset.act === 'book-toggle'){ const id = S.unitId, iid = t.dataset.item, v = t.checked; mutate(d => { d.units[id].bookDone = d.units[id].bookDone || {}; d.units[id].bookDone[iid] = v; }).catch(()=>{}); return; }
  if (t.dataset && t.dataset.bind && t.tagName === 'SELECT') setPath(S, t.dataset.bind, t.value);
});

document.addEventListener('click', e => {
  const el = e.target.closest('[data-act]'); if (!el || el.tagName === 'INPUT') return;
  const act = el.dataset.act;
  if (act === 'modal-bg'){ if (e.target === el){ S.modal = null; render(); } return; }
  switch (act){
    case 'signin': msalApp.loginRedirect({scopes:SCOPES, prompt:'select_account'}); break;
    case 'signout': if (msalApp) msalApp.logoutRedirect({account:msalApp.getActiveAccount()}); else location.reload(); break;
    case 'reload': location.reload(); break;
    case 'nav': S.view = el.dataset.v; S.editing = null; if (S.view === 'settings') S.pf = null; render(); break;
    case 'open-unit': S.unitId = el.dataset.id; S.view = 'unit'; render(); break;
    case 'add-stop': if (addCtl) addCtl.abort(); break;
    case 'rv-save': saveReview(); break;
    case 'rv-cancel': if (confirm('Discard this checklist?')){ S.add.review = null; render(); } break;
    case 'rv-unit-del': S.add.review.units.splice(+el.dataset.i, 1); if (!S.add.review.units.length) S.add.review = null; render(); break;
    case 'item-add': { const arr = getPath(S, el.dataset.base).items; arr.push({id:rid(), kind:'photo', text:''}); render(); const t = document.getElementById(el.dataset.base + '-t' + (arr.length-1)); if (t) t.focus(); break; }
    case 'item-del': getPath(S, el.dataset.base).items.splice(+el.dataset.i, 1); render(); break;
    case 'edit-open': { const u = S.units[S.unitId]; S.editing = {unitId:u.id, unit:JSON.parse(JSON.stringify({number:u.number||'', title:u.title||'', credits:u.credits||'', items:u.items||[]}))}; render(); break; }
    case 'edit-cancel': S.editing = null; render(); break;
    case 'edit-save': {
      const ed = S.editing; const items = ed.unit.items.filter(i=>String(i.text).trim()).map(i=>({id:i.id, kind:i.kind==='book'?'book':'photo', text:String(i.text).trim()}));
      mutate(d => { Object.assign(d.units[ed.unitId], {number:String(ed.unit.number).trim(), title:String(ed.unit.title).trim()||'Untitled unit', credits:String(ed.unit.credits).trim(), items}); }, 'Checklist saved').then(()=>{ S.editing = null; render(); }).catch(()=>{});
      break;
    }
    case 'unit-del': deleteUnit(); break;
    case 'unsign': if (confirm('Remove the signed copy? The unit goes back to “Ready for assessor”.')){ const id = S.unitId; mutate(d => { d.units[id].signed = null; d.units[id].creditsConfirmed = false; }).catch(()=>{}); } break;
    case 'cap-for': S.cap.unitId = el.dataset.unit; S.cap.itemId = el.dataset.item; S.cap.sugg = null; S.cap.msg = ''; if (el.dataset.job && S.plan.q) S.cap.job = S.plan.q; S.view = 'capture'; render(); break;
    case 'ph-del': { const p = S.cap.photos.splice(+el.dataset.i, 1)[0]; if (p) URL.revokeObjectURL(p.url); S.cap.sugg = null; render(); break; }
    case 'suggest': suggest(); break;
    case 'pick': S.cap.unitId = el.dataset.unit; S.cap.itemId = el.dataset.item; render(); break;
    case 'save-ev': saveEvidence(); break;
    case 'plan': runPlan(); break;
    case 'pf-save': {
      const {code, ...prof} = S.pf; try { localStorage.setItem('logbook-access-code', String(code||'').trim()); } catch {}
      mutate(d => { d.profile = Object.assign({}, d.profile, prof); }, 'Details saved').catch(()=>{}); break;
    }
    case 'open-root': store.rootUrl().then(u => { if (u) window.open(u, '_blank'); }).catch(e=>toast(storeMsg(e))); break;
    case 'open-file': openFile(el.dataset.id); break;
    case 'ev-open': S.modal = {type:'ev', id:el.dataset.id}; render(); break;
    case 'ev-del': deleteEvidence(S.modal.id); break;
    case 'modal-close': S.modal = null; render(); break;
    case 'pack-open': S.modal = {type:'pack', book:true, draft: !!el.dataset.draft}; render(); break;
    case 'pack-build': { const u = S.units[S.unitId]; buildPack(!!(u && u.book) && S.modal.book !== false, !!S.modal.draft); break; }
  }
});

start();
