/* ====================== STATE ====================== */
let pendingPhotos = [];   // {id, file, dataUrl, base64, mimeType}
let fise = [];            // extracted/edited fișe, loaded from localStorage
let openFisaId = null;
let templateArrayBuffer = null; // uploaded master .docx, raw bytes
let templateName = '';

const LS_FISE = 'pasuni_fise_v1';
const LS_KEY = 'pasuni_gemini_key';
const LS_MODEL = 'pasuni_gemini_model';
const LS_TEMPLATE = 'pasuni_template_b64';
const LS_TEMPLATE_NAME = 'pasuni_template_name';

/* ====================== INIT ====================== */
window.addEventListener('DOMContentLoaded', () => {
  loadFise();
  renderFiseList();
  updateReviewCount();
  bindUI();
  document.getElementById('apiKey').value = localStorage.getItem(LS_KEY) || '';
  document.getElementById('modelSelect').value = localStorage.getItem(LS_MODEL) || 'gemini-2.5-flash';
  loadTemplateFromStorage();
  if ('serviceWorker' in navigator) {
    navigator.serviceWorker.register('sw.js').catch(()=>{});
  }
});

function loadTemplateFromStorage(){
  const b64 = localStorage.getItem(LS_TEMPLATE);
  templateName = localStorage.getItem(LS_TEMPLATE_NAME) || '';
  if (b64){
    templateArrayBuffer = base64ToArrayBuffer(b64);
    updateTemplateStatus();
  }
}
function base64ToArrayBuffer(b64){
  const bin = atob(b64);
  const bytes = new Uint8Array(bin.length);
  for (let i=0;i<bin.length;i++) bytes[i] = bin.charCodeAt(i);
  return bytes.buffer;
}
function arrayBufferToBase64(buf){
  let binary = '';
  const bytes = new Uint8Array(buf);
  const chunk = 0x8000;
  for (let i=0;i<bytes.length;i+=chunk){
    binary += String.fromCharCode.apply(null, bytes.subarray(i,i+chunk));
  }
  return btoa(binary);
}
function updateTemplateStatus(){
  const el = document.getElementById('templateStatus');
  if (!el) return;
  el.textContent = templateArrayBuffer
    ? `Șablon încărcat: ${templateName}`
    : 'Niciun șablon Word încărcat — exportul va genera un document nou (aproximativ).';
  el.style.color = templateArrayBuffer ? '#2e5339' : '#b3413a';
}

function loadFise(){
  try { fise = JSON.parse(localStorage.getItem(LS_FISE) || '[]'); } catch(e){ fise = []; }
}
function saveFise(){
  localStorage.setItem(LS_FISE, JSON.stringify(fise));
  updateReviewCount();
}

/* ====================== TABS / SETTINGS ====================== */
function bindUI(){
  document.querySelectorAll('.tab-btn').forEach(btn=>{
    btn.addEventListener('click', ()=> switchTab(btn.dataset.tab));
  });
  document.getElementById('settingsBtn').addEventListener('click', ()=>{
    document.getElementById('panel-settings').style.display =
      document.getElementById('panel-settings').style.display === 'none' ? 'block' : 'none';
  });
  document.getElementById('closeSettings').addEventListener('click', ()=>{
    document.getElementById('panel-settings').style.display = 'none';
  });
  document.getElementById('saveSettings').addEventListener('click', ()=>{
    localStorage.setItem(LS_KEY, document.getElementById('apiKey').value.trim());
    localStorage.setItem(LS_MODEL, document.getElementById('modelSelect').value);
    toast('Setări salvate.');
    document.getElementById('panel-settings').style.display = 'none';
  });

  document.getElementById('uploadBox').addEventListener('click', ()=>{
    document.getElementById('fileInput').click();
  });
  document.getElementById('fileInput').addEventListener('change', onFilesChosen);
  document.getElementById('processBtn').addEventListener('click', processAllPhotos);
  document.getElementById('clearPhotosBtn').addEventListener('click', ()=>{
    pendingPhotos = []; renderPhotoGrid();
  });

  const tplInput = document.getElementById('templateInput');
  if (tplInput){
    tplInput.addEventListener('change', async (e)=>{
      const file = e.target.files[0];
      if (!file) return;
      const buf = await file.arrayBuffer();
      templateArrayBuffer = buf;
      templateName = file.name;
      try {
        localStorage.setItem(LS_TEMPLATE, arrayBufferToBase64(buf));
        localStorage.setItem(LS_TEMPLATE_NAME, templateName);
      } catch(err){
        toast('Șablon prea mare pentru stocare locală — va trebui reîncărcat la fiecare sesiune.');
      }
      updateTemplateStatus();
      toast('Șablon Word încărcat.');
    });
  }

  document.getElementById('genTextBtn').addEventListener('click', generateTextExport);
  document.getElementById('copyAllBtn').addEventListener('click', ()=>{
    const ta = document.getElementById('textOutput');
    ta.select(); document.execCommand('copy');
    toast('Copiat în clipboard.');
  });
  document.getElementById('genDocxBtn').addEventListener('click', generateDocx);
}

function switchTab(tab){
  document.querySelectorAll('.tab-btn').forEach(b=>b.classList.toggle('active', b.dataset.tab===tab));
  document.getElementById('tab-upload').style.display = tab==='upload' ? 'block':'none';
  document.getElementById('tab-review').style.display = tab==='review' ? 'block':'none';
  document.getElementById('tab-export').style.display = tab==='export' ? 'block':'none';
  if (tab==='review') renderFiseList();
}

function toast(msg){
  const host = document.getElementById('toastHost');
  const t = document.createElement('div');
  t.className='toast'; t.textContent=msg;
  host.appendChild(t);
  setTimeout(()=>t.remove(), 2600);
}

/* ====================== PHOTO UPLOAD ====================== */
function onFilesChosen(e){
  const files = Array.from(e.target.files || []);
  let remaining = files.length;
  if (!remaining) return;
  files.forEach(file=>{
    const reader = new FileReader();
    reader.onload = ()=>{
      const dataUrl = reader.result;
      const base64 = dataUrl.split(',')[1];
      pendingPhotos.push({
        id: 'p'+Date.now()+Math.random().toString(36).slice(2,7),
        file, dataUrl, base64, mimeType: file.type || 'image/jpeg', status: 'ready'
      });
      remaining--;
      if (remaining===0) renderPhotoGrid();
    };
    reader.readAsDataURL(file);
  });
  e.target.value = '';
}

function renderPhotoGrid(){
  const grid = document.getElementById('photoGrid');
  grid.innerHTML = '';
  pendingPhotos.forEach(p=>{
    const div = document.createElement('div');
    div.className='thumb';
    div.innerHTML = `<img src="${p.dataUrl}">
      <button class="rm" data-id="${p.id}">✕</button>
      <div class="status">${statusLabel(p.status)}</div>`;
    grid.appendChild(div);
  });
  grid.querySelectorAll('.rm').forEach(btn=>{
    btn.addEventListener('click', ()=>{
      pendingPhotos = pendingPhotos.filter(p=>p.id!==btn.dataset.id);
      renderPhotoGrid();
    });
  });
  document.getElementById('processBtn').disabled = pendingPhotos.length===0;
  document.getElementById('clearPhotosBtn').style.display = pendingPhotos.length ? 'inline-block':'none';
}
function statusLabel(s){
  return {ready:'gata', processing:'⏳ procesare...', done:'✓ extras', error:'✗ eroare'}[s] || s;
}

/* ====================== GEMINI EXTRACTION ====================== */
const SCHEMA_PROMPT = `Ești un asistent care transcrie fișe de teren scrise de mână pentru descrierea
parcelară a pajiștilor (Amenajament Pastoral, silvicultură din România). Primești o poză cu o pagină
dintr-un carnet de teren, format tabelar cu etichete în limba română (Tr. Păș., u.a., Supr.(ha),
Gr. funcț., Categ. folos., Unit. relief, Config. teren, T.S., T.P., Date stat. suplimentare, Încl.,
Exp., Altit., Unit. sol, Tip pajiște, Acoperire ierbacee, Gram (graminee), Leg (leguminoase),
Div. pl. (diverse plante), Pl. dăunătoare+toxice, Val. past., Arbuști, Gr. acop., Răsp., Veget.
forestieră, Date complementare, Lucr. exec., Lucr. propuse).

Rândurile de Gram/Leg/Div. pl./Pl. dăunătoare conțin un procent total, urmat de perechi (cod specie,
procent) — codul e un număr (ex: 31, 3, 205) care se referă la un cod dintr-un tabel de specii, NU
încerca tu să traduci codul în denumire de specie — întoarce doar cifra exact cum e scrisă. Cifrele
scrise de mână pot fi ambigue (mai ales 1/7, 3/8, 2/4, 0/6, 9/4) — dacă nu ești sigur, alege cea mai
plauzibilă și adaugă codul respectiv în lista "incert".

Simboluri frecvente: ○ = semințiș, Δ = nuieliș, □ = prăjiniș, Ψ = arbore izolat, simbol oiță = supratârlire.

Codul "Unit. relief" e un număr (30=versant inferior, 31=versant, 32=versant mijlociu, 33=versant superior,
11=luncă joasă, 12=luncă înaltă) — întoarce doar cifra. Codul "Config. teren" e o literă (A=plan, B=ondulat)
— întoarce doar litera. Codul "Unit. sol" e un cod numeric (ex: A108, 2212, 0209) — întoarce-l exact cum e scris.

Rândul "Arbuşti" NU conține text liber, ci o secvență de LITERE MARI (A-Z), fiecare literă fiind un cod de
specie de arbust dintr-un tabel separat (ex: "LINIK" înseamnă literele L, I, N, I, K citite separat, posibil
cu procente asociate). Întoarce-le ca listă de coduri-literă (fiecare literă = un element separat), NU ca text.

Întoarce STRICT un obiect JSON (fără markdown, fără text în plus), cu exact această structură:
{
 "u_a": "", "tr_pas": "", "supr_ha": "", "gr_funct": "", "categ_folos": "", "unit_relief": "",
 "config_teren": "", "t_s": "", "t_p": "",
 "date_stat_suplim": "",
 "incl": "", "exp": "", "alt": "", "unit_sol": "",
 "tip_pajiste": "", "acoperire_ierbacee": "",
 "gram_total": "", "gram": [{"cod":"","pct":""}],
 "leg_total": "", "leg": [{"cod":"","pct":""}],
 "div_pl_total": "", "div_pl": [{"cod":"","pct":""}],
 "daun_total": "", "daun": [{"cod":"","pct":""}],
 "val_past": "", "arbusti": [{"cod":"","pct":""}], "gr_acop": "", "rasp": "",
 "veget_forest": "", "date_compl": "", "lucr_exec": "", "lucr_propuse": "",
 "incert": []
}
Câmpurile pe care nu le poți citi rămân string gol "". Nu inventa date. Scrie procentele cu simbolul %
(ex "68%"). Pentru "alt" scrie ca "540 - 670 m".`;

async function processAllPhotos(){
  const key = localStorage.getItem(LS_KEY);
  const model = localStorage.getItem(LS_MODEL) || 'gemini-2.5-flash';
  if (!key){
    toast('Adaugă mai întâi cheia API Gemini din Setări (⚙️).');
    document.getElementById('panel-settings').style.display = 'block';
    return;
  }
  document.getElementById('processBtn').disabled = true;
  const hint = document.getElementById('processHint');
  let done=0;
  for (const p of pendingPhotos){
    p.status='processing'; renderPhotoGrid();
    hint.textContent = `Procesare ${++done}/${pendingPhotos.length}...`;
    try {
      const data = await callGemini(key, model, p.base64, p.mimeType);
      const fisa = normalizeFisa(data);
      fisa.id = 'f'+Date.now()+Math.random().toString(36).slice(2,7);
      fisa.thumb = p.dataUrl;
      fise.push(fisa);
      p.status='done';
    } catch(err){
      console.error(err);
      p.status='error';
      toast('Eroare la o poză: ' + err.message);
    }
    renderPhotoGrid();
  }
  saveFise();
  hint.textContent = `Gata. ${fise.length} fișe în total — vezi tab "Fișe".`;
  document.getElementById('processBtn').disabled = pendingPhotos.length===0;
  renderFiseList();
}

async function callGemini(key, model, base64, mimeType){
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent`;
  const body = {
    contents: [{
      parts: [
        { inline_data: { mime_type: mimeType, data: base64 } },
        { text: SCHEMA_PROMPT }
      ]
    }],
    generationConfig: {
      responseMimeType: "application/json",
      temperature: 0.1
    }
  };
  const res = await fetch(url, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json', 'x-goog-api-key': key },
    body: JSON.stringify(body)
  });
  if (!res.ok){
    const txt = await res.text();
    throw new Error(`Gemini ${res.status}: ${txt.slice(0,200)}`);
  }
  const json = await res.json();
  const text = json?.candidates?.[0]?.content?.parts?.[0]?.text;
  if (!text) throw new Error('Răspuns gol de la Gemini.');
  return JSON.parse(text);
}

function normalizeFisa(d){
  const safeArr = (a)=> Array.isArray(a) ? a.filter(x=>x && (x.cod||x.pct)) : [];
  return {
    u_a: d.u_a||'', tr_pas: d.tr_pas||'', supr_ha: d.supr_ha||'', gr_funct: d.gr_funct||'FP',
    categ_folos: d.categ_folos||'', unit_relief: d.unit_relief||'', config_teren: d.config_teren||'',
    t_s: d.t_s||'', t_p: d.t_p||'',
    date_stat_suplim: d.date_stat_suplim||'',
    incl: d.incl||'', exp: d.exp||'', alt: d.alt||'', unit_sol: d.unit_sol||'',
    tip_pajiste: d.tip_pajiste||'', acoperire_ierbacee: d.acoperire_ierbacee||'',
    gram_total: d.gram_total||'', gram: safeArr(d.gram),
    leg_total: d.leg_total||'', leg: safeArr(d.leg),
    div_pl_total: d.div_pl_total||'', div_pl: safeArr(d.div_pl),
    daun_total: d.daun_total||'', daun: safeArr(d.daun),
    val_past: d.val_past||'', arbusti: safeArr(d.arbusti), gr_acop: d.gr_acop||'', rasp: d.rasp||'',
    veget_forest: d.veget_forest||'', date_compl: d.date_compl||'',
    lucr_exec: d.lucr_exec||'', lucr_propuse: d.lucr_propuse||'',
    incert: Array.isArray(d.incert) ? d.incert : []
  };
}

/* ====================== LOOKUPS (plante / sol / arbuşti) ====================== */
function speciesName(cod){
  if (!cod) return '';
  const c = String(cod).trim();
  const p = CODEBOOK.plants[c];
  return p ? p.sci : null;
}
function soilName(cod){
  if (!cod) return null;
  const c = String(cod).trim();
  return CODEBOOK.soils[c] || null;
}
function shrubInfo(letter){
  if (!letter) return null;
  const c = String(letter).trim().toUpperCase();
  return CODEBOOK.shrubs[c] || null;
}
function reliefName(cod){
  if (!cod) return null;
  return CODEBOOK.relief[String(cod).trim()] || null;
}
function configTerenName(cod){
  if (!cod) return null;
  return CODEBOOK.config_teren[String(cod).trim().toUpperCase()] || null;
}

/* ====================== REVIEW LIST ====================== */
function updateReviewCount(){
  const el = document.getElementById('reviewCount');
  el.textContent = fise.length ? `(${fise.length})` : '';
}

function renderFiseList(){
  const host = document.getElementById('fiseList');
  if (!fise.length){
    host.innerHTML = `<div class="empty">Nicio fișă încă.<br>Adaugă poze în tab-ul "Poze noi" și apasă "Extrage date cu AI".</div>`;
    return;
  }
  host.innerHTML = '';
  fise.forEach(f=>{
    const card = document.createElement('div');
    card.className='card fisa-card';
    const warnCount = (f.incert||[]).length;
    card.innerHTML = `
      <div class="fisa-head" data-id="${f.id}">
        <div>
          <span class="tag">u.a. ${esc(f.u_a)||'?'}</span>
          <span class="meta"> ${esc(f.supr_ha)} ha · ${esc(f.tip_pajiste)}</span>
          ${warnCount ? `<span class="badge-warn">${warnCount} de verificat</span>`:''}
        </div>
        <div>▾</div>
      </div>
      <div class="fisa-body ${openFisaId===f.id?'open':''}" id="body-${f.id}"></div>
    `;
    host.appendChild(card);
    card.querySelector('.fisa-head').addEventListener('click', ()=>{
      openFisaId = openFisaId===f.id ? null : f.id;
      renderFiseList();
    });
    if (openFisaId===f.id){
      renderFisaBody(document.getElementById(`body-${f.id}`), f);
    }
  });
}

function esc(s){ return (s||'').toString().replace(/</g,'&lt;'); }

function renderFisaBody(container, f){
  container.classList.add('open');
  container.innerHTML = `
    <div class="section-title">Identificare</div>
    <div class="field-grid">
      ${fld(f,'u_a','u.a.')} ${fld(f,'tr_pas','Tr. Păș.')}
      ${fld(f,'supr_ha','Supr. (ha)')} ${fld(f,'gr_funct','Gr. funcț.')}
      ${fld(f,'categ_folos','Categ. folos.')} ${fld(f,'unit_relief','Unit. relief')}
      ${fld(f,'config_teren','Config. teren')} ${fld(f,'t_s','T.S.')}
      ${fld(f,'t_p','T.P.')}
    </div>
    <div class="section-title">Stațiune</div>
    <div class="field-grid">
      ${fld(f,'incl','Încl.')} ${fld(f,'exp','Exp.')}
      ${fld(f,'alt','Altit.')} ${fld(f,'unit_sol','Unit. sol')}
    </div>
    <label>Date stat. suplimentare</label>
    <textarea data-field="date_stat_suplim">${esc(f.date_stat_suplim)}</textarea>
    <div class="section-title">Vegetație erbacee</div>
    <div class="field-grid">
      ${fld(f,'tip_pajiste','Tip pajiște')} ${fld(f,'acoperire_ierbacee','Acoperire ierbacee')}
    </div>
    ${speciesBlock(f,'gram','Gram (graminee)')}
    ${speciesBlock(f,'leg','Leg (leguminoase)')}
    ${speciesBlock(f,'div_pl','Div. pl.')}
    ${speciesBlock(f,'daun','Pl. dăunătoare+toxice')}
    <div class="section-title">Pășunat</div>
    <div class="field-grid">
      ${fld(f,'val_past','Val. past.')} ${fld(f,'gr_acop','Gr. acop.')}
      ${fld(f,'rasp','Răsp.')}
    </div>
    ${speciesBlock(f,'arbusti','Arbuști (coduri literă A-Z)', {lookup:'shrub', showTotal:false})}
    <label>Veget. forestieră</label>
    <textarea data-field="veget_forest">${esc(f.veget_forest)}</textarea>
    <label>Date compl.</label>
    <textarea data-field="date_compl">${esc(f.date_compl)}</textarea>
    <label>Lucr. exec.</label>
    <textarea data-field="lucr_exec">${esc(f.lucr_exec)}</textarea>
    <label>Lucr. propuse</label>
    <textarea data-field="lucr_propuse">${esc(f.lucr_propuse)}</textarea>
    <div class="btnbar">
      <button class="btn small secondary dup-btn">📋 Copiază text (fișă)</button>
      <button class="btn small danger del-btn">🗑 Șterge fișă</button>
    </div>
  `;
  // bind simple field inputs
  container.querySelectorAll('input[data-field], textarea[data-field]').forEach(inp=>{
    inp.addEventListener('input', ()=>{
      f[inp.dataset.field] = inp.value;
      saveFise();
    });
  });
  // bind species rows
  ['gram','leg','div_pl','daun'].forEach(key=> bindSpeciesGroup(container, f, key, speciesName));
  bindSpeciesGroup(container, f, 'arbusti', codLookupShrub);
  container.querySelector('.dup-btn').addEventListener('click', ()=>{
    navigator.clipboard.writeText(fisaToText(f));
    toast('Text copiat pentru u.a. ' + (f.u_a||'?'));
  });
  container.querySelector('.del-btn').addEventListener('click', ()=>{
    if (!confirm('Ștergi fișa u.a. ' + (f.u_a||'?') + '?')) return;
    fise = fise.filter(x=>x.id!==f.id);
    saveFise(); renderFiseList();
  });
}

function fld(f, key, label){
  return `<div><label>${label}</label><input type="text" data-field="${key}" value="${esc(f[key])}"></div>`;
}

function codLookupShrub(cod){
  const info = shrubInfo(cod);
  return info ? `${info.sci}${info.pop ? ' ('+info.pop+')' : ''}` : null;
}

function speciesBlock(f, key, label, opts={}){
  const showTotal = opts.showTotal !== false;
  const lookup = opts.lookup === 'shrub' ? codLookupShrub : speciesName;
  const totalKey = key+'_total';
  const list = f[key] || [];
  let rows = list.map((s,i)=> speciesRowHtml(key, i, s, lookup)).join('');
  return `
    <div class="section-title">${label} ${showTotal ? `<input type="text" style="width:60px;display:inline-block;margin-left:6px;"
       data-total="${key}" value="${esc(f[totalKey])}">` : ''}</div>
    <div class="sp-list" data-list="${key}">${rows}</div>
    <button class="add-sp" data-add="${key}">+ adaugă ${opts.lookup==='shrub'?'arbust':'specie'}</button>
  `;
}

function speciesRowHtml(key, i, s, lookup){
  const name = lookup(s.cod);
  const nameHtml = s.cod ? (name || 'cod necunoscut') : '';
  const cls = (s.cod && !name) ? 'unknown' : '';
  return `<div class="sp-row" data-idx="${i}">
    <input class="cod" type="text" data-sp="cod" value="${esc(s.cod)}" placeholder="cod">
    <span class="name ${cls}">${esc(nameHtml)}</span>
    <input class="pct" type="text" data-sp="pct" value="${esc(s.pct)}" placeholder="%">
    <button class="rm-sp" title="Șterge">✕</button>
  </div>`;
}

function bindSpeciesGroup(container, f, key, lookup){
  const totalInput = container.querySelector(`input[data-total="${key}"]`);
  if (totalInput){
    totalInput.addEventListener('input', ()=>{ f[key+'_total'] = totalInput.value; saveFise(); });
  }
  const listEl = container.querySelector(`.sp-list[data-list="${key}"]`);
  if (!listEl) return;
  function rebind(){
    listEl.querySelectorAll('.sp-row').forEach(row=>{
      const idx = parseInt(row.dataset.idx);
      row.querySelectorAll('input[data-sp]').forEach(inp=>{
        inp.addEventListener('input', ()=>{
          f[key][idx][inp.dataset.sp] = inp.value;
          if (inp.dataset.sp==='cod'){
            const nm = lookup(inp.value);
            const nameSpan = row.querySelector('.name');
            nameSpan.textContent = inp.value ? (nm || 'cod necunoscut') : '';
            nameSpan.classList.toggle('unknown', !!(inp.value && !nm));
          }
          saveFise();
        });
      });
      row.querySelector('.rm-sp').addEventListener('click', ()=>{
        f[key].splice(idx,1);
        saveFise();
        listEl.innerHTML = f[key].map((s,i)=>speciesRowHtml(key,i,s,lookup)).join('');
        rebind();
      });
    });
  }
  rebind();
  const addBtn = container.querySelector(`button[data-add="${key}"]`);
  addBtn.addEventListener('click', ()=>{
    f[key].push({cod:'', pct:''});
    saveFise();
    listEl.innerHTML = f[key].map((s,i)=>speciesRowHtml(key,i,s,lookup)).join('');
    rebind();
  });
}

/* ====================== TEXT EXPORT ====================== */
function speciesListText(list, lookup){
  lookup = lookup || speciesName;
  const parts = list.filter(s=>s.cod).map(s=>{
    const nm = lookup(s.cod) || `(cod ${s.cod} necunoscut)`;
    return `${nm} ${s.pct||''}`.trim();
  });
  return parts.join('; ');
}

function fisaToText(f){
  const lines = [];
  const relief = reliefName(f.unit_relief);
  const conf = configTerenName(f.config_teren);
  const sol = soilName(f.unit_sol);
  lines.push(`===== u.a. ${f.u_a||'?'} (${f.tr_pas? 'Tr.Păș '+f.tr_pas+', ':''}${f.supr_ha||'?'} ha) =====`);
  lines.push(`Header: u.a=${f.u_a} | Supr=${f.supr_ha} | Gr.funcț=${f.gr_funct} | Categ.folos=${f.categ_folos} | Unit.relief=${f.unit_relief}${relief?' ('+relief+')':''} | Config.teren=${f.config_teren}${conf?' ('+conf+')':''} | T.S.=${f.t_s} | T.P.=${f.t_p}`);
  if (f.date_stat_suplim) lines.push(`Date stat suplim: ${f.date_stat_suplim}`);
  lines.push(`Încl: ${f.incl}   Exp: ${f.exp}   Alt: ${f.alt}   Unit.sol: ${f.unit_sol}${sol?' ('+sol+')':''}`);
  lines.push(`Tip pajiște: ${f.tip_pajiste}   Acoperire ierbacee: ${f.acoperire_ierbacee}`);
  if (f.gram && f.gram.length) lines.push(`Gram: ${f.gram_total} (${speciesListText(f.gram)})`);
  if (f.leg && f.leg.length) lines.push(`Leg: ${f.leg_total} (${speciesListText(f.leg)})`);
  if (f.div_pl && f.div_pl.length) lines.push(`Div. pl.: ${f.div_pl_total} (${speciesListText(f.div_pl)})`);
  if (f.daun && f.daun.length) lines.push(`Pl. dăunătoare+toxice: ${f.daun_total} (${speciesListText(f.daun)})`);
  const arbTxt = f.arbusti && f.arbusti.length ? speciesListText(f.arbusti, codLookupShrub) : '';
  lines.push(`Val. past: ${f.val_past}   Arbuști: ${arbTxt}   Gr. acop.: ${f.gr_acop}   Răsp.: ${f.rasp}`);
  if (f.veget_forest) lines.push(`Veget. forest.: ${f.veget_forest}`);
  if (f.date_compl) lines.push(`Date compl.: ${f.date_compl}`);
  if (f.lucr_exec) lines.push(`Lucr. exec.: ${f.lucr_exec}`);
  if (f.lucr_propuse) lines.push(`Lucr. propuse: ${f.lucr_propuse}`);
  return lines.join('\n');
}

function generateTextExport(){
  if (!fise.length){ toast('Nicio fișă de exportat.'); return; }
  const all = fise.map(fisaToText).join('\n\n');
  const ta = document.getElementById('textOutput');
  ta.value = all; ta.style.display='block';
  document.getElementById('copyBtnBar').style.display='flex';
}

/* ====================== DOCX EXPORT ====================== */
async function generateDocx(){
  if (!fise.length){ toast('Nicio fișă de exportat.'); return; }
  if (templateArrayBuffer){
    await generateDocxFromTemplate();
  } else {
    await generateDocxFallback();
  }
}

async function generateDocxFromTemplate(){
  toast('Se completează șablonul...');
  try {
    await DocxEditor.load(templateArrayBuffer.slice(0)); // slice(0) = don't detach original buffer
  } catch(err){
    toast('Eroare la citirea șablonului: ' + err.message);
    return;
  }
  const lookups = { plant: speciesName, soil: soilName, shrub: shrubInfo, relief: reliefName, config: configTerenName };
  const notFound = [];
  const filled = [];
  fise.forEach(f=>{
    const res = DocxEditor.fillFisa(f, lookups);
    if (res.ok) filled.push(f.u_a); else notFound.push(f.u_a);
  });
  const blob = await DocxEditor.exportBlob();
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url; a.download = (templateName.replace(/\.docx?$/i,'') || 'sablon') + '_completat.docx';
  document.body.appendChild(a); a.click(); a.remove();
  URL.revokeObjectURL(url);
  if (notFound.length){
    toast(`Descărcat. Atenție: u.a. ${notFound.join(', ')} nu au fost găsite în șablon (au fost sărite).`);
  } else {
    toast(`Descărcat — ${filled.length} fișe completate direct în șablonul tău.`);
  }
}

async function generateDocxFallback(){
  toast('Nu ai încărcat un șablon — se generează un document nou (aproximativ).');
  const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
          WidthType, BorderStyle, VerticalAlign, ShadingType } = docx;

  const headerShade = { fill: "D9E2D9", type: ShadingType.CLEAR, color: "auto" };

  function mergedRow(children, opts={}){
    return new TableRow({
      children: [ new TableCell({
        children: [ new Paragraph({ children }) ],
        columnSpan: 8,
        margins: {top:60,bottom:60,left:80,right:80},
        shading: opts.shade ? headerShade : undefined
      })]
    });
  }
  function headerCell(text){
    return new TableCell({
      children:[new Paragraph({children:[new TextRun({text, bold:true})], alignment:'center'})],
      shading: headerShade, verticalAlign: VerticalAlign.CENTER,
      margins:{top:60,bottom:60,left:40,right:40}
    });
  }
  function valCell(text){
    return new TableCell({
      children:[new Paragraph({children:[new TextRun({text: text||''})], alignment:'center'})],
      verticalAlign: VerticalAlign.CENTER, margins:{top:60,bottom:60,left:40,right:40}
    });
  }
  function labelRuns(label, value){
    return [ new TextRun({text: label, bold:true}), new TextRun({text: value||''}) ];
  }
  function speciesRuns(total, list){
    const runs = [ new TextRun({text: (total||'')+' (' , bold:false}) ];
    const items = (list||[]).filter(s=>s.cod);
    items.forEach((s,i)=>{
      const nm = speciesName(s.cod) || `cod ${s.cod}`;
      runs.push(new TextRun({text: nm, italics:true}));
      runs.push(new TextRun({text: ` ${s.pct||''}` + (i<items.length-1 ? '; ' : '')}));
    });
    runs.push(new TextRun({text:')'}));
    return runs;
  }

  const tables = [];
  fise.forEach(f=>{
    const rows = [];
    rows.push(new TableRow({children:[
      headerCell('Tr. Păș.'), headerCell('u.a.'), headerCell('Supr.(ha)'), headerCell('Gr.funcț'),
      headerCell('Categ.folos'), headerCell('Unit.relief'), headerCell('Config.teren'), headerCell('T.S.')
    ]}));
    rows.push(new TableRow({children:[
      valCell(f.tr_pas), valCell(f.u_a), valCell(f.supr_ha), valCell(f.gr_funct),
      valCell(f.categ_folos), valCell(f.unit_relief), valCell(f.config_teren), valCell(f.t_s)
    ]}));
    rows.push(mergedRow(labelRuns('Încl:  ', f.incl+'    Exp: '+f.exp+'    Alt: '+f.alt+'    Unit.sol: '+f.unit_sol)));
    rows.push(mergedRow(labelRuns('Date stat suplim: ', f.date_stat_suplim)));
    rows.push(mergedRow(labelRuns('Tip pajişte: ', f.tip_pajiste+'     Acoperire ierbacee: '+f.acoperire_ierbacee)));
    rows.push(mergedRow([new TextRun({text:'Gram: ', bold:true}), ...speciesRuns(f.gram_total, f.gram)]));
    rows.push(mergedRow([new TextRun({text:'Leg: ', bold:true}), ...speciesRuns(f.leg_total, f.leg)]));
    rows.push(mergedRow([new TextRun({text:'Div. pl.: ', bold:true}), ...speciesRuns(f.div_pl_total, f.div_pl)]));
    rows.push(mergedRow([new TextRun({text:'Pl. dăunătoare+toxice: ', bold:true}), ...speciesRuns(f.daun_total, f.daun)]));
    const arbTxt = (f.arbusti||[]).filter(s=>s.cod).map(s=> (shrubInfo(s.cod)?.sci) || `cod ${s.cod}`).join(', ');
    rows.push(mergedRow(labelRuns('Val. past: ', (f.val_past||'')+'   Arbuşti: '+arbTxt+'   Gr. acop.: '+(f.gr_acop||'')+'   Răsp.: '+(f.rasp||''))));
    rows.push(mergedRow(labelRuns('Veget. forest.: ', f.veget_forest)));
    rows.push(mergedRow(labelRuns('Date compl.: ', f.date_compl)));
    rows.push(mergedRow(labelRuns('Lucr. exec.: ', f.lucr_exec)));
    rows.push(mergedRow(labelRuns('Lucr. propuse: ', f.lucr_propuse)));

    tables.push(new Table({
      rows, width:{size:100, type: WidthType.PERCENTAGE}
    }));
    tables.push(new Paragraph({text:''}));
  });

  const doc = new Document({ sections:[{ properties:{}, children: tables.flatMap(t=>[t]) }] });
  const blob = await Packer.toBlob(doc);
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url; a.download = 'fise_pasuni_' + new Date().toISOString().slice(0,10) + '.docx';
  document.body.appendChild(a); a.click(); a.remove();
  URL.revokeObjectURL(url);
  toast('Document descărcat.');
}
