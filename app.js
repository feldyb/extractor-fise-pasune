/* ====================== STATE ====================== */
let pendingPhotos = [];   // {id, file, dataUrl, base64, mimeType}
let fise = [];            // extracted/edited fișe, loaded from localStorage
let openFisaId = null;
let templateArrayBuffer = null; // uploaded master .docx, raw bytes
let templateName = '';

const LS_FISE = 'pasuni_fise_v1';
const LS_PROVIDER = 'pasuni_provider';
const LS_KEY_GEMINI = 'pasuni_gemini_key';
const LS_MODEL_GEMINI = 'pasuni_gemini_model';
const LS_KEY_CLAUDE = 'pasuni_claude_key';
const LS_MODEL_CLAUDE = 'pasuni_claude_model';
const LS_KEY_OPENROUTER = 'pasuni_openrouter_key';
const LS_MODEL_OPENROUTER = 'pasuni_openrouter_model';
const LS_PACING = 'pasuni_pacing_sec';
const LS_TEMPLATE = 'pasuni_template_b64';
const LS_TEMPLATE_NAME = 'pasuni_template_name';

/* ====================== INIT ====================== */
window.addEventListener('DOMContentLoaded', () => {
  loadFise();
  renderFiseList();
  updateReviewCount();
  bindUI();
  document.getElementById('providerSelect').value = localStorage.getItem(LS_PROVIDER) || 'gemini';
  document.getElementById('apiKeyGemini').value = localStorage.getItem(LS_KEY_GEMINI) || '';
  document.getElementById('modelSelectGemini').value = localStorage.getItem(LS_MODEL_GEMINI) || 'gemini-3.5-flash';
  document.getElementById('apiKeyClaude').value = localStorage.getItem(LS_KEY_CLAUDE) || '';
  document.getElementById('modelSelectClaude').value = localStorage.getItem(LS_MODEL_CLAUDE) || 'claude-sonnet-5';
  document.getElementById('apiKeyOpenrouter').value = localStorage.getItem(LS_KEY_OPENROUTER) || '';
  const savedOrModel = localStorage.getItem(LS_MODEL_OPENROUTER) || 'google/gemma-4-31b-it:free';
  const orSelect = document.getElementById('modelSelectOpenrouter');
  const knownOr = Array.from(orSelect.options).some(o=>o.value===savedOrModel);
  if (knownOr){
    orSelect.value = savedOrModel;
  } else {
    orSelect.value = 'custom';
    document.getElementById('modelCustomOpenrouter').value = savedOrModel;
    document.getElementById('modelCustomOpenrouter').style.display = 'block';
  }
  document.getElementById('pacingInput').value = localStorage.getItem(LS_PACING) || '7';
  updateProviderPanels();
  loadTemplateFromStorage();
  renderLegend();
  if ('serviceWorker' in navigator) {
    navigator.serviceWorker.register('sw.js').catch(()=>{});
  }
});

function updateProviderPanels(){
  const provider = document.getElementById('providerSelect').value;
  document.getElementById('provider-gemini').style.display = provider==='gemini' ? 'block':'none';
  document.getElementById('provider-claude').style.display = provider==='claude' ? 'block':'none';
  document.getElementById('provider-openrouter').style.display = provider==='openrouter' ? 'block':'none';
  const pacingHint = document.getElementById('pacingHint');
  if (provider==='claude'){
    pacingHint.textContent = 'Claude nu are cotă gratuită/limită de tip "cotă zilnică" — poți lăsa 1-2s doar ca să eviți burst-uri de rate-limit.';
  } else if (provider==='openrouter'){
    pacingHint.textContent = 'Cont OpenRouter fără credit: 20 cereri/min, 50/zi. Cu 10$ credit adăugat: 1.000/zi pe modele gratuite — poți lăsa 3-4s.';
  } else {
    pacingHint.textContent = 'Nivel gratuit Gemini: ~10-15 cereri/minut → lasă 6-7s. Dacă ai activat facturarea (billing), poți coborî la 1-2s — limita urcă mult.';
  }
}

function getActiveProviderConfig(){
  const provider = localStorage.getItem(LS_PROVIDER) || 'gemini';
  if (provider === 'claude'){
    return { provider, key: localStorage.getItem(LS_KEY_CLAUDE), model: localStorage.getItem(LS_MODEL_CLAUDE) || 'claude-sonnet-5' };
  }
  if (provider === 'openrouter'){
    return { provider, key: localStorage.getItem(LS_KEY_OPENROUTER), model: localStorage.getItem(LS_MODEL_OPENROUTER) || 'google/gemma-4-31b-it:free' };
  }
  return { provider, key: localStorage.getItem(LS_KEY_GEMINI), model: localStorage.getItem(LS_MODEL_GEMINI) || 'gemini-3.5-flash' };
}

function renderLegend(){
  const host = document.getElementById('legendHost');
  if (!host) return;
  const symbols = Object.values(CODEBOOK.symbols).map(v=>`• ${v}`).join('<br>');
  const eroziune = Object.entries(CODEBOOK.eroziune).map(([k,v])=>`<b>${k}</b> — ${v}`).join('<br>');
  const lucrari = Object.entries(CODEBOOK.lucrari).map(([k,v])=>`<b>${k}</b> — ${v}`).join('<br>');
  host.innerHTML = `<div style="margin-bottom:10px;"><b style="color:var(--accent);">Simboluri (Veget. forest. / Date compl.)</b><br>${symbols}</div>
    <div style="margin-bottom:10px;"><b style="color:var(--accent);">Eroziune (Date stat. suplim.)</b><br>${eroziune}</div>
    <div><b style="color:var(--accent);">Lucrări (Lucr. exec. / Lucr. propuse)</b><br>${lucrari}</div>`;
}

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
  document.getElementById('providerSelect').addEventListener('change', updateProviderPanels);
  document.getElementById('modelSelectOpenrouter').addEventListener('change', (e)=>{
    document.getElementById('modelCustomOpenrouter').style.display = e.target.value==='custom' ? 'block':'none';
  });
  document.getElementById('saveSettings').addEventListener('click', ()=>{
    localStorage.setItem(LS_PROVIDER, document.getElementById('providerSelect').value);
    localStorage.setItem(LS_KEY_GEMINI, document.getElementById('apiKeyGemini').value.trim());
    localStorage.setItem(LS_MODEL_GEMINI, document.getElementById('modelSelectGemini').value);
    localStorage.setItem(LS_KEY_CLAUDE, document.getElementById('apiKeyClaude').value.trim());
    localStorage.setItem(LS_MODEL_CLAUDE, document.getElementById('modelSelectClaude').value);
    localStorage.setItem(LS_KEY_OPENROUTER, document.getElementById('apiKeyOpenrouter').value.trim());
    const orModelSel = document.getElementById('modelSelectOpenrouter').value;
    const orModel = orModelSel==='custom' ? document.getElementById('modelCustomOpenrouter').value.trim() : orModelSel;
    localStorage.setItem(LS_MODEL_OPENROUTER, orModel);
    localStorage.setItem(LS_PACING, document.getElementById('pacingInput').value);
    toast('Setări salvate.');
    document.getElementById('panel-settings').style.display = 'none';
  });

  document.getElementById('cameraBtn').addEventListener('click', ()=>{
    document.getElementById('fileInput').click();
  });
  document.getElementById('galleryBtn').addEventListener('click', ()=>{
    document.getElementById('galleryInput').click();
  });
  document.getElementById('fileInput').addEventListener('change', onFilesChosen);
  document.getElementById('galleryInput').addEventListener('change', onFilesChosen);
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
  document.getElementById('addManualBtn').addEventListener('click', addManualFisa);
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
const MAX_DIM = 1600; // max width/height sent to the AI — plenty for reading handwriting, much faster
const JPEG_QUALITY = 0.85;

function resizeImage(file){
  return new Promise((resolve, reject)=>{
    const img = new Image();
    const objUrl = URL.createObjectURL(file);
    img.onload = ()=>{
      let { width, height } = img;
      if (width > MAX_DIM || height > MAX_DIM){
        const scale = MAX_DIM / Math.max(width, height);
        width = Math.round(width * scale);
        height = Math.round(height * scale);
      }
      const canvas = document.createElement('canvas');
      canvas.width = width; canvas.height = height;
      const ctx = canvas.getContext('2d');
      ctx.drawImage(img, 0, 0, width, height);
      URL.revokeObjectURL(objUrl);
      const dataUrl = canvas.toDataURL('image/jpeg', JPEG_QUALITY);
      resolve(dataUrl);
    };
    img.onerror = (e)=>{ URL.revokeObjectURL(objUrl); reject(e); };
    img.src = objUrl;
  });
}

function onFilesChosen(e){
  const files = Array.from(e.target.files || []);
  let remaining = files.length;
  if (!remaining) return;
  const hint = document.getElementById('processHint');
  hint.textContent = `Se pregătesc ${remaining} poze...`;
  files.forEach(async (file)=>{
    try {
      const dataUrl = await resizeImage(file); // resized JPEG, sent to AI (fast)
      const base64 = dataUrl.split(',')[1];
      pendingPhotos.push({
        id: 'p'+Date.now()+Math.random().toString(36).slice(2,7),
        file, dataUrl, base64, mimeType: 'image/jpeg', status: 'ready'
      });
    } catch(err){
      console.error('Eroare redimensionare, folosesc originalul:', err);
      // fallback: use original file as-is if resize fails
      const reader = new FileReader();
      const dataUrl = await new Promise(res=>{ reader.onload=()=>res(reader.result); reader.readAsDataURL(file); });
      pendingPhotos.push({
        id: 'p'+Date.now()+Math.random().toString(36).slice(2,7),
        file, dataUrl, base64: dataUrl.split(',')[1], mimeType: file.type||'image/jpeg', status: 'ready'
      });
    }
    remaining--;
    if (remaining===0){ hint.textContent=''; renderPhotoGrid(); }
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

Simboluri frecvente (apar mai ales la "Veget. forest." și "Date compl."):
- Δ (triunghi) = nuieliș
- □ (pătrat) = prăjiniș
- ○ (cerc) = semințiș
- Ψ (un singur simbol de copac) = arbori izolați
- ΨΨ (două simboluri de copac) = pâlc
- linie ondulată + Ψ = arbori izolați la pârâu
- linie ondulată + ΨΨ = pâlc izolat de arbori la pârâu
- simbol oiță = supratârlire
Când transcrii "veget_forest" sau "date_compl", păstrează această distincție exactă (nu confunda "arbori izolați" cu "pâlc", și notează dacă simbolul are lângă el linia ondulată de pârâu).

Codurile "S1"-"S4" și "A1"-"A5" din rândul "Date stat. suplimentare" reprezintă grade de eroziune (S=eroziune
de suprafață, A=eroziune de adâncime, 1=cea mai slabă, 4/5=cea mai puternică) — apar des combinate, ex.
"S2 + A2". Citește-le exact ca literă+cifră.

Rândurile "Lucr. exec." și "Lucr. propuse" conțin coduri dintr-o listă fixă de lucrări: numere de trei cifre
începând cu 6 (621-634, 644, 645), sau "46", sau codurile "T1", "T2", "TC" — adesea urmate de un procent din
suprafață (ex. "625 pe 20%"). Citește cifrele/literele exact, nu inventa coduri care nu există în acest interval.

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
  const cfg = getActiveProviderConfig();
  if (!cfg.key){
    const provNames = { claude:'Anthropic', openrouter:'OpenRouter', gemini:'Gemini' };
    toast(`Adaugă mai întâi cheia API ${provNames[cfg.provider]||cfg.provider} din Setări (⚙️).`);
    document.getElementById('panel-settings').style.display = 'block';
    return;
  }
  document.getElementById('processBtn').disabled = true;
  const hint = document.getElementById('processHint');
  let done=0;
  const todo = pendingPhotos.filter(p => p.status !== 'done');
  for (const p of todo){
    p.status='processing'; renderPhotoGrid();
    hint.textContent = `Procesare ${++done}/${todo.length}...`;
    try {
      const data = await callVisionWithRetry(cfg, p.base64, p.mimeType, (msg)=>{ hint.textContent = msg; });
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
    // pauză configurabilă între poze, ca să respectăm limita de cereri/minut a contului
    if (p !== todo[todo.length-1]){
      const pacingSec = parseFloat(localStorage.getItem(LS_PACING) || '7') || 0;
      const remaining = todo.length - done;
      if (pacingSec > 0){
        hint.textContent = `Procesat ${done}/${todo.length} — pauză ${pacingSec}s (mai sunt ~${Math.ceil(remaining*pacingSec/60)} min)...`;
        await sleep(pacingSec * 1000);
      }
    }
  }
  saveFise();
  const okCount = pendingPhotos.filter(p=>p.status==='done').length;
  const errCount = pendingPhotos.filter(p=>p.status==='error').length;
  hint.textContent = errCount
    ? `Gata: ${okCount} reușite, ${errCount} cu eroare (apasă din nou "Extrage" ca să reîncerci doar pe cele eșuate).`
    : `Gata. ${fise.length} fișe în total — vezi tab "Fișe".`;
  document.getElementById('processBtn').disabled = pendingPhotos.length===0;
  renderFiseList();
}

function sleep(ms){ return new Promise(r=>setTimeout(r, ms)); }

// Reîncearcă automat la eroare 429 (quota/rate limit), cu așteptare progresivă
async function callVisionWithRetry(cfg, base64, mimeType, onStatus){
  const maxRetries = 4;
  let lastErr;
  for (let attempt=0; attempt<=maxRetries; attempt++){
    try {
      return cfg.provider === 'claude' ? await callClaude(cfg.key, cfg.model, base64, mimeType)
        : cfg.provider === 'openrouter' ? await callOpenRouter(cfg.key, cfg.model, base64, mimeType)
        : await callGemini(cfg.key, cfg.model, base64, mimeType);
    } catch(err){
      lastErr = err;
      const is429 = /\b429\b/.test(err.message);
      if (!is429 || attempt===maxRetries) throw err;
      const waitSec = Math.min(60, 10 * Math.pow(2, attempt)); // 10s, 20s, 40s, 60s
      for (let s=waitSec; s>0; s--){
        onStatus && onStatus(`Limită de cereri atinsă — reîncerc în ${s}s (încercarea ${attempt+2}/${maxRetries+1})...`);
        await sleep(1000);
      }
    }
  }
  throw lastErr;
}

async function callClaude(key, model, base64, mimeType){
  const url = 'https://api.anthropic.com/v1/messages';
  const body = {
    model,
    max_tokens: 2048,
    messages: [{
      role: 'user',
      content: [
        { type: 'image', source: { type: 'base64', media_type: mimeType, data: base64 } },
        { type: 'text', text: SCHEMA_PROMPT }
      ]
    }]
  };
  const res = await fetch(url, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'x-api-key': key,
      'anthropic-version': '2023-06-01',
      'anthropic-dangerous-direct-browser-access': 'true'
    },
    body: JSON.stringify(body)
  });
  if (!res.ok){
    const txt = await res.text();
    throw new Error(`Claude ${res.status}: ${txt.slice(0,200)}`);
  }
  const json = await res.json();
  let text = json?.content?.[0]?.text;
  if (!text) throw new Error('Răspuns gol de la Claude.');
  text = text.trim().replace(/^```json\s*/i, '').replace(/^```\s*/,'').replace(/```\s*$/,'');
  return JSON.parse(text);
}

async function callOpenRouter(key, model, base64, mimeType){
  const url = 'https://openrouter.ai/api/v1/chat/completions';
  const body = {
    model,
    messages: [{
      role: 'user',
      content: [
        { type: 'text', text: SCHEMA_PROMPT },
        { type: 'image_url', image_url: { url: `data:${mimeType};base64,${base64}` } }
      ]
    }],
    temperature: 0.1,
    max_tokens: 4096
  };
  const res = await fetch(url, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': `Bearer ${key}`,
      'HTTP-Referer': location.origin,
      'X-Title': 'Fise Pasuni OCR'
    },
    body: JSON.stringify(body)
  });
  const raw = await res.text();
  let json;
  try { json = JSON.parse(raw); } catch(e){ throw new Error(`OpenRouter ${res.status}: răspuns nevalid — ${raw.slice(0,200)}`); }

  if (!res.ok || json.error){
    const msg = json?.error?.message || raw.slice(0,200);
    throw new Error(`OpenRouter ${res.status}: ${msg}`);
  }
  const choice = json?.choices?.[0];
  let text = choice?.message?.content;
  if (!text){
    // unele modele :free pun tot răspunsul pe "reasoning" și lasă content gol, sau se opresc din lipsă de tokeni
    text = choice?.message?.reasoning || '';
    const reason = choice?.finish_reason || choice?.native_finish_reason || 'necunoscut';
    if (!text){
      throw new Error(`Răspuns gol de la OpenRouter (motiv oprire: ${reason}). Modelul "${model}" poate să nu proceseze bine imaginea — încearcă alt model din listă.`);
    }
  }
  text = text.trim().replace(/^```json\s*/i, '').replace(/^```\s*/,'').replace(/```\s*$/,'');
  // Unele modele gratuite nu respectă strict "doar JSON" -- extrage primul obiect {...} din răspuns
  const match = text.match(/\{[\s\S]*\}/);
  return JSON.parse(match ? match[0] : text);
}

async function callGemini(key, model, base64, mimeType){
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent`;
  const isGemini3 = /gemini-3/.test(model);
  const generationConfig = {
    responseMimeType: "application/json",
    thinkingConfig: isGemini3 ? { thinkingLevel: "low" } : { thinkingBudget: 0 }
  };
  if (!isGemini3) generationConfig.temperature = 0.1; // sampling params not recommended for Gemini 3.x
  const body = {
    contents: [{
      parts: [
        { inline_data: { mime_type: mimeType, data: base64 } },
        { text: SCHEMA_PROMPT }
      ]
    }],
    generationConfig
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

function addManualFisa(){
  const fisa = normalizeFisa({});
  fisa.id = 'f'+Date.now()+Math.random().toString(36).slice(2,7);
  fisa.thumb = null; // fără poză asociată
  fise.push(fisa);
  saveFise();
  openFisaId = fisa.id;
  switchTab('review');
  renderFiseList();
  toast('Fișă nouă adăugată — completeaz-o mai jos.');
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
function lucrareName(cod){
  if (!cod) return null;
  return CODEBOOK.lucrari[String(cod).trim().toUpperCase()] || null;
}

// Parsează un string liber "625 pe 20%; T1 pe 10%; 632" -> {cod: pct|''} pentru codurile recunoscute
function parseLucrString(text){
  const result = {};
  if (!text) return result;
  for (const cod of Object.keys(CODEBOOK.lucrari)){
    const re = new RegExp('\\b' + cod.replace(/[.*+?^${}()|[\]\\]/g,'\\$&') + '\\b(?:\\s*pe\\s*(\\d+(?:[.,]\\d+)?%))?', 'i');
    const m = text.match(re);
    if (m) result[cod] = m[1] || '';
  }
  return result;
}

function rebuildLucrString(container){
  const parts = [];
  Object.keys(CODEBOOK.lucrari).forEach(cod=>{
    const cb = container.querySelector(`input[data-lucr-cod="${cod}"]`);
    if (cb && cb.checked){
      const pctInp = container.querySelector(`input[data-lucr-pct="${cod}"]`);
      const pct = pctInp ? pctInp.value.trim() : '';
      parts.push(pct ? `${cod} pe ${pct}` : cod);
    }
  });
  return parts.join('; ');
}

function renderLucrChecklist(host, f){
  if (!host) return;
  const checked = parseLucrString(f.lucr_propuse);
  host.innerHTML = Object.entries(CODEBOOK.lucrari).map(([cod, denumire])=>{
    const isChecked = Object.prototype.hasOwnProperty.call(checked, cod);
    const pctVal = checked[cod] || '';
    return `<div class="checklist">
      <input type="checkbox" data-lucr-cod="${cod}" ${isChecked?'checked':''} style="flex:none;width:18px;height:18px;">
      <span style="flex:1;"><b>${cod}</b> — ${esc(denumire)}</span>
      <input type="text" data-lucr-pct="${cod}" value="${esc(pctVal)}" placeholder="%" style="width:52px;flex:none;padding:4px 6px;">
    </div>`;
  }).join('');
  host.querySelectorAll('input[data-lucr-cod], input[data-lucr-pct]').forEach(inp=>{
    inp.addEventListener('input', ()=>{
      f.lucr_propuse = rebuildLucrString(host);
      saveFise();
    });
  });
}
function eroziuneName(cod){
  if (!cod) return null;
  return CODEBOOK.eroziune[String(cod).trim().toUpperCase()] || null;
}
// Auto-adnotează codurile de eroziune (S1-S4/A1-A5) și de lucrări (621-645, T1/T2/TC, 46)
// găsite într-un text liber, adăugând denumirea între paranteze după fiecare cod recunoscut.
function annotateCodes(text){
  if (!text) return text;
  return text.replace(/\b(S[1-4]|A[1-5]|T[C12]|6\d{2}|46)\b(?!\s*%)/g, (m)=>{
    const nm = eroziuneName(m) || lucrareName(m);
    return nm ? `${m} (${nm})` : m;
  });
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
    <p class="hint code-hint" data-hintfor="date_stat_suplim">${esc(annotateCodes(f.date_stat_suplim))}</p>
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
    <p class="hint code-hint" data-hintfor="lucr_exec">${esc(annotateCodes(f.lucr_exec))}</p>
    <label>Lucr. propuse</label>
    <div id="lucrChecklist-${f.id}"></div>
    <div class="btnbar">
      <button class="btn small secondary dup-btn">📋 Copiază text (fișă)</button>
      <button class="btn small danger del-btn">🗑 Șterge fișă</button>
    </div>
  `;
  // bind simple field inputs
  container.querySelectorAll('input[data-field], textarea[data-field]').forEach(inp=>{
    inp.addEventListener('input', ()=>{
      f[inp.dataset.field] = inp.value;
      const hintEl = container.querySelector(`[data-hintfor="${inp.dataset.field}"]`);
      if (hintEl) hintEl.textContent = annotateCodes(inp.value);
      saveFise();
    });
  });
  // bind species rows
  ['gram','leg','div_pl','daun'].forEach(key=> bindSpeciesGroup(container, f, key, speciesName));
  bindSpeciesGroup(container, f, 'arbusti', codLookupShrub);
  renderLucrChecklist(container.querySelector(`#lucrChecklist-${f.id}`), f);
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
  if (f.date_stat_suplim) lines.push(`Date stat suplim: ${annotateCodes(f.date_stat_suplim)}`);
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
  if (f.lucr_exec) lines.push(`Lucr. exec.: ${annotateCodes(f.lucr_exec)}`);
  if (f.lucr_propuse) lines.push(`Lucr. propuse: ${annotateCodes(f.lucr_propuse)}`);
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
