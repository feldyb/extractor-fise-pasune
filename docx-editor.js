/* ====================================================================
   docx-editor.js
   Lucrează direct pe șablonul Word real al utilizatorului (nu generează
   un document nou). Deschide fișierul .docx (e un zip), găsește tabelul
   fișei după numărul de u.a., completează celulele goale păstrând exact
   formatarea (font Arial, mărime, bold/italic) din șablon, apoi
   reîmpachetează documentul pentru descărcare.
   ==================================================================== */

const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

const DocxEditor = (function(){

  let zip = null;
  let xmlDoc = null;
  let xmlString = null;

  async function load(arrayBuffer){
    zip = await JSZip.loadAsync(arrayBuffer);
    xmlString = await zip.file('word/document.xml').async('string');
    xmlDoc = new DOMParser().parseFromString(xmlString, 'application/xml');
    const err = xmlDoc.getElementsByTagName('parsererror');
    if (err.length) throw new Error('Șablonul Word nu a putut fi citit (XML invalid).');
    return true;
  }

  function isLoaded(){ return !!xmlDoc; }

  function getTables(){
    return Array.from(xmlDoc.getElementsByTagNameNS(W_NS, 'tbl'));
  }

  function cellText(tc){
    const ts = tc.getElementsByTagNameNS(W_NS, 't');
    let s = '';
    for (const t of ts) s += t.textContent;
    return s.trim();
  }

  // Extrage valorile existente din celula "Încl: ... Exp: ... Alt: ... Unit.sol: ..."
  // (funcționează atât pe un rând complet gol -- toate câmpurile ies '' -- cât și pe unul
  // parțial completat de o extragere anterioară).
  function parseInclExpAltSol(text){
    const inclM = text.match(/Încl\s*:?\s*([\s\S]*?)\s*Exp\s*:/i);
    const expM = text.match(/Exp\s*:?\s*([\s\S]*?)\s*Alt\s*:/i);
    const altM = text.match(/Alt\s*:?\s*([\s\S]*?)\s*Unit\.?\s*\.?\s*sol\s*:/i);
    const solM = text.match(/Unit\.?\s*\.?\s*sol\s*:?\s*([\s\S]*)$/i);
    return {
      incl: inclM ? inclM[1].trim() : '',
      exp: expM ? expM[1].trim() : '',
      alt: altM ? altM[1].trim() : '',
      unitsol: solM ? solM[1].trim() : ''
    };
  }

  function rowCells(tr){
    return Array.from(tr.children).filter(el => el.localName === 'tc');
  }

  // Identify "fișă" tables: 14 rows, 8-cell header value row (row idx1) with a u.a. value in cell[1]
  function listFisaTables(){
    const tables = getTables();
    const out = [];
    tables.forEach((tbl, tblIdx) => {
      const trs = Array.from(tbl.getElementsByTagNameNS(W_NS, 'tr'));
      if (trs.length !== 14) return;
      const row1cells = rowCells(trs[1]);
      if (row1cells.length < 8) return;
      const uaText = cellText(row1cells[1]);
      const gramText = cellText(rowCells(trs[5])[0]);
      out.push({
        tblIdx, tbl, trs,
        u_a: uaText,
        filled: gramText.length > 8  // has more than just "Gram:" label
      });
    });
    return out;
  }

  function findTableByUA(ua){
    const tables = listFisaTables();
    const target = String(ua).trim().toUpperCase();
    return tables.find(t => t.u_a.trim().toUpperCase() === target) || null;
  }

  /* ---------- run / text helpers ---------- */

  function makeRun(text, {bold=false, italic=false}={}){
    const r = xmlDoc.createElementNS(W_NS, 'w:r');
    const rPr = xmlDoc.createElementNS(W_NS, 'w:rPr');
    const fonts = xmlDoc.createElementNS(W_NS, 'w:rFonts');
    fonts.setAttributeNS(W_NS, 'w:ascii', 'Arial');
    fonts.setAttributeNS(W_NS, 'w:hAnsi', 'Arial');
    fonts.setAttributeNS(W_NS, 'w:cs', 'Arial');
    rPr.appendChild(fonts);
    if (bold) rPr.appendChild(xmlDoc.createElementNS(W_NS, 'w:b'));
    if (italic){
      rPr.appendChild(xmlDoc.createElementNS(W_NS, 'w:i'));
      rPr.appendChild(xmlDoc.createElementNS(W_NS, 'w:iCs'));
    }
    const sz = xmlDoc.createElementNS(W_NS, 'w:sz'); sz.setAttributeNS(W_NS,'w:val','16');
    const szCs = xmlDoc.createElementNS(W_NS, 'w:szCs'); szCs.setAttributeNS(W_NS,'w:val','16');
    rPr.appendChild(sz); rPr.appendChild(szCs);
    r.appendChild(rPr);
    const t = xmlDoc.createElementNS(W_NS, 'w:t');
    t.setAttributeNS('http://www.w3.org/XML/1998/namespace', 'xml:space', 'preserve');
    t.textContent = text;
    r.appendChild(t);
    return r;
  }

  // Clear all runs in the FIRST paragraph of a cell, replace with given parts
  function setCellParts(tc, parts){
    const p = tc.getElementsByTagNameNS(W_NS, 'p')[0];
    if (!p) return;
    Array.from(p.getElementsByTagNameNS(W_NS, 'r')).forEach(r => r.remove());
    parts.forEach(([text, opts]) => p.appendChild(makeRun(text, opts||{})));
  }

  // Append runs at the end of the first paragraph of a cell (keeps existing content)
  function appendCellParts(tc, parts){
    const p = tc.getElementsByTagNameNS(W_NS, 'p')[0];
    if (!p) return;
    parts.forEach(([text, opts]) => p.appendChild(makeRun(text, opts||{})));
  }

  function isCellEmpty(tc){
    // Blank template cells contain exactly one run (just the bold label).
    // Filled cells have multiple runs (label run + one or more data/species runs).
    const runs = tc.getElementsByTagNameNS(W_NS, 'r');
    return runs.length <= 1;
  }

  /* ---------- species text builder ---------- */

  function speciesParts(total, list, lookupFn){
    const parts = [[ (total||'') + ' (', {}]];
    const items = (list||[]).filter(s=>s.cod);
    items.forEach((s,i)=>{
      const nm = lookupFn(s.cod) || `cod ${s.cod}`;
      parts.push([nm, {italic:true}]);
      parts.push([` ${s.pct||''}` + (i<items.length-1 ? '; ' : ''), {}]);
    });
    parts.push([')', {}]);
    return parts;
  }

  /* ---------- main fill function ---------- */

  function fillFisa(fisa, lookups){
    const found = findTableByUA(fisa.u_a);
    if (!found) return { ok:false, reason:'no_match' };
    const trs = found.trs;

    // header value row (row idx 1): Tr.Păş, u.a, Supr, Gr.funct, Categ.folos, Unit.relief, Config.teren, T.S.
    const hcells = rowCells(trs[1]);
    // Real column order: Tr.Păş | u.a. | Supr.(ha) | Gr.funcț | T.S. | Categ.folos. | Unit.relief(text) | Config.teren(text)
    const CATEG_FOLOS = {P:'Pășune', PCA:'Pășune cu arbori', PCF:'Pășune cu fânaț', F:'Fâneață'};
    const reliefTxt = lookups.relief ? (lookups.relief(fisa.unit_relief) || fisa.unit_relief) : fisa.unit_relief;
    const configTxt = lookups.config ? (lookups.config(fisa.config_teren) || fisa.config_teren) : fisa.config_teren;
    const categTxt = CATEG_FOLOS[String(fisa.categ_folos||'').trim().toUpperCase()] || fisa.categ_folos;
    const headerVals = [fisa.tr_pas, fisa.u_a, fisa.supr_ha, fisa.gr_funct, fisa.t_s,
                         categTxt, reliefTxt, configTxt];
    headerVals.forEach((v,i)=>{
      if (v && cellText(hcells[i]).length === 0) setCellParts(hcells[i], [[v, {bold:true}]]);
    });

    const rowCell = (idx)=> rowCells(trs[idx])[0];

    // row2: Încl/Exp/Alt/Unit.sol (real template shows the SOIL NAME, not the raw code)
    // Completare "inteligentă": dacă rândul are deja ceva scris (ex. Încl/Exp de la o extragere
    // anterioară), nu-l mai sărim complet -- extragem ce e deja acolo și completăm doar ce lipsește
    // (ex. doar Altitudinea, adăugată manual ulterior).
    {
      const existing = parseInclExpAltSol(cellText(rowCell(2)));
      const solName = lookups.soil(fisa.unit_sol);
      const finalIncl = existing.incl || fisa.incl || '';
      const finalExp = existing.exp || fisa.exp || '';
      const finalAlt = existing.alt || fisa.alt || '';
      const finalSol = existing.unitsol || solName || fisa.unit_sol || '';
      // rescriem doar dacă avem ceva nou de adăugat față de ce era deja acolo
      const changed = finalIncl !== existing.incl || finalExp !== existing.exp ||
                       finalAlt !== existing.alt || finalSol !== existing.unitsol;
      if (changed || isCellEmpty(rowCell(2))){
        setCellParts(rowCell(2), [
          ['Încl:     ', {bold:true}], [finalIncl, {}],
          ['            Exp:      ', {bold:true}], [finalExp, {}],
          ['          Alt:       ', {bold:true}], [finalAlt, {}],
          ['                                  Unit.sol:  ', {bold:true}],
          [finalSol, {}]
        ]);
      }
    }
    // row3: Date stat suplim
    if (fisa.date_stat_suplim && isCellEmpty(rowCell(3))){
      setCellParts(rowCell(3), [['Date stat suplim: ', {bold:true}], [fisa.date_stat_suplim, {}]]);
    }
    // row4: Tip pajiste / Acoperire ierbacee
    if (isCellEmpty(rowCell(4))){
      setCellParts(rowCell(4), [
        ['Tip pajişte:                  ', {bold:true}], [fisa.tip_pajiste||'', {}],
        ['                                                          Acoperire ierbacee:   ', {bold:true}],
        [fisa.acoperire_ierbacee||'', {}]
      ]);
    }
    // row5 Gram
    if (isCellEmpty(rowCell(5)) && fisa.gram && fisa.gram.length){
      setCellParts(rowCell(5), [['Gram:   ', {bold:true}], ...speciesParts(fisa.gram_total, fisa.gram, lookups.plant)]);
    }
    // row6 Leg
    if (isCellEmpty(rowCell(6)) && fisa.leg && fisa.leg.length){
      setCellParts(rowCell(6), [['Leg:  ', {bold:true}], ...speciesParts(fisa.leg_total, fisa.leg, lookups.plant)]);
    }
    // row7 Div pl
    if (isCellEmpty(rowCell(7)) && fisa.div_pl && fisa.div_pl.length){
      setCellParts(rowCell(7), [['Div. pl:  ', {bold:true}], ...speciesParts(fisa.div_pl_total, fisa.div_pl, lookups.plant)]);
    }
    // row8 daunatoare
    if (isCellEmpty(rowCell(8)) && fisa.daun && fisa.daun.length){
      setCellParts(rowCell(8), [['Pl. dăunătoare + toxice:  ', {bold:true}], ...speciesParts(fisa.daun_total, fisa.daun, lookups.plant)]);
    }
    // row9 Val.past / Arbusti / Gr.acop / Rasp
    if (isCellEmpty(rowCell(9))){
      const arbustiParts = [];
      const arb = (fisa.arbusti||[]).filter(s=>s.cod);
      arb.forEach((s,i)=>{
        const nm = lookups.shrub(s.cod);
        arbustiParts.push([nm ? nm.sci : `cod ${s.cod}`, {italic:true}]);
        arbustiParts.push([(i<arb.length-1?', ':''), {}]);
      });
      setCellParts(rowCell(9), [
        ['Val. past :  ', {bold:true}], [fisa.val_past||'', {}],
        ['        Arbuşti.:  ', {bold:true}], ...arbustiParts,
        ['          Gr. acop.:        ', {bold:true}], [fisa.gr_acop||'', {}],
        ['                Răsp .: ', {bold:true}], [fisa.rasp||'', {}]
      ]);
    }
    // row10 veget forest
    if (fisa.veget_forest && isCellEmpty(rowCell(10))){
      setCellParts(rowCell(10), [
        ['Veget. forest .: ', {bold:true}], [fisa.veget_forest, {}],
        ['  Vârsta.:       Consist:       Răsp:          Volum:', {bold:true}]
      ]);
    }
    // row11 date compl
    if (fisa.date_compl && isCellEmpty(rowCell(11))){
      setCellParts(rowCell(11), [['Date compl:   ', {bold:true}], [fisa.date_compl, {}]]);
    }
    // row12 lucr exec
    if (fisa.lucr_exec && isCellEmpty(rowCell(12))){
      setCellParts(rowCell(12), [['Lucr.exec.: ', {bold:true}], [fisa.lucr_exec, {}]]);
    }
    // row13 lucr propuse
    if (fisa.lucr_propuse && isCellEmpty(rowCell(13))){
      setCellParts(rowCell(13), [['Lucr. prop.:   ', {bold:true}], [fisa.lucr_propuse, {}]]);
    }

    return { ok:true };
  }

  async function exportBlob(){
    const serializer = new XMLSerializer();
    const newXml = serializer.serializeToString(xmlDoc);
    zip.file('word/document.xml', newXml);
    return await zip.generateAsync({
      type: 'blob',
      compression: 'DEFLATE',
      compressionOptions: { level: 6 },
      mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    });
  }

  return { load, isLoaded, listFisaTables, findTableByUA, fillFisa, exportBlob };
})();
