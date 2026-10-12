/* =========================================================
   Données
   ========================================================= */
const PRESETS_TITRE = ['Déjeuner','Dîner','Séminaire et déjeuner','Cocktail','Anniversaire','Mariage','Repas de fin d\'année'];

const PRESETS_PRESTA = [
  {titre:"Café d'accueil", base:'TTC', tva:'10', contenu:"Thé, café, jus de fruits artisanaux\nViennoiseries"},
  {titre:'Apéritif', base:'TTC', tva:'20', contenu:''},
  {titre:'Salle de séminaire', base:'TTC', tva:'20', contenu:'Une table de réunion pour 10 personnes'},
  {titre:'Menu', base:'HT', tva:'10', contenu:"Mise en bouche\n\nEntrée\n\nPlat\n\nDessert\n\nCafé"},
  {titre:'Menu enfant', base:'TTC', tva:'10', contenu:''},
  {titre:'Prestation', base:'TTC', tva:'10', contenu:''}
];

const PRESETS_BLOC = [
  {titre:'Allergies / régimes', col:'G'},
  {titre:'Déroulé horaire', col:'G'},
  {titre:'Matériel', col:'D'},
  {titre:'Déco / plan de table', col:'D'},
  {titre:'Prestataires', col:'D'},
  {titre:'Exclusivité', col:'G'},
  {titre:'Infos diverses', col:'G'},
  {titre:'Bloc libre', col:'D'}
];

const EMPTY = () => ({
  date:'', pax:'', titre:'', heure:'', lieu:'',
  civ:'MME', nom:'', prenom:'', type:'SOCIETE', societe:'', tel:'', email:'',
  acDemande:'', acMontant:'', acVerse:'',
  prestas:[], blocs:[]
});

let S = load() || EMPTY();
let uid = Date.now();

function load(){ try{ const j = localStorage.getItem('fichator.draft'); return j ? Object.assign(EMPTY(), JSON.parse(j)) : null; }catch(e){ return null; } }
function save(){ try{ localStorage.setItem('fichator.draft', JSON.stringify(S)); }catch(e){} }

/* =========================================================
   Formatage
   ========================================================= */
const MOIS = ['janvier','février','mars','avril','mai','juin','juillet','août','septembre','octobre','novembre','décembre'];
const JOURS = ['dimanche','lundi','mardi','mercredi','jeudi','vendredi','samedi'];
function parseD(s){ if(!s) return null; const [y,m,d] = s.split('-').map(Number); return new Date(y, m-1, d); }
function dateLongue(s){ const d = parseD(s); return d ? `Le ${JOURS[d.getDay()]} ${d.getDate()} ${MOIS[d.getMonth()]} ${d.getFullYear()}` : 'Le …'; }
function dateCourte(s){ const d = parseD(s); if(!d) return ''; const p = n => String(n).padStart(2,'0'); return `${p(d.getDate())} ${p(d.getMonth()+1)} ${d.getFullYear()}`; }
function num(v){ if(v===''||v==null) return NaN; return parseFloat(String(v).replace(/\s/g,'').replace(',','.')); }
function euro(n){ if(isNaN(n)) return ''; const r = Math.round(n*100)/100; return (Number.isInteger(r) ? String(r) : r.toFixed(2).replace('.',',')) + '€'; }

function prixLigne(p){
  const a = num(p.prix);
  if(isNaN(a)) return '';
  const autre = p.base === 'TTC' ? 'HT' : 'TTC';
  let b;
  if(p.tva === 'manuel') b = num(p.prix2);
  else { const t = parseFloat(p.tva)/100; b = p.base === 'TTC' ? a/(1+t) : a*(1+t); }
  return isNaN(b) ? ` à ${euro(a)} ${p.base}` : ` à ${euro(a)} ${p.base} soit ${euro(b)} ${autre}`;
}

function nomFichier(){
  const parts = [dateCourte(S.date), S.civ, (S.nom||'').trim().toUpperCase(), (S.societe||'').trim()];
  return parts.filter(Boolean).join(' ').replace(/[\\/:*?"<>|]+/g,'-').replace(/\s+/g,' ').trim() + '.docx';
}
function manques(){
  const m = [];
  if(!S.date) m.push('date'); if(!S.nom) m.push('nom');
  if(!S.societe) m.push(S.type==='SOCIETE' ? 'société' : 'événement');
  return m;
}

/* =========================================================
   Modèle du document (sert à l'aperçu ET au Word)
   ========================================================= */
const TNR = 'Times New Roman', ARIAL = 'Arial';
const C = { rouge:'C8102E', titre:'3F8A8C', presta:'D11A44', alerte:'C0143C', gris:'595959' };
const GRID = [3100, 3269, 3269]; // largeurs des colonnes en twips (≈ 17 cm)

const R = (text, o={}) => Object.assign({text, font:TNR, size:12}, o);
const P = (runs, o={}) => Object.assign({runs: Array.isArray(runs) ? runs : [runs], align:'left'}, o);
const BL = (o={}) => P([], o);

// "texte *en rouge* texte" -> runs
function richRuns(line, base){
  const out = []; const parts = line.split('*');
  parts.forEach((t,i) => { if(!t) return; out.push(i % 2 ? Object.assign({}, base, {text:t, bold:true, color:C.presta}) : Object.assign({}, base, {text:t})); });
  return out;
}
function lignesCentrees(txt){
  return (txt||'').replace(/\r/g,'').split('\n').map(l => P(richRuns(l, R('', {font:ARIAL})), {align:'center'}));
}

function modele(){
  const contact = [S.civ, (S.nom||'').toUpperCase(), S.prenom, S.tel, S.email].filter(Boolean).join('  ');
  const versé = S.acVerse ? dateCourte(S.acVerse) : 'non versé';

  // Colonne gauche (remarques)
  const gauche = [
    P([R('REMARQUES',{underline:true}), R(' :')]),
    P(R('Fiche récapitulative non contractuelle', {size:8})),
    BL(),
    P(R(`Heure d'arrivée : ${S.heure||''}`)),
  ];
  if(S.lieu) gauche.push(P(R(`Lieu : ${S.lieu}`)));
  S.blocs.filter(b => b.col==='G').forEach(b => {
    gauche.push(BL());
    gauche.push(P(R(`${b.titre} :`, {underline:true, bold:true})));
    (b.contenu||'').replace(/\r/g,'').split('\n').forEach(l => gauche.push(P(richRuns(l, R('')))));
  });
  gauche.push(BL(), BL(),
    P([R('FACTURE',{underline:true}), R(' :')]),
    BL(), BL(),
    P(R("Le nombre définitif de convives est à confirmer 72h à l'avance (nombre facturé en cas d'absent)", {italic:true, color:C.alerte}), {align:'both'}),
    BL(), BL(),
    P(R(`Acompte demandé le: ${dateCourte(S.acDemande)}`), {align:'both'}),
    P(R(`Montant: ${S.acMontant ? euro(num(S.acMontant)) || S.acMontant : ''}`)),
    P(R(`Versé le: ${versé}`)),
    BL(), BL(), BL(),
    P(R("En raison de l'inflation, les prix annoncés sont garantis pour une durée de 3 mois à partir de ce jour.", {font:ARIAL, color:C.gris}), {align:'both'}),
    BL()
  );

  // Colonne droite (menu)
  const droite = [BL()];
  S.prestas.forEach((p,i) => {
    if(i) droite.push(BL());
    droite.push(P(R(`${p.titre||''}${prixLigne(p)}`, {font:ARIAL, bold:true, color:C.presta}), {align:'center'}));
    droite.push(...lignesCentrees(p.contenu));
  });
  S.blocs.filter(b => b.col==='D').forEach(b => {
    droite.push(BL());
    droite.push(P(R(b.titre||'', {font:ARIAL, bold:true, color:C.presta}), {align:'center'}));
    droite.push(...lignesCentrees(b.contenu));
  });
  droite.push(BL());

  return [
    { cells:[
      {span:1, rowspan:2, shade:'D9D9D9', valign:'center', paras:[
        P(R('FICHE MISE EN PLACE',{bold:true, size:11}),{align:'center', after:240}),
        P(R('ET',{bold:true, size:11}),{align:'center', after:240}),
        P(R('INFORMATIONS',{bold:true, size:11}),{align:'center', after:240}),
        P(R('BANQUETS',{bold:true, size:11}),{align:'center'})]},
      {span:2, height:900, paras:[P(R(dateLongue(S.date),{bold:true, size:14}),{align:'center'})]}
    ]},
    { cells:[
      {span:2, height:900, valign:'center', paras:[P(R('CHPG',{bold:true, size:20, color:C.rouge}),{align:'center'})]}
    ]},
    { cells:[
      {span:1, rowspan:2, paras:[BL(), BL(), P([R('NOMBRE DE PERSONNES',{underline:true}), R(' :')]), BL(), P(R(S.pax ? `${S.pax} pax` : '', {bold:true}))]},
      {span:2, height:900, valign:'center', paras:[P(R(S.titre||'', {bold:true, size:22, color:C.titre}),{align:'center'})]}
    ]},
    { cells:[
      {span:2, paras:[
        P([R(`${S.type==='SOCIETE'?'SOCIETE':'EVENEMENT'} : ${(S.societe||'').toUpperCase()}`, {bold:true, underline:true, size:11})]),
        BL(),
        P([R('CONTACT',{bold:true, underline:true, size:11}), R(` : ${contact}`, {size:11})], {after:60})
      ]}
    ]},
    { cells:[
      {span:1, paras:gauche},
      {span:2, paras:droite}
    ]},
    { cells:[
      {span:1, paras:[BL()]},
      {span:1, paras:[BL()]},
      {span:1, paras:[P(R('REPAS',{underline:true}))]}
    ]}
  ];
}

/* =========================================================
   Aperçu HTML
   ========================================================= */
function esc(s){ return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;'); }
function runHTML(r){
  const st = [`font-family:'${r.font}'`, `font-size:${r.size}pt`];
  if(r.bold) st.push('font-weight:700'); if(r.italic) st.push('font-style:italic');
  if(r.underline) st.push('text-decoration:underline'); if(r.color) st.push(`color:#${r.color}`);
  return `<span style="${st.join(';')}">${esc(r.text)}</span>`;
}
function parasHTML(paras){
  return paras.map(p => {
    const al = p.align==='both' ? 'justify' : p.align;
    const sz = (p.runs[0] && p.runs[0].size) || 12;
    return `<p style="text-align:${al};font-size:${sz}pt;line-height:1.15;margin-bottom:${(p.after||0)/20}pt">${p.runs.map(runHTML).join('') || '&nbsp;'}</p>`;
  }).join('');
}
function renderPreview(){
  const rows = modele();
  const total = GRID.reduce((a,b)=>a+b,0);
  let h = '<table><colgroup>' + GRID.map(w => `<col style="width:${(w/total*100).toFixed(2)}%">`).join('') + '</colgroup>';
  rows.forEach(row => {
    h += '<tr>';
    row.cells.forEach(c => {
      const st = [];
      if(c.shade) st.push(`background:#${c.shade}`);
      if(c.valign==='center') st.push('vertical-align:middle');
      if(c.height) st.push(`height:${c.height/20}pt`);
      h += `<td colspan="${c.span}" rowspan="${c.rowspan||1}" style="${st.join(';')}">${parasHTML(c.paras)}</td>`;
    });
    h += '</tr>';
  });
  document.getElementById('sheet').innerHTML = h + '</table>';
  const fn = nomFichier(), m = manques();
  document.getElementById('fname').textContent = fn;
  document.getElementById('fwarn').textContent = m.length ? `Manque : ${m.join(', ')}` : '';
  fitSheet();
}
function fitSheet(){
  const wrap = document.getElementById('sheetWrap'), sh = document.getElementById('sheet');
  const k = Math.min(1, wrap.clientWidth / 794);
  sh.style.transform = `scale(${k})`;
  wrap.style.height = (sh.offsetHeight * k) + 'px';
}
window.addEventListener('resize', fitSheet);

/* =========================================================
   Word (.docx) – XML WordprocessingML + zip maison
   ========================================================= */
function x(s){ return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;'); }
function runXML(r){
  let pr = `<w:rFonts w:ascii="${r.font}" w:hAnsi="${r.font}" w:cs="${r.font}"/>`;
  if(r.bold) pr += '<w:b/><w:bCs/>'; if(r.italic) pr += '<w:i/><w:iCs/>';
  if(r.color) pr += `<w:color w:val="${r.color}"/>`;
  pr += `<w:sz w:val="${r.size*2}"/><w:szCs w:val="${r.size*2}"/>`;
  if(r.underline) pr += '<w:u w:val="single"/>';
  return `<w:r><w:rPr>${pr}</w:rPr><w:t xml:space="preserve">${x(r.text)}</w:t></w:r>`;
}
function paraXML(p){
  const sz = ((p.runs[0] && p.runs[0].size) || 12) * 2;
  return `<w:p><w:pPr><w:spacing w:before="0" w:after="${p.after||0}"/><w:jc w:val="${p.align}"/><w:rPr><w:sz w:val="${sz}"/></w:rPr></w:pPr>${p.runs.map(runXML).join('')}</w:p>`;
}
function cellXML(c, w, cont){
  let pr = `<w:tcW w:w="${w}" w:type="dxa"/>`;
  if(c.span > 1) pr += `<w:gridSpan w:val="${c.span}"/>`;
  if(cont) pr += '<w:vMerge/>'; else if(c.rowspan > 1) pr += '<w:vMerge w:val="restart"/>';
  if(c.shade) pr += `<w:shd w:val="clear" w:color="auto" w:fill="${c.shade}"/>`;
  if(c.valign) pr += `<w:vAlign w:val="${c.valign}"/>`;
  const body = cont ? '<w:p/>' : (c.paras.length ? c.paras.map(paraXML).join('') : '<w:p/>');
  return `<w:tc><w:tcPr>${pr}</w:tcPr>${body}</w:tc>`;
}
function documentXML(){
  const rows = modele();
  const pending = {}; // index colonne -> {left, cell}
  let tr = '';
  rows.forEach(row => {
    let col = 0, i = 0, cells = '', hmax = 0;
    while(col < GRID.length){
      let c, cont = false;
      if(pending[col] && pending[col].left > 0){ c = pending[col].cell; cont = true; pending[col].left--; }
      else { c = row.cells[i++]; if(!c) break; if(c.rowspan > 1) pending[col] = {left:c.rowspan-1, cell:c}; }
      const w = GRID.slice(col, col + c.span).reduce((a,b)=>a+b,0);
      if(!cont && c.height) hmax = Math.max(hmax, c.height);
      cells += cellXML(c, w, cont);
      col += c.span;
    }
    const trPr = hmax ? `<w:trPr><w:trHeight w:val="${hmax}" w:hRule="atLeast"/></w:trPr>` : '';
    tr += `<w:tr>${trPr}${cells}</w:tr>`;
  });
  const b = 'w:val="single" w:sz="6" w:space="0" w:color="000000"';
  const tbl = `<w:tbl><w:tblPr><w:tblW w:w="${GRID.reduce((a,b)=>a+b,0)}" w:type="dxa"/><w:jc w:val="center"/>
<w:tblBorders><w:top ${b}/><w:left ${b}/><w:bottom ${b}/><w:right ${b}/><w:insideH ${b}/><w:insideV ${b}/></w:tblBorders>
<w:tblLayout w:type="fixed"/><w:tblCellMar><w:left w:w="108" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar><w:tblLook w:val="0000"/></w:tblPr>
<w:tblGrid>${GRID.map(g=>`<w:gridCol w:w="${g}"/>`).join('')}</w:tblGrid>${tr}</w:tbl>`;
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<w:body>${tbl}<w:p/><w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134" w:header="567" w:footer="567" w:gutter="0"/></w:sectPr></w:body></w:document>`;
}
const STYLES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman" w:eastAsia="Times New Roman"/><w:sz w:val="24"/><w:szCs w:val="24"/><w:lang w:val="fr-FR"/></w:rPr></w:rPrDefault>
<w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>
<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/><w:uiPriority w:val="99"/><w:semiHidden/><w:tblPr><w:tblInd w:w="0" w:type="dxa"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>
</w:styles>`;
const CT = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>`;
const RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`;
const DOCRELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`;

const CRC = (() => { const t = new Uint32Array(256); for(let n=0;n<256;n++){ let c=n; for(let k=0;k<8;k++) c = c&1 ? 0xEDB88320^(c>>>1) : c>>>1; t[n]=c>>>0; } return t; })();
function crc32(u8){ let c = 0xFFFFFFFF; for(let i=0;i<u8.length;i++) c = CRC[(c^u8[i])&255]^(c>>>8); return (c^0xFFFFFFFF)>>>0; }
function zip(files){ // files: [{name, data:string}] – stockage sans compression
  const enc = new TextEncoder(), chunks = [], central = []; let off = 0;
  const d = new Date(), dt = ((d.getHours()<<11)|(d.getMinutes()<<5)|(d.getSeconds()>>1)) & 0xFFFF, dd = (((d.getFullYear()-1980)<<9)|((d.getMonth()+1)<<5)|d.getDate()) & 0xFFFF;
  files.forEach(f => {
    const name = enc.encode(f.name), data = enc.encode(f.data), crc = crc32(data);
    const h = new DataView(new ArrayBuffer(30));
    h.setUint32(0,0x04034b50,true); h.setUint16(4,20,true); h.setUint16(6,0x0800,true); h.setUint16(8,0,true);
    h.setUint16(10,dt,true); h.setUint16(12,dd,true); h.setUint32(14,crc,true); h.setUint32(18,data.length,true); h.setUint32(22,data.length,true);
    h.setUint16(26,name.length,true); h.setUint16(28,0,true);
    chunks.push(new Uint8Array(h.buffer), name, data);
    const c = new DataView(new ArrayBuffer(46));
    c.setUint32(0,0x02014b50,true); c.setUint16(4,20,true); c.setUint16(6,20,true); c.setUint16(8,0x0800,true); c.setUint16(10,0,true);
    c.setUint16(12,dt,true); c.setUint16(14,dd,true); c.setUint32(16,crc,true); c.setUint32(20,data.length,true); c.setUint32(24,data.length,true);
    c.setUint16(28,name.length,true); c.setUint32(42,off,true);
    central.push(new Uint8Array(c.buffer), name);
    off += 30 + name.length + data.length;
  });
  const csize = central.reduce((a,b)=>a+b.length,0);
  const e = new DataView(new ArrayBuffer(22));
  e.setUint32(0,0x06054b50,true); e.setUint16(8,files.length,true); e.setUint16(10,files.length,true); e.setUint32(12,csize,true); e.setUint32(16,off,true);
  return new Blob([...chunks, ...central, new Uint8Array(e.buffer)], {type:'application/vnd.openxmlformats-officedocument.wordprocessingml.document'});
}
function buildDocx(){
  return zip([
    {name:'[Content_Types].xml', data:CT},
    {name:'_rels/.rels', data:RELS},
    {name:'word/document.xml', data:documentXML()},
    {name:'word/_rels/document.xml.rels', data:DOCRELS},
    {name:'word/styles.xml', data:STYLES}
  ]);
}

const API = window.fichator || null; // présent seulement dans le logiciel installé

async function enregistrer(sousAutreNom){
  const m = manques();
  if(m.length && !confirm(`Il manque : ${m.join(', ')}.\nEnregistrer quand même ?`)) return;
  const blob = buildDocx(), name = nomFichier();

  // Logiciel installé : écriture directe dans le dossier des fiches
  if(API){
    const bytes = new Uint8Array(await blob.arrayBuffer());
    const r = sousAutreNom ? await API.saveAs(name, bytes) : await API.save(name, bytes);
    if(r && r.ok){ toast(`Fiche enregistrée : ${r.name}`, {path:r.path}); refreshFolder(); }
    else if(r && r.error) toast(r.error, {error:true});
    return;
  }

  // Version navigateur (dev / dépannage)
  if(window.showSaveFilePicker){
    try{
      const h = await showSaveFilePicker({
        id:'fichator', suggestedName:name, startIn:'documents',
        types:[{description:'Document Word', accept:{'application/vnd.openxmlformats-officedocument.wordprocessingml.document':['.docx']}}]
      });
      const w = await h.createWritable(); await w.write(blob); await w.close();
      toast(`Fiche enregistrée : ${h.name}`);
      return;
    }catch(e){ if(e && e.name === 'AbortError') return; console.warn(e); }
  }
  const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = name;
  document.body.appendChild(a); a.click(); a.remove(); setTimeout(()=>URL.revokeObjectURL(a.href), 4000);
  toast('Fiche téléchargée — à glisser dans le dossier');
}

async function refreshFolder(){
  if(!API) return;
  const c = await API.getConfig();
  document.getElementById('folderBar').hidden = false;
  const el = document.getElementById('folder');
  el.textContent = c.folder ? c.folder + (c.folderOk ? '' : '  (introuvable — réseau ?)') : 'Pas encore choisi (il sera demandé au premier enregistrement)';
  el.style.color = c.folder && !c.folderOk ? '#a33' : '';
}

/* ---------- Mises à jour & nouveautés ---------- */
function cmpV(a, b){ const x = String(a).split('.').map(Number), y = String(b).split('.').map(Number); for(let i=0;i<3;i++){ if((x[i]||0)!==(y[i]||0)) return (x[i]||0)-(y[i]||0); } return 0; }
function showNews(depuis, titre){
  const list = (window.CHANGELOG||[]).filter(e => !depuis || cmpV(e.version, depuis) > 0);
  if(!list.length) return false;
  document.getElementById('newsTitle').textContent = titre || 'Quoi de neuf ?';
  document.getElementById('newsBody').innerHTML = list.map(e => `<h4>Version ${esc(e.version)}${e.date ? ' · ' + esc(e.date.split('-').reverse().join('/')) : ''}</h4><ul>${e.items.map(i=>`<li>${esc(i)}</li>`).join('')}</ul>`).join('');
  document.getElementById('news').hidden = false;
  return true;
}
function stripHtml(h){ const d = document.createElement('div'); d.innerHTML = h; return (d.textContent || '').trim(); }

async function initDesktop(){
  if(!API){ document.getElementById('ver').textContent = 'version navigateur'; document.getElementById('btnSaveAs').hidden = true; return; }
  const c = await API.getConfig();
  document.getElementById('ver').textContent = 'v' + c.version;
  refreshFolder();
  document.getElementById('btnFolder').onclick = async () => { await API.chooseFolder(); refreshFolder(); };
  // Après une mise à jour : on montre les nouveautés une fois
  if(c.lastSeenVersion && cmpV(c.version, c.lastSeenVersion) > 0) showNews(c.lastSeenVersion, `Fichator mis à jour en v${c.version} 🎉`);
  API.markSeen(c.version);

  const bar = document.getElementById('update'), txt = document.getElementById('updateText');
  let notes = '';
  API.onUpdate(u => {
    bar.hidden = false;
    if(u.state === 'downloading') txt.textContent = `Nouvelle version ${u.version} trouvée, téléchargement…`;
    if(u.state === 'progress') txt.textContent = `Téléchargement de la mise à jour… ${u.percent} %`;
    if(u.state === 'ready'){
      notes = stripHtml(u.notes || '');
      txt.textContent = `La version ${u.version} est prête. Elle s'installera à la fermeture, ou tout de suite :`;
      document.getElementById('updateGo').hidden = false;
      document.getElementById('updateNotes').hidden = !notes;
    }
  });
  document.getElementById('updateGo').onclick = () => API.installUpdate();
  document.getElementById('updateNotes').onclick = () => alert(notes);
}

/* =========================================================
   Formulaire
   ========================================================= */
function toast(t, o={}){
  const el = document.getElementById('toast'), open = document.getElementById('toastOpen'), rev = document.getElementById('toastReveal');
  document.getElementById('toastText').textContent = t;
  el.classList.toggle('err', !!o.error);
  open.hidden = rev.hidden = !(o.path && API);
  if(o.path && API){ open.onclick = () => API.open(o.path); rev.onclick = () => API.reveal(o.path); }
  el.classList.add('show'); clearTimeout(toast.t);
  toast.t = setTimeout(()=>el.classList.remove('show'), o.path || o.error ? 7000 : 2800);
}
function changed(){ save(); renderPreview(); }

function bindSimple(){
  document.querySelectorAll('[data-k]').forEach(el => {
    el.value = S[el.dataset.k] || '';
    el.oninput = () => { S[el.dataset.k] = el.value; changed(); };
  });
  segment('segCiv','civ'); segment('segType','type', () => {
    document.getElementById('lblSoc').textContent = S.type==='SOCIETE' ? 'Société' : 'Événement';
    document.querySelector('[data-k=societe]').placeholder = S.type==='SOCIETE' ? "Nom de l'entreprise" : 'Mariage, anniversaire de Léa…';
  });
}
function segment(id, key, after){
  const box = document.getElementById(id);
  const paint = () => { box.querySelectorAll('button').forEach(b => b.classList.toggle('on', b.dataset.v === S[key])); after && after(); };
  box.querySelectorAll('button').forEach(b => b.onclick = () => { S[key] = b.dataset.v; paint(); changed(); });
  paint();
}

function chips(id, list, label, onClick){
  const box = document.getElementById(id); box.innerHTML = '';
  list.forEach(it => { const b = document.createElement('button'); b.type='button'; b.className='chip'; b.textContent = label(it); b.onclick = () => onClick(it); box.appendChild(b); });
}

function moveIn(arr, i, d){ const j = i + d; if(j < 0 || j >= arr.length) return; [arr[i], arr[j]] = [arr[j], arr[i]]; }

function renderPrestas(){
  const box = document.getElementById('listPresta'); box.innerHTML = '';
  if(!S.prestas.length) box.innerHTML = '<p class="empty">Ajoute une prestation avec les boutons ci-dessus.</p>';
  S.prestas.forEach((p, i) => {
    const d = document.createElement('div'); d.className = 'item';
    d.innerHTML = `
      <div class="head">
        <input type="text" value="${esc(p.titre)}" data-f="titre" placeholder="Titre">
        <button class="btn icon" data-a="up" title="Monter">↑</button>
        <button class="btn icon" data-a="down" title="Descendre">↓</button>
        <button class="btn icon" data-a="del" title="Supprimer">✕</button>
      </div>
      <div class="row">
        <label class="f">Prix (€)<input type="text" value="${esc(p.prix||'')}" data-f="prix" placeholder="vide = pas de prix"></label>
        <label class="f">Saisi en<select data-f="base"><option ${p.base==='TTC'?'selected':''}>TTC</option><option ${p.base==='HT'?'selected':''}>HT</option></select></label>
        <label class="f">TVA<select data-f="tva">
          ${['10','20','5.5'].map(t=>`<option value="${t}" ${p.tva===t?'selected':''}>${t.replace('.',',')} %</option>`).join('')}
          <option value="manuel" ${p.tva==='manuel'?'selected':''}>Mixte (je tape l'autre prix)</option></select></label>
        ${p.tva==='manuel' ? `<label class="f">Prix ${p.base==='TTC'?'HT':'TTC'}<input type="text" value="${esc(p.prix2||'')}" data-f="prix2"></label>` : ''}
      </div>
      <div class="calc" data-calc></div>
      <textarea data-f="contenu" rows="5" placeholder="Une ligne par élément">${esc(p.contenu||'')}</textarea>`;
    const calc = d.querySelector('[data-calc]');
    const upd = () => { calc.textContent = prixLigne(p) ? '→ ' + (p.titre||'') + prixLigne(p) : ''; };
    upd();
    d.querySelectorAll('[data-f]').forEach(el => {
      const ev = el.tagName === 'SELECT' ? 'onchange' : 'oninput';
      el[ev] = () => { p[el.dataset.f] = el.value; if(el.dataset.f === 'tva' || el.dataset.f === 'base'){ renderPrestas(); } else upd(); changed(); };
    });
    d.querySelector('[data-a=up]').onclick = () => { moveIn(S.prestas, i, -1); renderPrestas(); changed(); };
    d.querySelector('[data-a=down]').onclick = () => { moveIn(S.prestas, i, 1); renderPrestas(); changed(); };
    d.querySelector('[data-a=del]').onclick = () => { S.prestas.splice(i,1); renderPrestas(); changed(); };
    box.appendChild(d);
  });
}

function renderBlocs(){
  const box = document.getElementById('listBlocs'); box.innerHTML = '';
  if(!S.blocs.length) box.innerHTML = '<p class="empty">Aucun bloc en plus. Clique sur ce dont ce client a besoin.</p>';
  S.blocs.forEach((b, i) => {
    const d = document.createElement('div'); d.className = 'item';
    d.innerHTML = `
      <div class="head">
        <input type="text" value="${esc(b.titre)}" data-f="titre" placeholder="Titre du bloc">
        <div class="seg"><button type="button" data-col="G" class="${b.col==='G'?'on':''}">Gauche</button><button type="button" data-col="D" class="${b.col==='D'?'on':''}">Droite</button></div>
        <button class="btn icon" data-a="up" title="Monter">↑</button>
        <button class="btn icon" data-a="down" title="Descendre">↓</button>
        <button class="btn icon" data-a="del" title="Supprimer">✕</button>
      </div>
      <textarea data-f="contenu" rows="3" placeholder="Une ligne par info">${esc(b.contenu||'')}</textarea>`;
    d.querySelectorAll('[data-f]').forEach(el => el.oninput = () => { b[el.dataset.f] = el.value; changed(); });
    d.querySelectorAll('[data-col]').forEach(el => el.onclick = () => { b.col = el.dataset.col; renderBlocs(); changed(); });
    d.querySelector('[data-a=up]').onclick = () => { moveIn(S.blocs, i, -1); renderBlocs(); changed(); };
    d.querySelector('[data-a=down]').onclick = () => { moveIn(S.blocs, i, 1); renderBlocs(); changed(); };
    d.querySelector('[data-a=del]').onclick = () => { S.blocs.splice(i,1); renderBlocs(); changed(); };
    box.appendChild(d);
  });
}

function init(){
  bindSimple();
  chips('chipsTitre', PRESETS_TITRE, t => t, t => { S.titre = t; document.querySelector('[data-k=titre]').value = t; changed(); });
  chips('chipsPresta', PRESETS_PRESTA, p => '+ ' + p.titre, p => { S.prestas.push(Object.assign({id:++uid, prix:'', prix2:''}, p)); renderPrestas(); changed(); });
  chips('chipsBlocs', PRESETS_BLOC, b => '+ ' + b.titre, b => { S.blocs.push(Object.assign({id:++uid, contenu:''}, b)); renderBlocs(); changed(); });
  renderPrestas(); renderBlocs(); renderPreview();
  document.getElementById('btnSave').onclick = () => enregistrer(false);
  document.getElementById('btnSaveAs').onclick = () => enregistrer(true);
  document.getElementById('showNews').onclick = e => { e.preventDefault(); showNews(null); };
  document.getElementById('newsClose').onclick = () => { document.getElementById('news').hidden = true; };
  document.addEventListener('keydown', e => { if((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 's'){ e.preventDefault(); enregistrer(false); } });
  initDesktop();
  document.getElementById('btnNew').onclick = () => {
    if(!confirm('Vider le formulaire pour une nouvelle fiche ?')) return;
    S = EMPTY(); save(); bindSimple(); renderPrestas(); renderBlocs(); renderPreview();
  };
}
init();
