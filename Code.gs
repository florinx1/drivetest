/**
 * ICG Drive Test — backend Google Apps Script
 * ------------------------------------------------------------------------
 * RECOMANDAT: creează Sheet-ul din contul Google al firmei, pentru ca
 * e-mailurile către clienți să plece de pe adresa firmei, nu de pe una personală.
 *
 * 1. Google Sheet nou („ICG Drive Test”) → Extensii → Apps Script → lipește acest fișier.
 * 2. Setări proiect (rotița) → Fus orar: (GMT+02:00) Bucharest.
 * 3. Rulează o dată funcția setup() și acceptă permisiunile
 *    (Sheets, Drive, Gmail pentru trimitere, acces extern pentru ANAF, declanșatoare).
 *    setup() creează foile, folderul din Drive, cele 4 vehicule, contul admin
 *    „Florin Loghin” cu PIN 1234 și alerta zilnică de la ora 8.
 * 4. Implementare → Implementare nouă → Aplicație web
 *      Execută ca: Eu · Cine are acces: Oricine
 * 5. Copiază URL-ul /exec în index.html → const API_URL = '...'
 * La modificări ale codului: Implementare → Gestionează implementările → Editează →
 * Versiune nouă (URL-ul rămâne același).
 */

const TZ = 'Europe/Bucharest';
const PHOTO_SLOTS = ['Față', 'Spate', 'Lateral stânga', 'Lateral dreapta', 'Bord / kilometraj'];
const VEH_DOCS = [['itpExp', 'ITP'], ['rcaExp', 'RCA'], ['cascoExp', 'CASCO'], ['rovinietaExp', 'Rovinietă']];
const BLOCKING_DOCS = ['itpExp', 'rcaExp', 'rovinietaExp'];

const COLS = {
  Vehicule: ['id','denumire','marca','model','combustibil','vin','nrInmatriculare','culoare','anFabricatie','valoareEUR','recomandareZile','casco','activ','ordine',
    'kmCurent','revizieUrmatoareKm','itpExp','rcaExp','cascoExp','rovinietaExp'],
  Utilizatori: ['id','nume','pin','telefon','email','rol','activ'],
  Rezervari: ['id','vehiculId','dataStart','dataSfarsit','agent','status','creatLa','nrContract',
    'client_tip','client_denumire','client_cui','client_regcom','client_adresa','client_reprezentant','client_functie','client_cnp','client_ci','client_telefon','client_email','client_domeniu','client_obs',
    'sofer_nume','sofer_cnp','sofer_ci','sofer_permis','sofer_categorii','sofer_permisExp','sofer_telefon','sofer2_nume','sofer2_permis',
    'predare_data','predare_ora','predare_km','predare_nivel','predare_loc','predare_oraRetur','predare_chei','predare_daune','predare_dotari','predare_dotariLipsa','predare_documente',
    'predare_reprezentantICG','predare_semnatarClient','predare_calitateSemnatar','predare_acceptClauze','predare_acordMarketing','predare_informareGDPR',
    'retur_data','retur_ora','retur_km','retur_nivel','retur_daune','retur_dauneDesc','retur_dotari','retur_curatenie','retur_primitDe','retur_interes','retur_intentie','retur_pasUrmator','retur_obs',
    'rezultat','rezultat_unitati','rezultat_obs',
    'drive_folder','contract_pdf','contract_email','pv_pdf','pv_email','reamintire_trimisa','alerta_intarziere','anulatMotiv','anulatDe'],
  Blocari: ['id','vehiculId','dataStart','dataSfarsit','motiv','creatDe','creatLa'],
  Asteptare: ['id','vehiculId','dataStart','dataSfarsit','agent','client','obs','status','creatLa','notificatLa'],
  Fotografii: ['id','rezervareId','etapa','slot','fileId','url','creatLa'],
  Semnaturi: ['rezervareId','tip','imagine','creatLa'],
  Setari: ['cheie','valoare']
};

const DEFAULT_SETARI = {
  firma_denumire: 'INTER CARGO GRUP S.R.L.',
  firma_sediu: 'Calea Moșilor nr. 284, Bl. 22A, Sc. 1, Ap. 22, Sector 2, București',
  firma_punct_lucru: 'Șos. de Centură nr. 5-7, 077025, Bragadiru, Jud. Ilfov',
  firma_cui: 'RO 18039372', firma_regcom: 'J2005017287400', firma_reprezentant: '', firma_functie: 'Administrator',
  firma_telefon: '0736 809 783', firma_email: 'office@fotonromania.com', telefon_asistenta: '',
  gdpr_email: 'office@fotonromania.com', politica_url: 'https://www.fotonromania.com/politica-de-confidentialitate/',
  penalitate_zi: '300', franciza_casco: '', contract_prefix: 'DT', contract_next: '1',
  email_alerte: '', email_cc: '', drive_folder_id: ''
};
const HIDDEN_SETARI = ['drive_folder_id'];

/* ======================= SETUP ======================= */
function setup() {
  const ss = SpreadsheetApp.getActive();
  Object.keys(COLS).forEach(name => {
    const sh = ss.getSheetByName(name) || ss.insertSheet(name);
    const cols = COLS[name];
    const existing = sh.getLastColumn() ? sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].filter(String) : [];
    const merged = existing.concat(cols.filter(c => existing.indexOf(c) < 0));
    if (sh.getMaxColumns() < merged.length) sh.insertColumnsAfter(sh.getMaxColumns(), merged.length - sh.getMaxColumns());
    sh.getRange(1, 1, 1, merged.length).setValues([merged]).setFontWeight('bold').setBackground('#1F3864').setFontColor('#FFFFFF');
    sh.setFrozenRows(1);
    sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).setNumberFormat('@');
  });
  if (rows_('Vehicule').length === 0) {
    [
      {id:'V9',denumire:'Tunland V9',marca:'FOTON',model:'Tunland V9 4x4 Double Cab',combustibil:'Diesel',recomandareZile:'3',casco:'Da',activ:'Da',ordine:'1'},
      {id:'G7',denumire:'Tunland G7',marca:'FOTON',model:'Tunland G7 4x4 Double Cab',combustibil:'Diesel',recomandareZile:'3',casco:'Da',activ:'Da',ordine:'2'},
      {id:'ETOANO',denumire:'eToano',marca:'FOTON',model:'eToano',combustibil:'Electric',recomandareZile:'7',casco:'Da',activ:'Da',ordine:'3'},
      {id:'CAVAN',denumire:'Cavan',marca:'CAVAN',model:'Cavan',combustibil:'Electric',recomandareZile:'7',casco:'Da',activ:'Da',ordine:'4'}
    ].forEach(v => upsert_('Vehicule', v));
  }
  if (rows_('Utilizatori').length === 0) upsert_('Utilizatori', { id: 'U1', nume: 'Florin Loghin', pin: '1234', rol: 'admin', activ: 'Da' });
  const s = setari_();
  Object.keys(DEFAULT_SETARI).forEach(k => { if (!(k in s)) setSetare_(k, DEFAULT_SETARI[k]); });
  if (!setari_().drive_folder_id) setSetare_('drive_folder_id', DriveApp.createFolder('ICG Drive Test – Documente').getId());
  installTriggers();
  SpreadsheetApp.flush();
}

/** Alerta zilnică la ora 8 (rulată automat de setup). */
function installTriggers() {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'dailyJob').forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('dailyJob').timeBased().everyDays(1).atHour(8).inTimezone(TZ).create();
}

/* ======================= WEB APP ======================= */
function doGet() { return json_({ ok: true, msg: 'API ICG Drive Test activ' }); }

function doPost(e) {
  let req;
  try { req = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: 'Cerere invalidă.' }); }
  const READ = ['getUsers', 'login', 'getAll', 'getDocData', 'anaf'];
  const lock = LockService.getScriptLock();
  let locked = false;
  try {
    if (READ.indexOf(req.action) < 0) { lock.waitLock(25000); locked = true; }
    return json_({ ok: true, data: handle_(req.action, req.payload || {}, req.auth) });
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  } finally {
    if (locked) try { lock.releaseLock(); } catch (x) {}
  }
}

const PUBLIC_ACTIONS = ['getUsers', 'login'];
const ADMIN_ACTIONS = ['saveBlocare','deleteBlocare','saveUtilizator','deleteUtilizator','saveVehicul','saveSetari','testAlerte'];

function auth_(a) {
  const u = rows_('Utilizatori').filter(x => x.activ !== 'Nu' && x.nume === (a && a.user))[0];
  if (!u || String(u.pin) !== String(a.pin)) throw new Error('AUTH: Sesiunea a expirat sau PIN-ul a fost schimbat. Autentifică-te din nou.');
  return u;
}

function handle_(action, p, auth) {
  const t = today_();
  let me = null;
  if (PUBLIC_ACTIONS.indexOf(action) < 0) {
    me = auth_(auth);
    if (ADMIN_ACTIONS.indexOf(action) >= 0 && me.rol !== 'admin') throw new Error('Doar administratorii pot face această operațiune.');
  }
  const S = setari_();
  const findR = id => { const r = rows_('Rezervari').filter(x => x.id === id)[0]; if (!r) throw new Error('Programarea nu mai există.'); return r; };
  const V = id => rows_('Vehicule').filter(v => v.id === id)[0];

  switch (action) {
    case 'getUsers': return rows_('Utilizatori').filter(u => u.activ !== 'Nu').map(u => u.nume);
    case 'login': {
      const u = rows_('Utilizatori').filter(x => x.activ !== 'Nu' && x.nume === p.user)[0];
      if (!u || String(u.pin) !== String(p.pin)) { Utilities.sleep(800); throw new Error('Nume sau PIN greșit.'); }
      return { user: u.nume, rol: u.rol };
    }
    case 'getAll': { const o = pub_(); o.me = { nume: me.nume, rol: me.rol }; return o; }
    case 'changePin': {
      if (!/^\d{4,8}$/.test(p.pin || '')) throw new Error('PIN-ul trebuie să aibă 4–8 cifre.');
      upsert_('Utilizatori', { id: me.id, pin: p.pin }); return true;
    }

    case 'createRezervare': {
      if (!p.vehiculId || !p.dataStart || !p.dataSfarsit) throw new Error('Alege vehiculul și perioada.');
      if (p.dataStart < t) throw new Error('Nu se poate programa în trecut.');
      if (p.dataSfarsit < p.dataStart) throw new Error('Data de final este înaintea datei de început.');
      const c = conflicts_(p.vehiculId, p.dataStart, p.dataSfarsit);
      if (c.length) throw new Error('Perioada tocmai a fost ocupată: ' + c[0].label + ' (' + fmt_(c[0].s) + ' – ' + fmt_(c[0].e) + ').');
      const r = { id: uid_('R'), vehiculId: p.vehiculId, dataStart: p.dataStart, dataSfarsit: p.dataSfarsit, agent: me.nume, status: 'PROGRAMAT', creatLa: t };
      if (p.asteptareId) {
        const a = rows_('Asteptare').filter(x => x.id === p.asteptareId)[0];
        if (a) { upsert_('Asteptare', { id: a.id, status: 'INCHISA' }); r.client_denumire = a.client || ''; }
      }
      upsert_('Rezervari', r);
      const out = pub_(); out.newId = r.id; return out;
    }
    case 'saveClient': {
      const r = findR(p.id);
      if (r.status !== 'PROGRAMAT') throw new Error('Datele clientului se pot modifica doar înainte de predare.');
      const upd = { id: r.id };
      Object.keys(p.fields || {}).filter(k => k.indexOf('client_') === 0).forEach(k => upd[k] = p.fields[k]);
      upsert_('Rezervari', upd); return pub_();
    }
    case 'predare': {
      const r = findR(p.id), v = V(r.vehiculId);
      if (r.status !== 'PROGRAMAT') throw new Error('Vehiculul a fost deja predat sau programarea este închisă.');
      if (t < r.dataStart) throw new Error('Predarea se confirmă începând cu ziua programării (' + fmt_(r.dataStart) + ').');
      const bd = blockingDocs_(v, t);
      if (bd.length) throw new Error('Vehiculul nu poate fi predat: ' + bd.join(', ') + ' expirat. Actualizează documentele în Admin → Vehicule.');
      if ((p.fields || {}).predare_informareGDPR !== 'Da') throw new Error('Confirmă că clientul a primit Nota de informare privind prelucrarea datelor.');
      const so = stillOut_(r);
      if (so) throw new Error('Vehiculul nu a fost returnat de clientul anterior (' + (so.client_denumire || '—') + ', agent ' + so.agent + ', termen ' + fmt_(so.dataSfarsit) + '). Predarea se poate face după confirmarea returului.');
      if ((p.fotos || []).filter(f => PHOTO_SLOTS.indexOf(f.slot) >= 0).length < PHOTO_SLOTS.length) throw new Error('Lipsesc pozele obligatorii.');
      const upd = { id: r.id, status: 'PREDAT' };
      Object.keys(p.fields || {}).filter(k => /^(sofer|sofer2|predare)_/.test(k)).forEach(k => upd[k] = p.fields[k]);
      const n = parseInt(S.contract_next || '1', 10);
      upd.nrContract = (S.contract_prefix || 'DT') + '-' + t.slice(0, 4) + '-' + ('00' + n).slice(-3);
      setSetare_('contract_next', String(n + 1));
      const merged = Object.assign({}, r, upd);
      const folder = folderFor_(merged, v);
      upd.drive_folder = folder.getUrl();
      saveFotos_(merged, 'predare', p.fotos, folder, t);
      saveSigs_(r.id, p.semnaturi, t);
      upsert_('Rezervari', upd);
      if (num_(upd.predare_km) > num_(v.kmCurent)) upsert_('Vehicule', { id: v.id, kmCurent: upd.predare_km });
      return pub_();
    }
    case 'prelungire': {
      const r = findR(p.id);
      if (['PROGRAMAT', 'PREDAT'].indexOf(r.status) < 0) throw new Error('Programarea este închisă.');
      if (!p.dataSfarsit || p.dataSfarsit < r.dataStart) throw new Error('Data nouă este înaintea începutului.');
      const c = conflicts_(r.vehiculId, r.dataStart, p.dataSfarsit, r.id);
      if (c.length) throw new Error('Nu se poate prelungi: urmează ' + c[0].label + ' din ' + fmt_(c[0].s) + '.');
      upsert_('Rezervari', { id: r.id, dataSfarsit: p.dataSfarsit, reamintire_trimisa: '' });
      checkWaitlist_(r.vehiculId); return pub_();
    }
    case 'retur': {
      const r = findR(p.id), v = V(r.vehiculId);
      if (r.status !== 'PREDAT') throw new Error('Vehiculul nu figurează ca predat.');
      if ((p.fotos || []).filter(f => PHOTO_SLOTS.indexOf(f.slot) >= 0).length < PHOTO_SLOTS.length) throw new Error('Lipsesc pozele obligatorii.');
      const upd = { id: r.id, status: 'TERMINAT', rezultat: r.rezultat || 'În lucru' };
      Object.keys(p.fields || {}).filter(k => k.indexOf('retur_') === 0).forEach(k => upd[k] = p.fields[k]);
      const folder = folderFor_(r, v);
      saveFotos_(r, 'retur', p.fotos, folder, t);
      saveSigs_(r.id, p.semnaturi, t);
      upsert_('Rezervari', upd);
      if (num_(upd.retur_km) > num_(v.kmCurent)) upsert_('Vehicule', { id: v.id, kmCurent: upd.retur_km });
      checkWaitlist_(r.vehiculId);
      return pub_();
    }
    case 'anulare': {
      const r = findR(p.id);
      if (r.status !== 'PROGRAMAT') throw new Error('Doar programările nepredate pot fi anulate.');
      upsert_('Rezervari', { id: r.id, status: 'ANULAT', anulatMotiv: p.motiv || '', anulatDe: me.nume });
      checkWaitlist_(r.vehiculId); return pub_();
    }
    case 'getDocData': {
      const fotos = rows_('Fotografii').filter(f => f.rezervareId === p.id).map(f => {
        let data = '';
        try { const b = DriveApp.getFileById(f.fileId).getBlob(); data = 'data:' + b.getContentType() + ';base64,' + Utilities.base64Encode(b.getBytes()); } catch (e) {}
        return { etapa: f.etapa, slot: f.slot, data: data };
      });
      return { semnaturi: rows_('Semnaturi').filter(s => s.rezervareId === p.id), fotos: fotos };
    }
    case 'savePdf': {
      const r = findR(p.id), v = V(r.vehiculId);
      const blob = Utilities.newBlob(Utilities.base64Decode(p.b64), 'application/pdf', p.filename || 'document.pdf');
      const folder = folderFor_(r, v);
      folder.getFilesByName(blob.getName()).hasNext() && trashByName_(folder, blob.getName());
      const file = folder.createFile(blob);
      const k = p.kind === 'pv' ? 'pv' : 'contract';
      const upd = { id: r.id }; upd[k + '_pdf'] = file.getUrl(); upd.drive_folder = folder.getUrl();
      let emailed = '';
      if (p.email && r.client_email) {
        sendDocEmail_(r, v, S, me, blob, k);
        emailed = r.client_email; upd[k + '_email'] = emailed;
      }
      upsert_('Rezervari', upd);
      const out = pub_(); out.saved = { url: file.getUrl(), emailed: emailed }; return out;
    }
    case 'saveRezultat': {
      findR(p.id);
      upsert_('Rezervari', { id: p.id, rezultat: p.rezultat || '', rezultat_unitati: p.rezultat_unitati || '', rezultat_obs: p.rezultat_obs || '' });
      return pub_();
    }
    case 'addAsteptare': {
      if (!p.vehiculId || !p.dataStart || !p.dataSfarsit || p.dataSfarsit < p.dataStart) throw new Error('Alege o perioadă validă.');
      if (p.dataSfarsit < t) throw new Error('Perioada este în trecut.');
      if (!conflicts_(p.vehiculId, p.dataStart < t ? t : p.dataStart, p.dataSfarsit).length) throw new Error('Vehiculul este liber în această perioadă. Programează-l direct.');
      upsert_('Asteptare', { id: uid_('A'), vehiculId: p.vehiculId, dataStart: p.dataStart, dataSfarsit: p.dataSfarsit, agent: me.nume, client: p.client || '', obs: p.obs || '', status: 'ACTIVA', creatLa: t });
      return pub_();
    }
    case 'deleteAsteptare': {
      const a = rows_('Asteptare').filter(x => x.id === p.id)[0];
      if (a) { if (a.agent !== me.nume && me.rol !== 'admin') throw new Error('Poți șterge doar cererile tale.'); upsert_('Asteptare', { id: a.id, status: 'INCHISA' }); }
      return pub_();
    }
    case 'anaf': return anaf_(p.cui);

    /* ---- admin ---- */
    case 'saveBlocare': {
      if (!p.dataStart || !p.dataSfarsit || p.dataSfarsit < p.dataStart) throw new Error('Perioadă invalidă.');
      (p.vehiculIds || []).forEach(v => upsert_('Blocari', { id: uid_('B'), vehiculId: v, dataStart: p.dataStart, dataSfarsit: p.dataSfarsit, motiv: p.motiv || '', creatDe: me.nume, creatLa: t }));
      return pub_();
    }
    case 'deleteBlocare': {
      const b = rows_('Blocari').filter(x => x.id === p.id)[0];
      deleteRow_('Blocari', p.id); if (b) checkWaitlist_(b.vehiculId); return pub_();
    }
    case 'saveUtilizator': {
      if (!p.nume) throw new Error('Completează numele.');
      if (p.pin && !/^\d{4,8}$/.test(p.pin)) throw new Error('PIN-ul trebuie să aibă 4–8 cifre.');
      if (p.id) {
        const upd = { id: p.id, activ: 'Da' };
        ['nume', 'telefon', 'email', 'rol'].forEach(k => { if (k in p) upd[k] = p[k]; });
        if (p.pin) upd.pin = p.pin;
        upsert_('Utilizatori', upd);
      } else {
        if (!p.pin) throw new Error('Setează un PIN pentru utilizatorul nou.');
        if (rows_('Utilizatori').some(x => x.activ !== 'Nu' && x.nume === p.nume)) throw new Error('Există deja un utilizator cu acest nume.');
        upsert_('Utilizatori', { id: uid_('U'), nume: p.nume, pin: p.pin, telefon: p.telefon || '', email: p.email || '', rol: p.rol || 'agent', activ: 'Da' });
      }
      return pub_();
    }
    case 'deleteUtilizator': {
      const u = rows_('Utilizatori').filter(x => x.id === p.id)[0];
      if (u && u.nume === me.nume) throw new Error('Nu îți poți dezactiva propriul cont.');
      upsert_('Utilizatori', { id: p.id, activ: 'Nu' }); return pub_();
    }
    case 'saveVehicul': {
      if (!V(p.id)) throw new Error('Vehicul inexistent.');
      upsert_('Vehicule', p); return pub_();
    }
    case 'saveSetari': {
      Object.keys(p).forEach(k => { if (HIDDEN_SETARI.indexOf(k) < 0) setSetare_(k, p[k]); });
      return pub_();
    }
    case 'testAlerte': {
      const r = sendAdminSummary_(true);
      return { sent: r.sent, to: r.to, count: r.count };
    }
  }
  throw new Error('Acțiune necunoscută: ' + action);
}

/* ======================= DISPONIBILITATE (identică cu aplicația) ======================= */
function occIntervals_(vehId, excludeId) {
  const t = today_(), out = [];
  rows_('Blocari').filter(b => b.vehiculId === vehId).forEach(b => out.push({ s: b.dataStart, e: b.dataSfarsit, label: b.motiv || 'Indisponibil' }));
  rows_('Rezervari').filter(r => r.vehiculId === vehId && r.id !== excludeId).forEach(r => {
    const who = (r.client_denumire || 'client necompletat') + ' · ' + r.agent;
    if (r.status === 'PROGRAMAT') out.push({ s: r.dataStart, e: r.dataSfarsit, label: 'Programat: ' + who });
    else if (r.status === 'PREDAT') out.push({ s: r.dataStart, e: r.dataSfarsit > t ? r.dataSfarsit : t, label: 'În test: ' + who });
    else if (r.status === 'TERMINAT') out.push({ s: r.dataStart, e: r.retur_data || r.dataSfarsit, label: 'Finalizat: ' + who });
  });
  return out;
}
function conflicts_(vehId, s, e, excludeId) { return occIntervals_(vehId, excludeId).filter(i => s <= i.e && i.s <= e); }
function stillOut_(r) { return rows_('Rezervari').filter(x => x.vehiculId === r.vehiculId && x.status === 'PREDAT' && x.id !== r.id)[0]; }
function lateBlocker_(r, t) {
  if (r.status !== 'PROGRAMAT' || r.dataStart > addDays_(t, 7)) return null;
  const x = stillOut_(r);
  return x && x.dataSfarsit < t && x.dataStart < r.dataStart ? x : null;
}
function blockingDocs_(v, t) { return BLOCKING_DOCS.filter(k => v[k] && v[k] < t).map(k => VEH_DOCS.filter(x => x[0] === k)[0][1]); }

/* ======================= LISTĂ DE AȘTEPTARE ======================= */
function checkWaitlist_(vehId) {
  const t = today_(), users = rows_('Utilizatori'), veh = rows_('Vehicule');
  rows_('Asteptare').filter(a => a.status === 'ACTIVA' && (!vehId || a.vehiculId === vehId)).forEach(a => {
    if (a.dataSfarsit < t) { upsert_('Asteptare', { id: a.id, status: 'INCHISA' }); return; }
    const s = a.dataStart < t ? t : a.dataStart;
    if (conflicts_(a.vehiculId, s, a.dataSfarsit).length) return;
    upsert_('Asteptare', { id: a.id, status: 'LIBER', notificatLa: t });
    const u = users.filter(x => x.nume === a.agent)[0], v = veh.filter(x => x.id === a.vehiculId)[0] || {};
    if (u && u.email) mail_({
      to: u.email,
      subject: 'Liber: ' + v.denumire + ' ' + fmt_(s) + ' – ' + fmt_(a.dataSfarsit),
      htmlBody: wrap_('Vehiculul cerut s-a eliberat', '<p>' + esc_(v.denumire) + ' este acum liber în perioada <b>' + fmt_(s) + ' – ' + fmt_(a.dataSfarsit) + '</b>, pentru care ai trimis o cerere pe lista de așteptare (client: ' + esc_(a.client || '—') + ').</p><p>Deschide aplicația ICG Drive Test și programează-l din secțiunea „De făcut astăzi” înainte să îl rezerve altcineva.</p>')
    });
  });
}

/* ======================= DRIVE ======================= */
function rootFolder_() {
  const id = setari_().drive_folder_id;
  try { if (id) return DriveApp.getFolderById(id); } catch (e) {}
  const f = DriveApp.createFolder('ICG Drive Test – Documente'); setSetare_('drive_folder_id', f.getId()); return f;
}
function sub_(parent, name) { const it = parent.getFoldersByName(name); return it.hasNext() ? it.next() : parent.createFolder(name); }
function folderFor_(r, v) {
  const client = clean_(r.client_denumire || 'Client necompletat');
  const name = clean_((r.nrContract || r.id) + ' – ' + (v ? v.denumire : r.vehiculId) + ' – ' + fmt_(r.dataStart));
  return sub_(sub_(rootFolder_(), client), name);
}
function saveFotos_(r, etapa, fotos, folder, t) {
  const dir = sub_(folder, etapa === 'retur' ? 'Poze retur' : 'Poze predare');
  (fotos || []).forEach(f => {
    const m = String(f.data || '').match(/^data:(image\/\w+);base64,(.+)$/); if (!m) return;
    const file = dir.createFile(Utilities.newBlob(Utilities.base64Decode(m[2]), m[1], clean_(f.slot) + '.jpg'));
    try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) { /* domeniul poate bloca partajarea prin link */ }
    upsert_('Fotografii', { id: uid_('F'), rezervareId: r.id, etapa: etapa, slot: f.slot, fileId: file.getId(), url: file.getUrl(), creatLa: t });
  });
}
function trashByName_(folder, name) { const it = folder.getFilesByName(name); while (it.hasNext()) it.next().setTrashed(true); }

/* ======================= E-MAIL ======================= */
function mail_(o) {
  const S = setari_();
  const msg = { to: o.to, subject: o.subject, htmlBody: o.htmlBody, name: 'Inter Cargo Grup – Drive Test' };
  if (o.cc) msg.cc = o.cc;
  if (S.firma_email) msg.replyTo = S.firma_email;
  if (o.attachments) msg.attachments = o.attachments;
  MailApp.sendEmail(msg);
}
function wrap_(title, inner) {
  const S = setari_();
  return '<div style="font-family:Arial,sans-serif;max-width:620px;color:#16181B;line-height:1.5">' +
    '<div style="background:#1F3864;color:#fff;padding:14px 18px;font-size:18px;font-weight:bold">' + esc_(title) + '</div>' +
    '<div style="padding:16px 18px;border:1px solid #D5DDE7;border-top:0">' + inner +
    '<p style="margin-top:22px;color:#555;font-size:13px">' + esc_(S.firma_denumire || 'Inter Cargo Grup') + '<br>' + esc_(S.firma_punct_lucru || S.firma_sediu || '') +
    (S.telefon_asistenta ? '<br>Asistență drive test: ' + esc_(S.telefon_asistenta) : '') +
    '<br><span style="font-size:12px">Informații despre prelucrarea datelor personale: ' + esc_(S.politica_url || 'https://www.fotonromania.com/politica-de-confidentialitate/') + ' · cereri: ' + esc_(S.gdpr_email || S.firma_email || '') + '</span></p></div></div>';
}
function sendDocEmail_(r, v, S, me, blob, kind) {
  const cc = [S.email_cc, me.email].filter(String).join(',');
  const isPv = kind === 'pv';
  const subject = (isPv ? 'Proces-verbal de restituire ' : 'Contract de comodat drive test ') + (r.nrContract || '') + ' – ' + v.marca + ' ' + v.denumire;
  const body = isPv
    ? '<p>Bună ziua,</p><p>Vă mulțumim că ați testat ' + esc_(v.marca + ' ' + v.denumire) + '. Atașat găsiți procesul-verbal de restituire a vehiculului din data de ' + fmt_(r.retur_data) + ', împreună cu fotografiile de la predare și restituire.</p><p>Pentru orice întrebare sau pentru o ofertă, ne puteți contacta oricând.</p><p>Cu stimă,<br>' + esc_(me.nume) + '</p>'
    : '<p>Bună ziua,</p><p>Atașat găsiți contractul de comodat nr. <b>' + esc_(r.nrContract) + '</b> pentru testarea vehiculului <b>' + esc_(v.marca + ' ' + v.denumire) + '</b>' + (v.nrInmatriculare ? ' (' + esc_(v.nrInmatriculare) + ')' : '') + ', împreună cu procesul-verbal de predare, fotografiile vehiculului și Nota de informare privind prelucrarea datelor personale (Anexa 3).</p><p>Perioada de test: <b>' + fmt_(r.dataStart) + ' – ' + fmt_(r.dataSfarsit) + '</b>' + (r.predare_oraRetur ? ', retur până la ora ' + esc_(r.predare_oraRetur) : '') + '.</p><p>În caz de avarie sau accident, vă rugăm să ne anunțați în cel mult 2 ore' + (S.telefon_asistenta ? ', la ' + esc_(S.telefon_asistenta) : '') + '.</p><p>Vă dorim un test plăcut!<br>' + esc_(me.nume) + '</p>';
  mail_({ to: r.client_email, cc: cc, subject: subject, htmlBody: wrap_(isPv ? 'Proces-verbal de restituire' : 'Contract de comodat – drive test', body), attachments: [blob] });
}

/* ======================= ALERTE ZILNICE ======================= */
function dailyJob() {
  const t = today_(), tomorrow = addDays_(t, 1), veh = rows_('Vehicule');
  checkWaitlist_(null);
  rows_('Rezervari').filter(r => r.status === 'PREDAT' && r.dataSfarsit === tomorrow && r.client_email && r.reamintire_trimisa !== tomorrow).forEach(r => {
    const v = veh.filter(x => x.id === r.vehiculId)[0] || {};
    try {
      mail_({ to: r.client_email, subject: 'Reamintire: returnarea vehiculului ' + v.denumire + ' mâine, ' + fmt_(r.dataSfarsit),
        htmlBody: wrap_('Reamintire retur drive test', '<p>Bună ziua,</p><p>Vă reamintim că perioada de drive test pentru <b>' + esc_(v.marca + ' ' + v.denumire) + '</b> (' + esc_(v.nrInmatriculare || '') + ') se încheie mâine, <b>' + fmt_(r.dataSfarsit) + (r.predare_oraRetur ? ', ora ' + esc_(r.predare_oraRetur) : '') + '</b>, la ' + esc_(r.predare_loc || 'sediul nostru') + '.</p><p>Vă rugăm să returnați vehiculul cu nivelul de ' + (/electric/i.test(v.combustibil || '') ? 'încărcare a bateriei' : 'combustibil') + ' de la predare. Dacă doriți prelungirea testului, ne puteți contacta astăzi.</p><p>Vă mulțumim!<br>' + esc_(r.agent) + '</p>') });
      upsert_('Rezervari', { id: r.id, reamintire_trimisa: tomorrow });
    } catch (e) {}
  });
  notifyLateImpact_();
  sendAdminSummary_(false);
}
/** Anunță agentul care are următoarea programare pe un vehicul întârziat (o singură dată pe programare). */
function notifyLateImpact_() {
  const t = today_(), users = rows_('Utilizatori'), veh = rows_('Vehicule');
  rows_('Rezervari').filter(r => !r.alerta_intarziere).forEach(r => {
    const lb = lateBlocker_(r, t); if (!lb) return;
    const u = users.filter(x => x.nume === r.agent)[0], late = users.filter(x => x.nume === lb.agent)[0], v = veh.filter(x => x.id === r.vehiculId)[0] || {};
    if (u && u.email) {
      try {
        mail_({ to: u.email, cc: late && late.email && late.email !== u.email ? late.email : '',
          subject: 'Atenție: ' + v.denumire + ' întârzie – programarea ta din ' + fmt_(r.dataStart),
          htmlBody: wrap_('Vehicul returnat cu întârziere', '<p>' + esc_(v.denumire) + ' nu a fost încă returnat de <b>' + esc_(lb.client_denumire || 'clientul anterior') + '</b> (termen ' + fmt_(lb.dataSfarsit) + ', agent ' + esc_(lb.agent) + ').</p><p>Programarea ta pentru <b>' + esc_(r.client_denumire || 'client necompletat') + '</b> începe pe <b>' + fmt_(r.dataStart) + '</b>. Anunță clientul din timp sau mută programarea. Predarea rămâne blocată în aplicație până la confirmarea returului.</p>') });
      } catch (e) {}
    }
    upsert_('Rezervari', { id: r.id, alerta_intarziere: t });
  });
}
function buildAlerts_() {
  const t = today_(), veh = rows_('Vehicule'), V = id => veh.filter(v => v.id === id)[0] || { denumire: id }, out = [];
  rows_('Rezervari').forEach(r => {
    const v = V(r.vehiculId).denumire, c = r.client_denumire || 'client necompletat';
    if (r.status === 'PREDAT' && r.dataSfarsit < t) out.push({ level: 'bad', text: 'Retur întârziat ' + zile_(diff_(r.dataSfarsit, t)) + ': ' + v + ' – ' + c + ' (termen ' + fmt_(r.dataSfarsit) + ', agent ' + r.agent + ')' });
    else if (r.status === 'PREDAT' && r.dataSfarsit === t) out.push({ level: 'info', text: 'Retur azi: ' + v + ' – ' + c + ' (agent ' + r.agent + ')' });
    if (r.status === 'PROGRAMAT' && r.dataStart <= t && r.dataSfarsit >= t) out.push({ level: 'info', text: 'Predare azi: ' + v + ' – ' + c + ' (agent ' + r.agent + ')' });
    if (r.status === 'PROGRAMAT' && !r.client_denumire && r.dataStart >= t && diff_(t, r.dataStart) <= 2) out.push({ level: 'warn', text: 'Programare fără date client: ' + v + ', ' + fmt_(r.dataStart) + ' (agent ' + r.agent + ')' });
    const lb = lateBlocker_(r, t);
    if (lb) out.push({ level: 'bad', text: 'Programare afectată de întârziere: ' + v + ' – ' + c + ', din ' + fmt_(r.dataStart) + ' (agent ' + r.agent + '). Vehiculul nu a fost returnat de ' + (lb.client_denumire || 'clientul anterior') + ' (termen ' + fmt_(lb.dataSfarsit) + ', agent ' + lb.agent + ').' });
    if (r.status === 'PROGRAMAT' && r.dataSfarsit < t) out.push({ level: 'warn', text: 'Programare nepredată și expirată: ' + v + ' – ' + c + ', ' + fmt_(r.dataStart) + '–' + fmt_(r.dataSfarsit) });
  });
  veh.filter(v => v.activ !== 'Nu').forEach(v => {
    VEH_DOCS.forEach(d => { if (!v[d[0]]) return; const n = diff_(t, v[d[0]]);
      if (n < 0) out.push({ level: 'bad', text: v.denumire + ': ' + d[1] + ' expirat din ' + fmt_(v[d[0]]) });
      else if (n <= 30) out.push({ level: 'warn', text: v.denumire + ': ' + d[1] + ' expiră în ' + zile_(n) + ' (' + fmt_(v[d[0]]) + ')' }); });
    if (num_(v.kmCurent) && num_(v.revizieUrmatoareKm)) { const rem = num_(v.revizieUrmatoareKm) - num_(v.kmCurent);
      if (rem <= 0) out.push({ level: 'bad', text: v.denumire + ': revizie depășită cu ' + (-rem) + ' km' });
      else if (rem <= 1000) out.push({ level: 'warn', text: v.denumire + ': revizie în ' + rem + ' km' }); }
  });
  rows_('Asteptare').filter(a => a.status === 'LIBER').forEach(a => out.push({ level: 'info', text: 'Listă de așteptare: ' + V(a.vehiculId).denumire + ' liber ' + fmt_(a.dataStart) + '–' + fmt_(a.dataSfarsit) + ' pentru ' + (a.client || 'client') + ' (agent ' + a.agent + ')' }));
  const ord = { bad: 0, warn: 1, info: 2 }; return out.sort((a, b) => ord[a.level] - ord[b.level]);
}
function sendAdminSummary_(force) {
  const S = setari_(), to = String(S.email_alerte || '').trim(), al = buildAlerts_();
  if (!to || (!al.length && !force)) return { sent: false, to: to, count: al.length };
  if (!al.length) return { sent: false, to: to, count: 0 };
  const color = { bad: '#BD3526', warn: '#A96400', info: '#4A5A6E' }, label = { bad: 'URGENT', warn: 'ATENȚIE', info: 'INFO' };
  const html = '<table style="border-collapse:collapse;width:100%">' + al.map(a => '<tr><td style="padding:6px 8px;border-bottom:1px solid #eee;color:' + color[a.level] + ';font-weight:bold;font-size:12px;white-space:nowrap">' + label[a.level] + '</td><td style="padding:6px 8px;border-bottom:1px solid #eee">' + esc_(a.text) + '</td></tr>').join('') + '</table>';
  mail_({ to: to, subject: 'Drive test – situația zilei ' + fmt_(today_()) + ' (' + al.filter(a => a.level === 'bad').length + ' urgente)', htmlBody: wrap_('Situația flotei demo – ' + fmt_(today_()), html) });
  return { sent: true, to: to, count: al.length };
}

/* ======================= ANAF ======================= */
function anaf_(cui) {
  const c = String(cui || '').replace(/\D/g, '');
  if (!c) throw new Error('Introdu un CUI valid.');
  const res = UrlFetchApp.fetch('https://webservicesp.anaf.ro/api/PlatitorTvaRest/v9/tva', {
    method: 'post', contentType: 'application/json', muteHttpExceptions: true,
    payload: JSON.stringify([{ cui: Number(c), data: today_() }])
  });
  if (res.getResponseCode() !== 200) throw new Error('Serviciul ANAF nu răspunde acum (cod ' + res.getResponseCode() + '). Completează manual sau reîncearcă.');
  const j = JSON.parse(res.getContentText());
  const f = (j.found || [])[0];
  if (!f) throw new Error('CUI-ul ' + c + ' nu a fost găsit în baza ANAF.');
  const g = f.date_generale || {}, tva = !!(f.inregistrare_scop_Tva && f.inregistrare_scop_Tva.scpTVA);
  return { cui: (tva ? 'RO' : '') + c, denumire: g.denumire || '', adresa: g.adresa || '', nrRegCom: g.nrRegCom || '', telefon: g.telefon || '', tva: tva };
}

/* ======================= HELPERS ======================= */
function pub_() {
  const s = setari_(); HIDDEN_SETARI.forEach(k => delete s[k]);
  return {
    vehicule: rows_('Vehicule'),
    utilizatori: rows_('Utilizatori').map(u => { const o = Object.assign({}, u); delete o.pin; return o; }),
    rezervari: rows_('Rezervari'), blocari: rows_('Blocari'), asteptare: rows_('Asteptare'),
    fotografii: rows_('Fotografii').map(f => ({ id: f.id, rezervareId: f.rezervareId, etapa: f.etapa, slot: f.slot, url: f.url, thumb: 'https://drive.google.com/thumbnail?id=' + f.fileId + '&sz=w480' })),
    setari: s
  };
}
const CACHE_ = {};
function sheet_(name) { const sh = SpreadsheetApp.getActive().getSheetByName(name); if (!sh) throw new Error('Lipsește foaia ' + name + '. Rulează setup().'); return sh; }
function norm_(v) { if (v instanceof Date) return Utilities.formatDate(v, TZ, 'yyyy-MM-dd'); return v === null || v === undefined ? '' : String(v); }
function rows_(name) {
  if (CACHE_[name]) return CACHE_[name].map(o => Object.assign({}, o));
  const sh = sheet_(name), last = sh.getLastRow();
  let out = [];
  if (last >= 2) {
    const vals = sh.getRange(1, 1, last, sh.getLastColumn()).getValues(), h = vals[0];
    out = vals.slice(1).filter(r => r.some(String)).map(r => { const o = {}; h.forEach((k, i) => { if (k) o[k] = norm_(r[i]); }); return o; });
  }
  CACHE_[name] = out;
  return out.map(o => Object.assign({}, o));
}
function headers_(sh, keys) {
  let h = sh.getRange(1, 1, 1, Math.max(1, sh.getLastColumn())).getValues()[0].filter(String);
  const missing = keys.filter(k => h.indexOf(k) < 0);
  if (missing.length) {
    if (sh.getMaxColumns() < h.length + missing.length) sh.insertColumnsAfter(sh.getMaxColumns(), h.length + missing.length - sh.getMaxColumns());
    sh.getRange(1, h.length + 1, 1, missing.length).setValues([missing]).setFontWeight('bold');
    h = h.concat(missing);
  }
  return h;
}
function upsert_(name, obj) {
  delete CACHE_[name];
  const sh = sheet_(name), key = name === 'Setari' ? 'cheie' : 'id';
  const h = headers_(sh, Object.keys(obj));
  const last = sh.getLastRow(), ki = h.indexOf(key);
  let rowIdx = -1;
  if (last >= 2) {
    const ids = sh.getRange(2, ki + 1, last - 1, 1).getValues();
    for (let i = 0; i < ids.length; i++) if (String(ids[i][0]) === String(obj[key])) { rowIdx = i + 2; break; }
  }
  if (rowIdx < 0) {
    sh.getRange(last + 1, 1, 1, h.length).setNumberFormat('@').setValues([h.map(k => (k in obj ? String(obj[k]) : ''))]);
  } else {
    const range = sh.getRange(rowIdx, 1, 1, h.length), cur = range.getValues()[0];
    range.setNumberFormat('@').setValues([h.map((k, i) => (k in obj ? String(obj[k]) : norm_(cur[i])))]);
  }
}
function deleteRow_(name, id) {
  delete CACHE_[name];
  const sh = sheet_(name), last = sh.getLastRow(); if (last < 2) return;
  const ids = sh.getRange(2, 1, last - 1, 1).getValues();
  for (let i = ids.length - 1; i >= 0; i--) if (String(ids[i][0]) === String(id)) sh.deleteRow(i + 2);
}
function setari_() { const o = {}; rows_('Setari').forEach(r => o[r.cheie] = r.valoare); return o; }
function setSetare_(k, v) { upsert_('Setari', { cheie: k, valoare: v }); }
function saveSigs_(rezId, sigs, t) {
  const sh = sheet_('Semnaturi'), h = headers_(sh, COLS.Semnaturi);
  Object.keys(sigs || {}).forEach(tip => {
    const img = sigs[tip]; if (!img) return;
    if (img.length > 49000) throw new Error('Semnătura este prea mare. Șterge-o și semnează din nou, mai simplu.');
    const o = { rezervareId: rezId, tip: tip, imagine: img, creatLa: t };
    sh.getRange(sh.getLastRow() + 1, 1, 1, h.length).setNumberFormat('@').setValues([h.map(k => (k in o ? o[k] : ''))]);
  });
}
function today_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'); }
function addDays_(s, n) { const p = s.split('-').map(Number); const d = new Date(Date.UTC(p[0], p[1] - 1, p[2] + n)); return Utilities.formatDate(d, 'UTC', 'yyyy-MM-dd'); }
function diff_(a, b) { const pa = a.split('-').map(Number), pb = b.split('-').map(Number); return Math.round((Date.UTC(pb[0], pb[1] - 1, pb[2]) - Date.UTC(pa[0], pa[1] - 1, pa[2])) / 86400000); }
function zile_(n) { return n === 1 ? '1 zi' : n + ' zile'; }
function fmt_(s) { return s ? String(s).split('-').reverse().join('.') : ''; }
function num_(v) { const n = parseFloat(String(v || '').replace(',', '.')); return isNaN(n) ? 0 : n; }
function clean_(s) { return String(s).replace(/[\\/:*?"<>|#]+/g, ' ').trim().slice(0, 90); }
function esc_(s) { return String(s == null ? '' : s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c])); }
function uid_(p) { return p + Date.now().toString(36).toUpperCase() + Math.random().toString(36).slice(2, 5).toUpperCase(); }
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
