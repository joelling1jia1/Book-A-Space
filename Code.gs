/** Bound script: open the Church Space Booking sheet → Extensions → Apps Script → paste this.
 *  Deploy → New deployment → Web app → Execute as: Me, Access: Anyone. */
const GID = 1228728446;                 // the booking tab
const TZ = 'Asia/Kuala_Lumpur';
// app key -> EXACT text written to column D (Space)
const SPACES = {main:'Main Hall', toddler:'Toddler Room', hangout:'Hangout Place Living Hall',
                study:'Study Room', multi:'Multipurpose', kids:'Hope Kids'};
const FALLBACK_LGS = ['Others'];
const MON = ['jan','feb','mar','apr','may','jun','jul','aug','sep','oct','nov','dec'];

function sh_() { return SpreadsheetApp.getActive().getSheets().find(s => s.getSheetId() === GID); }
function hdr_(sh) {
  const a = sh.getRange(1,1,15,1).getDisplayValues();
  for (let i=0;i<a.length;i++) if (String(a[i][0]).trim().toLowerCase()==='no.') return i+1;
  return 3;
}
function key_(name) {                    // sheet text -> app key (tolerates old "Living Room")
  const n = String(name).trim().toLowerCase();
  if (n.startsWith('hangout')) return 'hangout';
  for (const k in SPACES) if (SPACES[k].toLowerCase()===n) return k;
  return null;
}
function pDate_(s) {
  s = String(s).trim(); let m;
  if ((m = s.match(/^(\d{1,2})\s*[-\/ ]\s*([A-Za-z]+)\s*[-\/ ]\s*(\d{2,4})$/))) {
    const mo = MON.indexOf(m[2].slice(0,3).toLowerCase()); if (mo<0) return null;
    return fmt_(+m[3], mo+1, +m[1]);
  }
  if ((m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/))) return fmt_(+m[3], +m[2], +m[1]);
  return null;
}
function fmt_(y,mo,d){ if (y<100) y+=2000; if (y<2000||y>2100) return null;
  return y+'-'+('0'+mo).slice(-2)+'-'+('0'+d).slice(-2); }
function pTime_(s) {
  const m = String(s).trim().replace(/\s+/g,'').match(/^(\d{1,2})(?:[:.](\d{2}))?(?::\d{2})?([AaPp][Mm])?$/);
  if (!m) return null; let h=+m[1]; const mi=+(m[2]||0), ap=(m[3]||'').toLowerCase();
  if (ap==='pm'&&h<12) h+=12; if (ap==='am'&&h===12) h=0;
  return h>23||mi>59 ? null : h*60+mi;
}
function lgs_(sh, h) {
  const rule = sh.getRange(h+1,3).getDataValidation();
  if (rule) {
    const t = rule.getCriteriaType(), v = rule.getCriteriaValues();
    if (t===SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) return v[0];
    if (t===SpreadsheetApp.DataValidationCriteria.VALUE_IN_RANGE)
      return v[0].getValues().flat().filter(String);
  }
  const col = sh.getRange(h+1,3,sh.getLastRow()-h,1).getDisplayValues().flat().map(x=>x.trim()).filter(String);
  return [...new Set(col)].length ? [...new Set(col)] : FALLBACK_LGS;
}
function load_() {
  const sh = sh_(), h = hdr_(sh), n = sh.getMaxRows()-h;
  const rows = sh.getRange(h+1,1,n,9).getDisplayValues();
  const today = Utilities.formatDate(new Date(),TZ,'yyyy-MM-dd'), out = [];
  rows.forEach((r,i) => {
    const date = pDate_(r[4]), space = key_(r[3]), s = pTime_(r[5]);
    if (!date || !space || s===null || date < today) return;
    let e = pTime_(r[6]); if (e===null || e<=s) e = s+120;
    out.push({id:h+1+i, name:r[1].trim(), lg:r[2].trim(), space, date, s, e, desc:r[7].trim(), phone:r[8].trim()});
  });
  out.sort((a,b)=> a.date<b.date?-1:a.date>b.date?1:a.s-b.s);
  return {sh,h,rows,bookings:out};
}
function json_(o){ return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

function doGet() {
  const d = load_();
  return json_({bookings:d.bookings, lgs:lgs_(d.sh,d.h)});
}
function doPost(e) {
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const b = JSON.parse(e.postData.contents);
    return json_(b.action==='cancel' ? cancel_(b) : create_(b));
  } catch (err) { return json_({ok:false,error:String(err)}); }
  finally { lock.releaseLock(); }
}
function create_(b) {
  const d = load_(), s = pTime_(b.start), e = pTime_(b.end);
  if (!b.name||!b.lg||!b.desc||s===null||e===null||e<=s||!b.dates.length||!b.spaces.length) return {ok:false,error:'Missing or invalid details'};
  if (b.spaces.length>3||b.dates.length>12) return {ok:false,error:'Too many spaces or dates'};
  const todo = [], skipped = [];
  b.dates.forEach(date => b.spaces.forEach(k => {
    if (!SPACES[k]) return;
    const hit = d.bookings.find(x=>x.space===k&&x.date===date&&x.s<e&&s<x.e);
    hit ? skipped.push({date,space:SPACES[k]}) : todo.push({date,k});
  }));
  if (skipped.length && !b.skipClashes) return {ok:false,error:'Clash detected',skipped};
  if (!todo.length) return {ok:false,error:'All selected slots are already booked',skipped};
  // free rows = Name, Space and Date all empty (fills the pre-numbered blank rows first)
  const free = []; d.rows.forEach((r,i)=>{ if(!r[1].trim()&&!r[3].trim()&&!r[4].trim()) free.push(d.h+1+i); });
  let maxNo = 0; d.rows.forEach(r=>{ const n=parseInt(r[0],10); if(n>maxNo) maxNo=n; });
  todo.forEach((t,i) => {
    const row = free[i] || (d.sh.insertRowAfter(d.sh.getMaxRows()), d.sh.getMaxRows());
    const no = d.rows[row-d.h-1] && parseInt(d.rows[row-d.h-1][0],10) || ++maxNo;
    const p = t.date.split('-');
    const cells = d.sh.getRange(row,1,1,9);
    d.sh.getRange(row,9).setNumberFormat('@');
    d.sh.getRange(row,5).setNumberFormat('d-MMM-yyyy');
    d.sh.getRange(row,6,1,2).setNumberFormat('h:mm AM/PM');
    cells.setValues([[no,b.name.trim(),b.lg,SPACES[t.k],new Date(+p[0],+p[1]-1,+p[2]),b.start,b.end,b.desc.trim(),String(b.phone||'')]]);
  });
  SpreadsheetApp.flush();
  return {ok:true,created:todo.length,skipped};
}
function cancel_(b) {
  const sh = sh_(), r = sh.getRange(b.id,1,1,9).getDisplayValues()[0];
  if (key_(r[3])!==b.space || pDate_(r[4])!==b.date || pTime_(r[5])!==b.start)
    return {ok:false,error:'This booking has changed. Please refresh.'};
  sh.getRange(b.id,2,1,8).clearContent();   // keeps No. (A) and the clash formula (J)
  return {ok:true};
}
