//                                                                
//  REJOES Dashboard - Google Sheets
//  Adauga la sfarsitul Code.gs
//                                                                

const DASH_SHEET = 'DASHBOARD';
const DASH_COMPANY = 'S Colect REC S.R.L.';
const DASH_ADDR = 'Str. Republicii nr. 20, 305400 Jimbolia, Jud. Timis | CIF: RO 38513453';
const DASH_EMAIL = 'info@tm-res.com';

const DC = {
  dark:  '#0a1e14', green: '#1a3828', gold:  '#c9a84c',
  white: '#ffffff', gray:  '#f5f5f5', red:   '#c0392b',
  orange:'#e67e22', blue:  '#2980b9', text:  '#2c3e50',
  muted: '#7f8c8d', border:'#d4c5a0'
};

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Rejoes Dashboard')
    .addItem('Regenereaza Dashboard (luna curenta)', 'setupDashboard')
    .addItem('Raport perioada personalizata...', 'buildCustomPeriod')
    .addSeparator()
    .addItem('Trimite Raport Full prin Email...', 'sendFullReport')
    .addToUi();
}

function setupDashboard() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dates = getDefaultDates();
  buildDash(ss, dates);
  SpreadsheetApp.getUi().alert('Dashboard creat cu succes!');
}

function getDefaultDates() {
  var now = new Date();
  return { start: new Date(now.getFullYear(), now.getMonth(), 1), end: now };
}

function buildDash(ss, dates) {
  var sh = ss.getSheetByName(DASH_SHEET);
  if (sh) ss.deleteSheet(sh);
  sh = ss.insertSheet(DASH_SHEET, 0);

  var ncols = 10;
  var colW = [20, 220, 110, 100, 100, 110, 110, 120, 120, 20];
  colW.forEach(function(w, i) { sh.setColumnWidth(i+1, w); });

  var row = 1;
  var dateStr = dFmt(dates.start) + ' - ' + dFmt(dates.end);

  // Header
  row = addHeader(sh, row, ncols, dateStr);

  // Get data
  var pons = getRange(ss, 'Pontaj', dates);
  var raps = getRange(ss, 'Raportare', dates);

  // KPIs
  row = addTitle(sh, row, 'INDICATORI CHEIE', ncols);
  row = addKPIs(sh, row, pons, raps);

  // Pontaj
  row = addTitle(sh, row, 'PONTAJ - CLASAMENT ORE LUCRATE', ncols);
  row = addPontaj(sh, row, pons, dates);

  // Vanzari
  row = addTitle(sh, row, 'VANZARI - CLASAMENT MAGAZINE', ncols);
  row = addVanzari(sh, row, raps);

  // Diferente
  row = addTitle(sh, row, 'DIFERENTE CASA', ncols);
  row = addDiferente(sh, row, raps);

  // Footer
  addDashFooter(sh, row, ncols, dateStr);

  sh.setFrozenRows(4);
  SpreadsheetApp.flush();
}

function addHeader(sh, row, ncols, dateStr) {
  sh.setRowHeight(row, 6);
  sh.getRange(row,1,1,ncols).merge().setBackground(DC.dark); row++;
  sh.setRowHeight(row, 56);
  sh.getRange(row,2,1,4).merge()
    .setValue(DASH_COMPANY)
    .setFontFamily('Georgia').setFontSize(20).setFontWeight('bold')
    .setFontColor(DC.gold).setBackground(DC.dark).setVerticalAlignment('middle');
  sh.getRange(row,6,1,4).merge()
    .setValue('RAPORT OPERATIONAL\n' + dateStr)
    .setFontFamily('Georgia').setFontSize(10).setFontWeight('bold')
    .setFontColor(DC.border).setBackground(DC.dark)
    .setVerticalAlignment('middle').setHorizontalAlignment('right').setWrap(true);
  sh.getRange(row,1).setBackground(DC.dark);
  sh.getRange(row,10).setBackground(DC.dark);
  row++;
  sh.setRowHeight(row, 26);
  sh.getRange(row,1,1,ncols).merge()
    .setValue('   ' + DASH_ADDR + '  |  ' + DASH_EMAIL + '  |  Generat: ' + dFmtFull(new Date()))
    .setBackground(DC.green).setFontColor(DC.border).setFontSize(9).setVerticalAlignment('middle');
  row++;
  sh.setRowHeight(row, 12);
  sh.getRange(row,1,1,ncols).merge().setBackground(DC.dark); row++;
  return row;
}

function addTitle(sh, row, title, ncols) {
  sh.setRowHeight(row, 6);
  sh.getRange(row,1,1,ncols).merge().setBackground(DC.gray); row++;
  sh.setRowHeight(row, 32);
  sh.getRange(row,1,1,ncols).merge()
    .setValue('  ' + title)
    .setBackground(DC.green).setFontColor(DC.gold)
    .setFontFamily('Georgia').setFontSize(12).setFontWeight('bold')
    .setVerticalAlignment('middle');
  row++;
  sh.setRowHeight(row, 3);
  sh.getRange(row,1,1,ncols).merge().setBackground(DC.gold); row++;
  return row;
}

function addKPIs(sh, row, pons, raps) {
  var ore = pons.reduce(function(s,r){return s+parseH(r[6]);},0);
  var ang = new Set(pons.map(function(r){return r[1];})).size;
  var vanz = raps.reduce(function(s,r){return s+(parseFloat(r[3])||0);},0);
  var dif  = raps.reduce(function(s,r){return s+(parseFloat(r[7])||0);},0);
  var bon  = raps.reduce(function(s,r){return s+(parseInt(r[8])||0);},0);

  sh.setRowHeight(row, 14); row++;
  var kpis = [
    {l:'Angajati activi', v:ang, c:DC.blue},
    {l:'Total ore lucrate', v:ore.toFixed(0)+' h', c:DC.green},
    {l:'Vanzari totale', v:fmtR(vanz), c:'#1a6b3a'},
    {l:'Bonuri fiscale', v:bon, c:DC.blue},
    {l:'Diferenta casa', v:fmtR(dif), c:Math.abs(dif)<1?'#1a6b3a':Math.abs(dif)<50?DC.orange:DC.red}
  ];

  kpis.forEach(function(k, i) {
    var col = 2 + i*2 - 1;
    sh.setRowHeight(row, 48);
    sh.getRange(row, col).setValue(k.l)
      .setFontSize(9).setFontColor(DC.muted).setFontWeight('bold')
      .setBackground('#fafafa').setVerticalAlignment('bottom')
      .setBorder(true,true,false,true,false,false,'#e0d5b5',SpreadsheetApp.BorderStyle.SOLID);
    sh.setRowHeight(row+1, 40);
    sh.getRange(row+1, col).setValue(k.v)
      .setFontFamily('Georgia').setFontSize(18).setFontWeight('bold')
      .setFontColor(k.c).setBackground('#fafafa').setVerticalAlignment('top')
      .setBorder(false,true,true,true,false,false,'#e0d5b5',SpreadsheetApp.BorderStyle.SOLID);
  });
  row += 2;
  sh.setRowHeight(row, 14); row++;
  return row;
}


// Employee work schedule norms
function getEmpNorm(empName) {
  var name = empName.toUpperCase();
  // Concediu maternitate
  if (name.indexOf('BOTAS') >= 0) return {type: 'maternitate', hPerDay: 0, schedule: 'Concediu maternitate'};
  // 8h/zi standard (Luni-Vineri)
  if (name.indexOf('BARABAN') >= 0 || name.indexOf('ONOFREI') >= 0 || name.indexOf('CIOBANICA') >= 0)
    return {type: 'standard', hPerDay: 8, schedule: '8h/zi L-V'};
  // 8h Luni-Vineri + 4h Sambata
  if (name.indexOf('BRSTYAK') >= 0 || name.indexOf('SPINU') >= 0 ||
      name.indexOf('BAKAITY') >= 0 || name.indexOf('ZDROBA') >= 0 ||
      (name.indexOf('CIOBANU') >= 0 && name.indexOf('LUMINITA') >= 0))
    return {type: 'sabsemilib', hPerDay: 8, hSab: 4, schedule: '8h L-V + 4h Sb'};
  // 10h/zi, 2 lucrate + 2 libere
  return {type: '2plus2', hPerDay: 10, schedule: '10h, 2+2'};
}

function calcNorma(empName, dates) {
  var norm = getEmpNorm(empName);
  if (norm.type === 'maternitate') return 0;
  var days = Math.round((dates.end - dates.start) / 86400000) + 1;
  if (norm.type === 'standard') {
    // Count working days Mon-Fri
    var wd = 0;
    for (var i = 0; i < days; i++) {
      var d = new Date(dates.start.getTime() + i*86400000);
      var dow = d.getDay();
      if (dow >= 1 && dow <= 5) wd++;
    }
    return wd * 8;
  }
  if (norm.type === 'sabsemilib') {
    var wd = 0, sabs = 0;
    for (var i = 0; i < days; i++) {
      var d = new Date(dates.start.getTime() + i*86400000);
      var dow = d.getDay();
      if (dow >= 1 && dow <= 5) wd++;
      if (dow === 6) sabs++;
    }
    return wd * 8 + sabs * 4;
  }
  if (norm.type === '2plus2') {
    return Math.floor(days / 4) * 2 * 10 + Math.min(days % 4, 2) * 10;
  }
  return 0;
}

function addPontaj(sh, row, pons, dates) {
  var emp = {};
  var maternitate = [];
  pons.forEach(function(r) {
    var e = r[1]; if (!e || e === 'TEST') return; // skip TEST
    var norm = getEmpNorm(e);
    if (norm.type === 'maternitate') {
      if (maternitate.indexOf(e) < 0) maternitate.push(e);
      return;
    }
    if (!emp[e]) emp[e] = {ore:0, zile:0, lib:0};
    var h = parseH(r[6]);
    if (h>0) { emp[e].ore+=h; emp[e].zile++; } else { emp[e].lib++; }
  });

  var sorted = Object.keys(emp).sort(function(a,b){return emp[b].ore-emp[a].ore;});
  var hdrs = ['#','Angajat','Program','Zile lucrate','Ore lucrate','Norma luna','+/- Norma','Status'];
  sh.setRowHeight(row, 26);
  hdrs.forEach(function(h,i) {
    sh.getRange(row,i+1).setValue(h)
      .setBackground(DC.dark).setFontColor(DC.gold)
      .setFontWeight('bold').setFontSize(10)
      .setHorizontalAlignment('center').setVerticalAlignment('middle');
  });
  row++;

  sorted.forEach(function(e, idx) {
    var d = emp[e];
    var norm = getEmpNorm(e);
    var normaOre = calcNorma(e, dates);
    var diff = d.ore - normaOre;
    var bg = idx%2===0?DC.white:'#f9f7f0';
    sh.setRowHeight(row, 22);
    var status = diff >= 0 ? 'Norma indeplinita' : Math.abs(diff) <= 8 ? 'Aproape' : 'Sub norma';
    var sColor = diff >= 0 ? '#1a6b3a' : Math.abs(diff) <= 8 ? DC.orange : DC.red;
    var diffStr = (diff >= 0 ? '+' : '') + diff.toFixed(1) + ' h';
    var vals = [idx+1, e, norm.schedule, d.zile,
      d.ore.toFixed(1)+' h', normaOre.toFixed(0)+' h', diffStr, status];
    vals.forEach(function(v,i) {
      var cell = sh.getRange(row,i+1).setValue(v).setBackground(bg).setFontSize(10).setVerticalAlignment('middle');
      if(i===6) cell.setFontColor(diff>=0?'#1a6b3a':DC.red).setFontWeight('bold');
      if(i===7) cell.setFontColor(sColor).setFontWeight('bold');
    });
    row++;
  });

  // Maternitate section
  if (maternitate.length) {
    sh.setRowHeight(row, 26);
    sh.getRange(row,1,1,8).merge()
      .setValue('Concediu maternitate: ' + maternitate.join(', '))
      .setBackground('#e8f0fe').setFontColor('#1a56db')
      .setFontSize(10).setFontStyle('italic').setVerticalAlignment('middle');
    row++;
  }
  row++; return row;
}

function addVanzari(sh, row, raps) {
  var mag = {};
  raps.forEach(function(r) {
    var m = r[1]; if (!m) return;
    if (!mag[m]) mag[m] = {vanz:0, pos:0, num:0, bon:0, zile:0, dates:{}};
    mag[m].vanz += parseFloat(r[3])||0;
    mag[m].pos  += parseFloat(r[4])||0;
    mag[m].num  += parseFloat(r[6])||0;
    mag[m].bon  += parseInt(r[8])||0;
    if (r[0]) mag[m].dates[String(r[0])] = 1;
    mag[m].zile = Object.keys(mag[m].dates).length || 1;
  });

  var sorted = Object.keys(mag).sort(function(a,b){return mag[b].vanz-mag[a].vanz;});
  var hdrs = ['#','Magazin','Vanzari total','POS total','Numerar total','Nr. bonuri','Medie/zi','Performanta'];
  sh.setRowHeight(row, 26);
  hdrs.forEach(function(h,i) {
    sh.getRange(row,i+1).setValue(h)
      .setBackground(DC.dark).setFontColor(DC.gold)
      .setFontWeight('bold').setFontSize(10)
      .setHorizontalAlignment('center').setVerticalAlignment('middle');
  });
  row++;

  sorted.forEach(function(m, idx) {
    var d = mag[m];
    var bg = idx%2===0?DC.white:'#f9f7f0';
    sh.setRowHeight(row, 22);
    var perf = d.vanz>=5000?'Top performer':d.vanz>=2000?'Bun':'Necesita atentie';
    var pc = d.vanz>=5000?'#1a6b3a':d.vanz>=2000?DC.orange:DC.red;
    var mediezi = d.zile>0 ? fmtR(d.vanz/d.zile) : '0.00 RON';
    var vals = [idx+1, m, fmtR(d.vanz), fmtR(d.pos), fmtR(d.num),
      d.bon, mediezi, perf];
    vals.forEach(function(v,i) {
      var cell = sh.getRange(row,i+1).setValue(v).setBackground(bg).setFontSize(10).setVerticalAlignment('middle');
      if(i===7) cell.setFontColor(pc).setFontWeight('bold');
    });
    row++;
  });
  row++; return row;
}

function addDiferente(sh, row, raps) {
  var difs = raps.filter(function(r){return Math.abs(parseFloat(r[7])||0)>0.01;});
  difs.sort(function(a,b){return Math.abs(parseFloat(b[7])||0)-Math.abs(parseFloat(a[7])||0);});

  if (!difs.length) {
    sh.setRowHeight(row, 38);
    sh.getRange(row,1,1,8).merge()
      .setValue('Felicitari! Nicio diferenta de casa in perioada selectata. Toate casierii sunt eroi!')
      .setBackground('#e8f8f0').setFontColor('#1a6b3a').setFontWeight('bold')
      .setFontSize(11).setVerticalAlignment('middle');
    return row+2;
  }

  var total = difs.reduce(function(s,r){return s+Math.abs(parseFloat(r[7])||0);},0);
  if (total>100) {
    sh.setRowHeight(row, 34);
    sh.getRange(row,1,1,8).merge()
      .setValue('ATENTIE! Diferente de ' + fmtR(total) + ' total. Ne trebuie explicatii, nu scuze creative!')
      .setBackground('#fff3e0').setFontColor(DC.red).setFontWeight('bold')
      .setFontSize(10).setVerticalAlignment('middle');
    row++;
  }

  var hdrs = ['Data','Magazin','Casier','Vanzari Z','Num. teoretic','Num. predat','Diferenta','Motiv'];
  sh.setRowHeight(row, 26);
  hdrs.forEach(function(h,i) {
    sh.getRange(row,i+1).setValue(h)
      .setBackground(DC.red).setFontColor(DC.white)
      .setFontWeight('bold').setFontSize(10)
      .setHorizontalAlignment('center').setVerticalAlignment('middle');
  });
  row++;

  difs.forEach(function(r, idx) {
    var dif = parseFloat(r[7])||0;
    var bg = idx%2===0?'#fff8f8':'#fff0f0';
    sh.setRowHeight(row, 22);
    var vals = [r[0],r[1],r[2],fmtR(parseFloat(r[3])||0),
      fmtR(parseFloat(r[5])||0),fmtR(parseFloat(r[6])||0),
      fmtR(dif), r[13]||'Nemotivat'];
    vals.forEach(function(v,i) {
      var cell = sh.getRange(row,i+1).setValue(v).setBackground(bg).setFontSize(10).setVerticalAlignment('middle');
      if(i===6) cell.setFontColor(dif<0?DC.red:DC.orange).setFontWeight('bold');
    });
    row++;
  });
  row++; return row;
}

function addDashFooter(sh, row, ncols, dateStr) {
  sh.setRowHeight(row, 6);
  sh.getRange(row,1,1,ncols).merge().setBackground(DC.gold); row++;
  sh.setRowHeight(row, 44);
  sh.getRange(row,1,1,ncols).merge()
    .setValue('REJOES Operation Management L2  |  ' + DASH_COMPANY + '  |  ' + dateStr + '\nRaport generat automat  |  ' + DASH_EMAIL + '  |  Confidential')
    .setBackground(DC.dark).setFontColor(DC.border)
    .setFontSize(9).setVerticalAlignment('middle').setHorizontalAlignment('center').setWrap(true);
}

function buildCustomPeriod() {
  var ui = SpreadsheetApp.getUi();
  var r1 = ui.prompt('Data START (DD.MM.YYYY):');
  if (r1.getSelectedButton() !== ui.Button.OK) return;
  var r2 = ui.prompt('Data END (DD.MM.YYYY):');
  if (r2.getSelectedButton() !== ui.Button.OK) return;
  var from = dashParseDate(r1.getResponseText().trim());
  var to   = dashParseDate(r2.getResponseText().trim());
  if (!from||!to) { ui.alert('Format invalid. Foloseste DD.MM.YYYY'); return; }
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  buildDash(ss, {start:from, end:to});
  ui.alert('Dashboard actualizat!');
}

function sendFullReport() {
  var ui = SpreadsheetApp.getUi();
  var r1 = ui.prompt('Data START (DD.MM.YYYY):');
  if (r1.getSelectedButton() !== ui.Button.OK) return;
  var r2 = ui.prompt('Data END (DD.MM.YYYY):');
  if (r2.getSelectedButton() !== ui.Button.OK) return;
  var r3 = ui.prompt('Destinatari email (virgula intre adrese):', 'adrian.m@gmail.com, dana.lucean@yahoo.com, olikvp@gmail.com', ui.ButtonSet.OK_CANCEL);
  if (r3.getSelectedButton() !== ui.Button.OK) return;

  var from  = dashParseDate(r1.getResponseText().trim());
  var to    = dashParseDate(r2.getResponseText().trim());
  var recip = r3.getResponseText().trim();
  if (!from||!to) { ui.alert('Format invalid!'); return; }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  buildDash(ss, {start:from, end:to});

  var sh = ss.getSheetByName(DASH_SHEET);
  var pdfUrl = 'https://docs.google.com/spreadsheets/d/' + ss.getId() +
    '/export?format=pdf&size=A4&portrait=false&fitw=true&sheetnames=false' +
    '&printtitle=false&pagenumbers=true&gridlines=false&fzr=false&gid=' + sh.getSheetId();
  var pdfBlob = UrlFetchApp.fetch(pdfUrl, {
    headers:{Authorization:'Bearer '+ScriptApp.getOAuthToken()}
  }).getBlob().setName('Rejoes_Raport_' + dFmt(from) + '_' + dFmt(to) + '.pdf');

  var xlsUrl = 'https://docs.google.com/spreadsheets/d/' + ss.getId() + '/export?format=xlsx';
  var xlsBlob = UrlFetchApp.fetch(xlsUrl, {
    headers:{Authorization:'Bearer '+ScriptApp.getOAuthToken()}
  }).getBlob().setName('Rejoes_Raport_' + dFmt(from) + '_' + dFmt(to) + '.xlsx');

  var recipList = recip.split(',').map(function(e){return e.trim();}).filter(Boolean);
  var toAddr = recipList.shift();
  var ccAddr = recipList.join(',');

  MailApp.sendEmail(toAddr, 'Rejoes - Raport Operational ' + dFmt(from) + ' - ' + dFmt(to), '', {
    htmlBody: '<h2>Raport Operational Rejoes</h2><p>Perioada: '+dFmt(from)+' - '+dFmt(to)+'</p><p>Gasiti raportul complet atasat in format PDF si Excel.</p><p>'+DASH_COMPANY+'</p>',
    attachments:[pdfBlob, xlsBlob],
    cc: ccAddr||'',
    name:'Rejoes Raport Automat'
  });
  ui.alert('Raport trimis cu succes catre: ' + recip);
}

function getRange(ss, sheetName, dates) {
  var ws = ss.getSheetByName(sheetName);
  if (!ws || ws.getLastRow()<2) return [];
  var data = ws.getRange(2,1,ws.getLastRow()-1,ws.getLastColumn()).getValues();
  return data.filter(function(r) {
    var d = dashParseAny(r[0]);
    return d && d>=dates.start && d<=dates.end;
  });
}

function dashParseAny(v) {
  if (!v) return null;
  if (v instanceof Date) return v;
  var s = String(v).trim();
  var m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (m) return new Date(+m[1],+m[2]-1,+m[3]);
  m = s.match(/^(\d{2})\.(\d{2})\.(\d{4})/);
  if (m) return new Date(+m[3],+m[2]-1,+m[1]);
  return null;
}

function dashParseDate(s) {
  var m = s.match(/^(\d{2})\.(\d{2})\.(\d{4})$/);
  if (!m) return null;
  return new Date(+m[3],+m[2]-1,+m[1]);
}

function parseH(v) {
  if (!v) return 0;
  var m = String(v).match(/([\d.]+)/);
  return m ? parseFloat(m[1]) : 0;
}

function dFmt(d) {
  if (!d) return '';
  return String(d.getDate()).padStart(2,'0')+'.'+String(d.getMonth()+1).padStart(2,'0')+'.'+d.getFullYear();
}

function dFmtFull(d) {
  return dFmt(d)+' '+String(d.getHours()).padStart(2,'0')+':'+String(d.getMinutes()).padStart(2,'0');
}

function fmtR(v) {
  return parseFloat(v||0).toFixed(2)+' RON';
}
