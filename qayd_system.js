// ══════════════════════════════════════════════════════
// Qayd — قيد واحد → يظهر في كل مكان
// Part 1: Core + Settings + Entry Log
// ══════════════════════════════════════════════════════

var C = {pri:'#1B3A5C',acc:'#E8A838',suc:'#2E7D32',dan:'#C62828',wht:'#FFFFFF',drk:'#263238',bg:'#F5F5F5',
  hdr:'#BDD6EE',hdr2:'#FFF2CC',hdr3:'#D9D2E9',lGrn:'#E8F5E9',lRed:'#FFEBEE',lYlw:'#FFF9C4'};

var SN = {
  entry:'سجل القيود',settings:'الإعدادات',
  cash:'الصندوق',rajhi:'بنك الراجحي',ahli:'بنك الأهلي',
  projects:'المشاريع',custody:'العهد',interco:'التقاص',
  incomeW:'قائمة الدخل الأسبوعية',incomeM:'قائمة الدخل الشهرية',
  dash:'لوحة التحكم',reports:'التقارير'
};

// ═══ HELPERS ═══
function getOrCreate_(ss,n){var s=ss.getSheetByName(n);if(!s)s=ss.insertSheet(n);return s;}
function fmtH_(sh,r,c,bg){sh.getRange(r,1,1,c).setBackground(bg).setFontColor(C.wht).setFontWeight('bold').setHorizontalAlignment('center');}
function batch_(sh,sr,er,col,fn){
  var f=[];for(var r=sr;r<=er;r++)f.push([fn(r)]);
  sh.getRange(sr,col,f.length,1).setFormulas(f);
}
function addDV_(sh,r,c,n,list){
  var rule=SpreadsheetApp.newDataValidation().requireValueInList(list,true).setAllowInvalid(false).build();
  sh.getRange(r,c,n,1).setDataValidation(rule);
}

// ═══ MENU ═══
function onOpen(){
  SpreadsheetApp.getUi().createMenu('⚡ Qayd')
    .addItem('🔧 بناء النظام','setupQayd')
    .addItem('🔒 حماية المعادلات','protectQayd')
    .addToUi();
}

// ═══ MAIN SETUP ═══
function setupQayd(){
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var ui=SpreadsheetApp.getUi();
  var r=ui.alert('⚡ Qayd','سيتم بناء النظام الكامل. هل تريد المتابعة؟',ui.ButtonSet.YES_NO);
  if(r!==ui.Button.YES)return;

  buildSettings_(ss);SpreadsheetApp.flush();
  buildEntryLog_(ss);SpreadsheetApp.flush();
  buildCashSheet_(ss);SpreadsheetApp.flush();
  buildBankSheet_(ss,SN.rajhi,'الراجحي');SpreadsheetApp.flush();
  buildBankSheet_(ss,SN.ahli,'الأهلي');SpreadsheetApp.flush();
  buildProjectsSummary_(ss);SpreadsheetApp.flush();
  buildCustodySheet_(ss);SpreadsheetApp.flush();
  buildIntercoSheet_(ss);SpreadsheetApp.flush();
  buildWeeklyIncome_(ss);SpreadsheetApp.flush();
  buildMonthlyIncome_(ss);SpreadsheetApp.flush();
  buildDashboard_(ss);SpreadsheetApp.flush();

  // Hide processing sheets
  [SN.cash,SN.rajhi,SN.ahli,SN.custody,SN.interco].forEach(function(n){
    var s=ss.getSheetByName(n);if(s)s.hideSheet();
  });

  var def=ss.getSheetByName('Sheet1');
  if(def&&ss.getSheets().length>1)ss.deleteSheet(def);
  ss.getSheetByName(SN.entry).activate();
  ui.alert('✅ تم بناء Qayd بنجاح!');
}

// ═══ SETTINGS ═══
function buildSettings_(ss){
  var sh=getOrCreate_(ss,SN.settings);
  sh.clear();sh.setRightToLeft(true);

  sh.getRange(1,1).setValue('إعدادات Qayd').setFontSize(16).setFontWeight('bold');
  sh.getRange(1,1,1,4).merge().setBackground(C.pri).setFontColor(C.wht).setHorizontalAlignment('center');

  // Company info
  sh.getRange(3,1).setValue('اسم الشركة:').setFontWeight('bold');
  sh.getRange(3,2).setValue('شركة إعمار لاند');
  sh.getRange(4,1).setValue('السنة المالية:').setFontWeight('bold');
  sh.getRange(4,2).setValue(new Date().getFullYear());

  // Expense categories
  sh.getRange(6,1).setValue('تصنيفات المصروفات').setFontWeight('bold').setBackground(C.hdr);
  sh.getRange(6,1,1,2).merge();
  var cats=['مصروفات إدارية عامة','محروقات (بنزين+زيت+ديزل)','وجبات ومشروبات','مشتريات عامة',
    'رواتب ويوميات وسلف','صيانة سيارات','أصول ثابتة','إيجارات','كهرباء وماء','اتصالات وإنترنت',
    'مصروفات تسويقية','رسوم حكومية','مصروفات نقل','مصروفات متنوعة'];
  for(var i=0;i<cats.length;i++) sh.getRange(7+i,1).setValue(i+1), sh.getRange(7+i,2).setValue(cats[i]);

  // Revenue categories
  var rr=7+cats.length+1;
  sh.getRange(rr,1).setValue('تصنيفات الإيرادات').setFontWeight('bold').setBackground(C.lGrn);
  sh.getRange(rr,1,1,2).merge();
  var revs=['مبيعات نقدية','مبيعات آجلة','إيرادات مشاريع','إيرادات خدمات','إيرادات أخرى'];
  for(var j=0;j<revs.length;j++) sh.getRange(rr+1+j,1).setValue(j+1), sh.getRange(rr+1+j,2).setValue(revs[j]);

  // Projects
  var pr=rr+revs.length+2;
  sh.getRange(pr,1).setValue('المشاريع').setFontWeight('bold').setBackground(C.hdr2);
  sh.getRange(pr,1,1,3).merge();
  sh.getRange(pr+1,1,1,3).setValues([['كود','اسم المشروع','العميل']]).setFontWeight('bold');
  var projs=[['P01','مشروع السجون','وزارة الداخلية'],['P02','مشروع المزاحمية','القطاع الخاص'],
    ['P03','فلل جدة','شركة البناء'],['P04','مشروع الرمال','درر العواصم']];
  sh.getRange(pr+2,1,projs.length,3).setValues(projs);

  // Employees
  var er=pr+projs.length+3;
  sh.getRange(er,1).setValue('الموظفون').setFontWeight('bold').setBackground(C.hdr3);
  sh.getRange(er,1,1,2).merge();
  var emps=['سعدية','طارق','محمد فهيم','أحمد','عبدالله'];
  for(var e=0;e<emps.length;e++) sh.getRange(er+1+e,1).setValue(e+1), sh.getRange(er+1+e,2).setValue(emps[e]);

  // Intercompany entities
  var ir=er+emps.length+2;
  sh.getRange(ir,1).setValue('كيانات التقاص').setFontWeight('bold').setBackground(C.acc).setFontColor(C.wht);
  sh.getRange(ir,1,1,2).merge();
  var ents=['إعمار','مسار','الفرع الرياض'];
  for(var n=0;n<ents.length;n++) sh.getRange(ir+1+n,1).setValue(n+1), sh.getRange(ir+1+n,2).setValue(ents[n]);

  // Named ranges
  ss.setNamedRange('rng_ExpCats',sh.getRange(7,2,cats.length,1));
  ss.setNamedRange('rng_RevCats',sh.getRange(rr+1,2,revs.length,1));
  ss.setNamedRange('rng_Projects',sh.getRange(pr+2,2,projs.length,1));
  ss.setNamedRange('rng_Employees',sh.getRange(er+1,2,emps.length,1));
  ss.setNamedRange('rng_Entities',sh.getRange(ir+1,2,ents.length,1));

  sh.setColumnWidth(1,30);sh.setColumnWidth(2,220);sh.setColumnWidth(3,160);
  sh.setTabColor(C.pri);
}

// ═══ ENTRY LOG — The Single Input Sheet ═══
// This is the ONLY sheet Essam types into. Everything else auto-populates.
function buildEntryLog_(ss){
  var sh=getOrCreate_(ss,SN.entry);
  sh.clear();
  if(sh.getMaxRows()<1010) sh.insertRowsAfter(sh.getMaxRows(),1010-sh.getMaxRows());
  sh.getRange(1,1,sh.getMaxRows(),sh.getMaxColumns()).clearDataValidations();
  sh.setRightToLeft(true);

  // Title
  sh.getRange(1,1).setValue('⚡ Qayd — سجل القيود اليومي').setFontSize(16).setFontWeight('bold');
  sh.getRange(1,1,1,10).merge().setBackground(C.pri).setFontColor(C.wht).setHorizontalAlignment('center');

  // Live balance bar (row 2) — instant feedback while entering
  var balItems=[
    {l:'💰 الصندوق',f:"=SUMIFS(D4:D1003,F4:F1003,\"صندوق\",C4:C1003,\"وارد\")-SUMIFS(D4:D1003,F4:F1003,\"صندوق\",C4:C1003,\"صادر\")"},
    {l:'🏦 الراجحي',f:"=SUMIFS(D4:D1003,F4:F1003,\"بنك الراجحي\",C4:C1003,\"وارد\")-SUMIFS(D4:D1003,F4:F1003,\"بنك الراجحي\",C4:C1003,\"صادر\")"},
    {l:'🏦 الأهلي',f:"=SUMIFS(D4:D1003,F4:F1003,\"بنك الأهلي\",C4:C1003,\"وارد\")-SUMIFS(D4:D1003,F4:F1003,\"بنك الأهلي\",C4:C1003,\"صادر\")"},
    {l:'📊 الإجمالي',f:"=B2+D2+F2"}
  ];
  for(var b=0;b<balItems.length;b++){
    var bc=1+b*2;
    sh.getRange(2,bc).setValue(balItems[b].l).setFontWeight('bold').setFontSize(11).setBackground(C.lYlw);
    sh.getRange(2,bc+1).setFormula(balItems[b].f).setFontWeight('bold').setFontSize(13).setNumberFormat('#,##0.00').setBackground(C.lYlw);
  }
  // Alert indicator
  sh.getRange(2,9).setValue('⚡').setFontSize(14).setBackground(C.acc).setFontColor(C.wht).setHorizontalAlignment('center');
  sh.getRange(2,10).setFormula('=IF(H2<0,"⚠️ عجز!","✅ متوازن")').setFontWeight('bold').setFontSize(11).setBackground(C.acc).setFontColor(C.wht);

  // Column headers (row 3) — only 8 input columns
  var headers=['التاريخ','البيان','نوع الحركة','المبلغ','التصنيف','المصدر','المشروع','الطرف المقابل','#','الشهر'];
  sh.getRange(3,1,1,10).setValues([headers]);
  fmtH_(sh,3,10,C.drk);
  var widths=[110,250,80,120,180,120,130,150,40,80];
  for(var w=0;w<widths.length;w++) sh.setColumnWidth(w+1,widths[w]);

  // Data validation dropdowns
  var DR=1000;
  addDV_(sh,4,3,DR,['وارد','صادر','مردود','تحويل']);
  addDV_(sh,4,6,DR,['صندوق','بنك الراجحي','بنك الأهلي']);

  // Link to named ranges for categories
  var settSh=ss.getSheetByName(SN.settings);
  if(settSh){
    var expR=ss.getRangeByName('rng_ExpCats');
    var revR=ss.getRangeByName('rng_RevCats');
    var projR=ss.getRangeByName('rng_Projects');
    var empR=ss.getRangeByName('rng_Employees');
    // Combined expense+revenue categories for dropdown
    if(expR&&revR){
      var allCats=expR.getValues().concat(revR.getValues()).map(function(r){return r[0];}).filter(function(v){return v!=='';});
      addDV_(sh,4,5,DR,allCats);
    }
    if(projR){
      var rule=SpreadsheetApp.newDataValidation().requireValueInRange(projR,true).setAllowInvalid(true).build();
      sh.getRange(4,7,DR,1).setDataValidation(rule);
    }
    if(empR){
      var ents=ss.getRangeByName('rng_Entities');
      var allPeople=empR.getValues().concat(ents?ents.getValues():[]).map(function(r){return r[0];}).filter(function(v){return v!=='';});
      addDV_(sh,4,8,DR,allPeople);
    }
  }

  // Auto-number and auto-month formulas
  batch_(sh,4,4+DR-1,9,function(r){return '=IF(A'+r+'<>"",ROW()-3,"")';});
  batch_(sh,4,4+DR-1,10,function(r){return '=IF(A'+r+'<>"",TEXT(A'+r+',"MMMM"),"")';});

  // Formatting
  sh.getRange(4,1,DR,1).setNumberFormat('yyyy-mm-dd');
  sh.getRange(4,4,DR,1).setNumberFormat('#,##0.00');
  sh.setFrozenRows(3);
  sh.setTabColor(C.acc);

  // ═══ EMBEDDED SUMMARY — Below data (like Essam's style) ═══
  var sumRow=4+DR+2;
  sh.getRange(sumRow,1).setValue('📋 ملخص سريع').setFontSize(14).setFontWeight('bold');
  sh.getRange(sumRow,1,1,8).merge().setBackground(C.pri).setFontColor(C.wht).setHorizontalAlignment('center');

  var sumHeaders=['التصنيف','إجمالي صادر','إجمالي وارد','الصافي'];
  sh.getRange(sumRow+1,1,1,4).setValues([sumHeaders]).setFontWeight('bold').setBackground(C.hdr);

  // Auto summary per category
  var cats=['مصروفات إدارية عامة','محروقات (بنزين+زيت+ديزل)','وجبات ومشروبات','مشتريات عامة',
    'رواتب ويوميات وسلف','صيانة سيارات','إيجارات','مبيعات نقدية','مبيعات آجلة','إيرادات مشاريع'];
  for(var ci=0;ci<cats.length;ci++){
    var cr=sumRow+2+ci;
    sh.getRange(cr,1).setValue(cats[ci]);
    sh.getRange(cr,2).setFormula('=SUMIFS(D$4:D$1003,E$4:E$1003,"'+cats[ci]+'",C$4:C$1003,"صادر")').setNumberFormat('#,##0.00');
    sh.getRange(cr,3).setFormula('=SUMIFS(D$4:D$1003,E$4:E$1003,"'+cats[ci]+'",C$4:C$1003,"وارد")').setNumberFormat('#,##0.00');
    sh.getRange(cr,4).setFormula('=C'+cr+'-B'+cr).setNumberFormat('#,##0.00');
  }
  var totR=sumRow+2+cats.length;
  sh.getRange(totR,1).setValue('الإجمالي الكلي').setFontWeight('bold').setBackground(C.lYlw);
  sh.getRange(totR,2).setFormula('=SUM(B'+(sumRow+2)+':B'+(totR-1)+')').setFontWeight('bold').setNumberFormat('#,##0.00');
  sh.getRange(totR,3).setFormula('=SUM(C'+(sumRow+2)+':C'+(totR-1)+')').setFontWeight('bold').setNumberFormat('#,##0.00');
  sh.getRange(totR,4).setFormula('=C'+totR+'-B'+totR).setFontWeight('bold').setNumberFormat('#,##0.00').setFontSize(12);
}

// ══════════════════════════════════════════════════════
// Qayd Part 2: Auto-Processing Sheets (Hidden)
// All fed automatically from سجل القيود via FILTER
// ══════════════════════════════════════════════════════

// ═══ CASH SHEET (auto from entry log where source=صندوق) ═══
function buildCashSheet_(ss){
  var sh=getOrCreate_(ss,SN.cash);sh.clear();sh.setRightToLeft(true);
  sh.getRange(1,1).setValue('حركات الصندوق (تلقائي)').setFontSize(14).setFontWeight('bold');
  sh.getRange(1,1,1,8).merge().setBackground(C.pri).setFontColor(C.wht).setHorizontalAlignment('center');
  sh.getRange(2,1,1,8).setValues([['التاريخ','البيان','نوع','المبلغ','التصنيف','المشروع','الطرف','الرصيد']]);
  fmtH_(sh,2,8,C.drk);
  // FILTER from entry log — pure FILTER, no QUERY wrapper (avoids Col1 type ambiguity)
  var e="'"+SN.entry+"'!";
  sh.getRange(3,1).setFormula(
    '=IFERROR(SORT(FILTER({'+e+'A4:A1003,'+e+'B4:B1003,'+e+'C4:C1003,'+e+'D4:D1003,'+e+'E4:E1003,'+e+'G4:G1003,'+e+'H4:H1003},'+e+'F4:F1003="صندوق",'+e+'A4:A1003>0),1,TRUE),"")'
  );
  // Running balance
  batch_(sh,3,502,8,function(r){
    return '=IF(A'+r+'="","",SUMPRODUCT((A$3:A'+r+'<>"")*IF(C$3:C'+r+'="وارد",D$3:D'+r+',-D$3:D'+r+')))';
  });
  sh.getRange(3,1,500,1).setNumberFormat('yyyy-mm-dd');
  sh.getRange(3,4,500,1).setNumberFormat('#,##0.00');
  sh.getRange(3,8,500,1).setNumberFormat('#,##0.00');
  sh.setFrozenRows(2);sh.setTabColor('#4CAF50');
}

// ═══ BANK SHEET (reusable for Rajhi/Ahli) ═══
function buildBankSheet_(ss,sheetName,bankLabel){
  var sh=getOrCreate_(ss,sheetName);sh.clear();sh.setRightToLeft(true);
  sh.getRange(1,1).setValue('حركات '+bankLabel+' (تلقائي)').setFontSize(14).setFontWeight('bold');
  sh.getRange(1,1,1,8).merge().setBackground('#1565C0').setFontColor(C.wht).setHorizontalAlignment('center');
  sh.getRange(2,1,1,8).setValues([['التاريخ','البيان','نوع','المبلغ','التصنيف','المشروع','الطرف','الرصيد']]);
  fmtH_(sh,2,8,C.drk);
  var e="'"+SN.entry+"'!";
  sh.getRange(3,1).setFormula(
    '=IFERROR(SORT(FILTER({'+e+'A4:A1003,'+e+'B4:B1003,'+e+'C4:C1003,'+e+'D4:D1003,'+e+'E4:E1003,'+e+'G4:G1003,'+e+'H4:H1003},'+e+'F4:F1003="'+sheetName+'",'+e+'A4:A1003>0),1,TRUE),"")'
  );
  batch_(sh,3,502,8,function(r){
    return '=IF(A'+r+'="","",SUMPRODUCT((A$3:A'+r+'<>"")*IF(C$3:C'+r+'="وارد",D$3:D'+r+',-D$3:D'+r+')))';
  });
  sh.getRange(3,1,500,1).setNumberFormat('yyyy-mm-dd');
  sh.getRange(3,4,500,1).setNumberFormat('#,##0.00');
  sh.getRange(3,8,500,1).setNumberFormat('#,##0.00');
  sh.setFrozenRows(2);sh.setTabColor('#1565C0');
}

// ═══ PROJECTS SUMMARY (auto P&L per project) ═══
function buildProjectsSummary_(ss){
  var sh=getOrCreate_(ss,SN.projects);sh.clear();sh.setRightToLeft(true);
  sh.getRange(1,1).setValue('ربحية المشاريع (تلقائي)').setFontSize(14).setFontWeight('bold');
  sh.getRange(1,1,1,8).merge().setBackground(C.acc).setFontColor(C.wht).setHorizontalAlignment('center');
  sh.getRange(2,1,1,8).setValues([['المشروع','إيرادات','مصروفات','صافي ربح/خسارة','نسبة الربح %','عدد الحركات','آخر حركة','حالة']]);
  fmtH_(sh,2,8,C.drk);
  var e="'"+SN.entry+"'!";
  var projs=['مشروع السجون','مشروع المزاحمية','فلل جدة','مشروع الرمال'];
  for(var i=0;i<projs.length;i++){
    var r=3+i,p=projs[i];
    sh.getRange(r,1).setValue(p);
    sh.getRange(r,2).setFormula('=SUMIFS('+e+'D$4:D$1003,'+e+'G$4:G$1003,"'+p+'",'+e+'C$4:C$1003,"وارد")').setNumberFormat('#,##0.00');
    sh.getRange(r,3).setFormula('=SUMIFS('+e+'D$4:D$1003,'+e+'G$4:G$1003,"'+p+'",'+e+'C$4:C$1003,"صادر")').setNumberFormat('#,##0.00');
    sh.getRange(r,4).setFormula('=B'+r+'-C'+r).setNumberFormat('#,##0.00');
    sh.getRange(r,5).setFormula('=IF(B'+r+'>0,ROUND(D'+r+'/B'+r+'*100,1),"")');
    sh.getRange(r,6).setFormula('=COUNTIF('+e+'G$4:G$1003,"'+p+'")');
    sh.getRange(r,7).setFormula('=IFERROR(INDEX('+e+'A$4:A$1003,MATCH(2,1/('+e+'G$4:G$1003="'+p+'"),1)),"")').setNumberFormat('yyyy-mm-dd');
    sh.getRange(r,8).setFormula('=IF(D'+r+'>0,"✅ ربح",IF(D'+r+'<0,"❌ خسارة","⚪ متعادل"))');
  }
  sh.setFrozenRows(2);sh.setTabColor('#FF6F00');
}

// ═══ CUSTODY (auto per employee) ═══
function buildCustodySheet_(ss){
  var sh=getOrCreate_(ss,SN.custody);sh.clear();sh.setRightToLeft(true);
  sh.getRange(1,1).setValue('تتبع العهد (تلقائي)').setFontSize(14).setFontWeight('bold');
  sh.getRange(1,1,1,5).merge().setBackground('#7B1FA2').setFontColor(C.wht).setHorizontalAlignment('center');
  sh.getRange(2,1,1,5).setValues([['الموظف','عهد مسلمة','عهد مسواة','رصيد عهدة','حالة']]);
  fmtH_(sh,2,5,C.drk);
  var e="'"+SN.entry+"'!";
  var emps=['سعدية','طارق','محمد فهيم','أحمد','عبدالله'];
  for(var i=0;i<emps.length;i++){
    var r=3+i,emp=emps[i];
    sh.getRange(r,1).setValue(emp);
    sh.getRange(r,2).setFormula('=SUMIFS('+e+'D$4:D$1003,'+e+'H$4:H$1003,"'+emp+'",'+e+'C$4:C$1003,"صادر")').setNumberFormat('#,##0.00');
    sh.getRange(r,3).setFormula('=SUMIFS('+e+'D$4:D$1003,'+e+'H$4:H$1003,"'+emp+'",'+e+'C$4:C$1003,"وارد")').setNumberFormat('#,##0.00');
    sh.getRange(r,4).setFormula('=B'+r+'-C'+r).setNumberFormat('#,##0.00');
    sh.getRange(r,5).setFormula('=IF(D'+r+'>0,"⚠️ عهدة مفتوحة",IF(D'+r+'=0,"✅ مسوّاة","🔄 دائن"))');
  }
  sh.setTabColor('#7B1FA2');
}

// ═══ INTERCOMPANY (Emaar ↔ Masar auto-reconciliation) ═══
function buildIntercoSheet_(ss){
  var sh=getOrCreate_(ss,SN.interco);sh.clear();sh.setRightToLeft(true);
  sh.getRange(1,1).setValue('حسابات التقاص بين الكيانات (تلقائي)').setFontSize(14).setFontWeight('bold');
  sh.getRange(1,1,1,5).merge().setBackground('#E65100').setFontColor(C.wht).setHorizontalAlignment('center');
  sh.getRange(2,1,1,5).setValues([['الكيان','له (وارد)','عليه (صادر)','الرصيد الصافي','الاتجاه']]);
  fmtH_(sh,2,5,C.drk);
  var e="'"+SN.entry+"'!";
  var ents=['إعمار','مسار','الفرع الرياض'];
  for(var i=0;i<ents.length;i++){
    var r=3+i,ent=ents[i];
    sh.getRange(r,1).setValue(ent);
    sh.getRange(r,2).setFormula('=SUMIFS('+e+'D$4:D$1003,'+e+'H$4:H$1003,"'+ent+'",'+e+'C$4:C$1003,"وارد")').setNumberFormat('#,##0.00');
    sh.getRange(r,3).setFormula('=SUMIFS('+e+'D$4:D$1003,'+e+'H$4:H$1003,"'+ent+'",'+e+'C$4:C$1003,"صادر")').setNumberFormat('#,##0.00');
    sh.getRange(r,4).setFormula('=B'+r+'-C'+r).setNumberFormat('#,##0.00');
    sh.getRange(r,5).setFormula('=IF(D'+r+'>0,"⬅️ لنا عنده",IF(D'+r+'<0,"➡️ لهم عندنا","⚪ متساوي"))');
  }
  sh.setTabColor('#E65100');
}

// ══════════════════════════════════════════════════════
// Qayd Part 3: Income Statements + Dashboard + Protection
// ══════════════════════════════════════════════════════

// ═══ WEEKLY INCOME STATEMENT ═══
function buildWeeklyIncome_(ss){
  var sh=getOrCreate_(ss,SN.incomeW);sh.clear();sh.setRightToLeft(true);
  sh.getRange(1,1).setValue('قائمة الدخل الأسبوعية').setFontSize(16).setFontWeight('bold');
  sh.getRange(1,1,1,6).merge().setBackground(C.suc).setFontColor(C.wht).setHorizontalAlignment('center');

  // Week selector
  sh.getRange(2,1).setValue('من تاريخ:').setFontWeight('bold');
  sh.getRange(2,3).setValue('إلى تاريخ:').setFontWeight('bold');
  sh.getRange(2,2).setValue(new Date()).setNumberFormat('yyyy-mm-dd').setBackground(C.lYlw);
  sh.getRange(2,4).setValue(new Date()).setNumberFormat('yyyy-mm-dd').setBackground(C.lYlw);

  var e="'"+SN.entry+"'!";
  var from='$B$2',to='$D$2';

  // Income Statement Structure
  var rows=[
    {r:4,l:'المبيعات والإيرادات',bg:C.hdr,bold:true,isHeader:true},
    {r:5,l:'مبيعات نقدية',cat:'مبيعات نقدية',type:'وارد'},
    {r:6,l:'مبيعات آجلة',cat:'مبيعات آجلة',type:'وارد'},
    {r:7,l:'إيرادات مشاريع',cat:'إيرادات مشاريع',type:'وارد'},
    {r:8,l:'إيرادات خدمات',cat:'إيرادات خدمات',type:'وارد'},
    {r:9,l:'مردودات المبيعات',cat:'مبيعات نقدية',type:'مردود',neg:true},
    {r:10,l:'صافي المبيعات',bg:C.lGrn,bold:true,sum:'B5:B8',sub:'B9'},
    {r:12,l:'المشتريات',bg:C.hdr2,bold:true,isHeader:true},
    {r:13,l:'مشتريات عامة',cat:'مشتريات عامة',type:'صادر'},
    {r:14,l:'مردودات المشتريات',cat:'مشتريات عامة',type:'مردود'},
    {r:15,l:'صافي المشتريات',bg:C.lYlw,bold:true,sum:'B13',sub:'B14'},
    {r:17,l:'مجمل الربح',bg:'#C8E6C9',bold:true,big:true,formula:'=B10-B15'},
    {r:19,l:'المصروفات التشغيلية',bg:C.hdr3,bold:true,isHeader:true},
    {r:20,l:'مصروفات إدارية عامة',cat:'مصروفات إدارية عامة',type:'صادر'},
    {r:21,l:'محروقات',cat:'محروقات (بنزين+زيت+ديزل)',type:'صادر'},
    {r:22,l:'وجبات ومشروبات',cat:'وجبات ومشروبات',type:'صادر'},
    {r:23,l:'رواتب ويوميات وسلف',cat:'رواتب ويوميات وسلف',type:'صادر'},
    {r:24,l:'صيانة سيارات',cat:'صيانة سيارات',type:'صادر'},
    {r:25,l:'إيجارات',cat:'إيجارات',type:'صادر'},
    {r:26,l:'كهرباء وماء',cat:'كهرباء وماء',type:'صادر'},
    {r:27,l:'مصروفات تسويقية',cat:'مصروفات تسويقية',type:'صادر'},
    {r:28,l:'مصروفات متنوعة',cat:'مصروفات متنوعة',type:'صادر'},
    {r:29,l:'إجمالي المصروفات',bg:C.lRed,bold:true,sum:'B20:B28'},
    {r:31,l:'صافي الدخل',bg:'#1B5E20',color:C.wht,bold:true,big:true,formula:'=B17-B29'},
  ];

  rows.forEach(function(item){
    sh.getRange(item.r,1).setValue(item.l);
    if(item.bold) sh.getRange(item.r,1).setFontWeight('bold');
    if(item.bg) sh.getRange(item.r,1,1,2).setBackground(item.bg);
    if(item.color) sh.getRange(item.r,1,1,2).setFontColor(item.color);
    if(item.big) sh.getRange(item.r,2).setFontSize(14);

    if(item.cat){
      var sign=item.neg?-1:1;
      sh.getRange(item.r,2).setFormula(
        '=SUMPRODUCT(('+e+'A$4:A$1003>='+from+')*('+e+'A$4:A$1003<='+to+')*('+e+'E$4:E$1003="'+item.cat+'")*('+e+'C$4:C$1003="'+item.type+'")*'+e+'D$4:D$1003)*'+sign
      );
    }
    if(item.sum&&!item.sub) sh.getRange(item.r,2).setFormula('=SUM('+item.sum+')');
    if(item.sum&&item.sub) sh.getRange(item.r,2).setFormula('=SUM('+item.sum+')-'+item.sub);
    if(item.formula) sh.getRange(item.r,2).setFormula(item.formula);

    sh.getRange(item.r,2).setNumberFormat('#,##0.00');
  });

  sh.setColumnWidth(1,250);sh.setColumnWidth(2,150);
  sh.setFrozenRows(3);sh.setTabColor(C.suc);
}

// ═══ MONTHLY INCOME STATEMENT (Cash + Accrual side by side) ═══
function buildMonthlyIncome_(ss){
  var sh=getOrCreate_(ss,SN.incomeM);sh.clear();sh.setRightToLeft(true);
  sh.getRange(1,1).setValue('قائمة الدخل الشهرية — نقدي واستحقاق').setFontSize(14).setFontWeight('bold');
  sh.getRange(1,1,1,4).merge().setBackground(C.suc).setFontColor(C.wht).setHorizontalAlignment('center');

  sh.getRange(2,1).setValue('الشهر:').setFontWeight('bold');
  sh.getRange(2,2).setValue('يناير').setBackground(C.lYlw);
  addDV_(sh,2,2,1,['يناير','فبراير','مارس','أبريل','مايو','يونيو','يوليو','أغسطس','سبتمبر','أكتوبر','نوفمبر','ديسمبر']);

  sh.getRange(3,1).setValue('البيان').setFontWeight('bold');
  sh.getRange(3,2).setValue('أساس نقدي').setFontWeight('bold').setBackground('#A5D6A7');
  sh.getRange(3,3).setValue('أساس استحقاق').setFontWeight('bold').setBackground('#BBDEFB');
  sh.getRange(3,4).setValue('الفرق').setFontWeight('bold').setBackground(C.lYlw);

  var e="'"+SN.entry+"'!";
  var mo='$B$2';
  var items=[
    {r:5,l:'إجمالي المبيعات',type:'وارد'},
    {r:6,l:'مردودات المبيعات',type:'مردود',neg:true},
    {r:7,l:'صافي المبيعات',formula:'=B5-B6',bold:true,bg:C.lGrn},
    {r:9,l:'مشتريات',type:'صادر',cat:'مشتريات عامة'},
    {r:10,l:'رواتب',type:'صادر',cat:'رواتب ويوميات وسلف'},
    {r:11,l:'محروقات',type:'صادر',cat:'محروقات (بنزين+زيت+ديزل)'},
    {r:12,l:'إيجارات',type:'صادر',cat:'إيجارات'},
    {r:13,l:'مصروفات إدارية',type:'صادر',cat:'مصروفات إدارية عامة'},
    {r:14,l:'مصروفات أخرى',type:'صادر',other:true},
    {r:15,l:'إجمالي المصروفات',formula:'=SUM(B9:B14)',bold:true,bg:C.lRed},
    {r:17,l:'صافي الدخل',formula:'=B7-B15',bold:true,big:true,bg:'#1B5E20',color:C.wht},
  ];

  items.forEach(function(item){
    sh.getRange(item.r,1).setValue(item.l);
    if(item.bold) sh.getRange(item.r,1).setFontWeight('bold');
    if(item.bg) sh.getRange(item.r,1,1,4).setBackground(item.bg);
    if(item.color) sh.getRange(item.r,1,1,4).setFontColor(item.color);
    if(item.big) sh.getRange(item.r,2,1,3).setFontSize(13);

    if(item.type&&!item.cat){
      // Cash basis: exclude آجلة sales
      sh.getRange(item.r,2).setFormula(
        '=SUMPRODUCT(('+e+'J$4:J$1003='+mo+')*('+e+'C$4:C$1003="'+item.type+'")*('+e+'E$4:E$1003<>"مبيعات آجلة")*'+e+'D$4:D$1003)'+(item.neg?'*-1':'')
      );
      // Accrual basis: include ALL (including آجلة)
      sh.getRange(item.r,3).setFormula(
        '=SUMPRODUCT(('+e+'J$4:J$1003='+mo+')*('+e+'C$4:C$1003="'+item.type+'")*'+e+'D$4:D$1003)'+(item.neg?'*-1':'')
      );
    }
    if(item.cat){
      // Cash basis: same category filter
      sh.getRange(item.r,2).setFormula(
        '=SUMPRODUCT(('+e+'J$4:J$1003='+mo+')*('+e+'E$4:E$1003="'+item.cat+'")*('+e+'C$4:C$1003="'+item.type+'")*'+e+'D$4:D$1003)'
      );
      // Accrual basis: same for expenses (no distinction yet)
      sh.getRange(item.r,3).setFormula(
        '=SUMPRODUCT(('+e+'J$4:J$1003='+mo+')*('+e+'E$4:E$1003="'+item.cat+'")*('+e+'C$4:C$1003="'+item.type+'")*'+e+'D$4:D$1003)'
      );
    }
    if(item.formula){
      sh.getRange(item.r,2).setFormula(item.formula);
      // Accrual totals mirror the structure but from col C
      sh.getRange(item.r,3).setFormula(item.formula.replace(/B/g,'C'));
    }

    sh.getRange(item.r,4).setFormula('=C'+item.r+'-B'+item.r);
    [2,3,4].forEach(function(c){sh.getRange(item.r,c).setNumberFormat('#,##0.00');});
  });

  [200,130,130,100].forEach(function(w,i){sh.setColumnWidth(i+1,w);});
  sh.setFrozenRows(3);sh.setTabColor('#1B5E20');
}

// ═══ DASHBOARD ═══
function buildDashboard_(ss){
  var sh=getOrCreate_(ss,SN.dash);sh.clear();sh.setRightToLeft(true);
  sh.getRange(1,1).setValue('⚡ Qayd — لوحة التحكم').setFontSize(18).setFontWeight('bold');
  sh.getRange(1,1,1,8).merge().setBackground(C.pri).setFontColor(C.wht).setHorizontalAlignment('center');

  var e="'"+SN.entry+"'!";

  // ═ Section 1: Live Balances ═
  sh.getRange(3,1).setValue('💰 الأرصدة اللحظية').setFontSize(14).setFontWeight('bold');
  sh.getRange(3,1,1,8).merge().setBackground(C.acc).setFontColor(C.wht).setHorizontalAlignment('center');

  var bals=[['الصندوق','صندوق'],['بنك الراجحي','بنك الراجحي'],['بنك الأهلي','بنك الأهلي']];
  for(var i=0;i<bals.length;i++){
    var r=4+i;
    sh.getRange(r,1).setValue(bals[i][0]).setFontWeight('bold').setFontSize(12);
    sh.getRange(r,2).setFormula(
      '=SUMIFS('+e+'D$4:D$1003,'+e+'F$4:F$1003,"'+bals[i][1]+'",'+e+'C$4:C$1003,"وارد")-SUMIFS('+e+'D$4:D$1003,'+e+'F$4:F$1003,"'+bals[i][1]+'",'+e+'C$4:C$1003,"صادر")'
    ).setFontSize(14).setFontWeight('bold').setNumberFormat('#,##0.00');
  }
  sh.getRange(7,1).setValue('📊 الرصيد الكلي').setFontWeight('bold').setFontSize(13).setBackground(C.lYlw);
  sh.getRange(7,2).setFormula('=SUM(B4:B6)').setFontWeight('bold').setFontSize(16).setNumberFormat('#,##0.00').setBackground(C.lYlw);

  // ═ Section 2: Quick KPIs ═
  sh.getRange(9,1).setValue('📈 مؤشرات الأداء').setFontSize(14).setFontWeight('bold');
  sh.getRange(9,1,1,8).merge().setBackground('#1565C0').setFontColor(C.wht).setHorizontalAlignment('center');

  var kpis=[
    ['إجمالي الإيرادات','=SUMIF('+e+'C$4:C$1003,"وارد",'+e+'D$4:D$1003)'],
    ['إجمالي المصروفات','=SUMIF('+e+'C$4:C$1003,"صادر",'+e+'D$4:D$1003)'],
    ['صافي الدخل','=B10-B11'],
    ['عدد القيود','=COUNTA('+e+'A$4:A$1003)'],
    ['متوسط قيمة القيد','=IF(B13>0,ROUND((B10+B11)/B13,0),0)'],
  ];
  for(var k=0;k<kpis.length;k++){
    var kr=10+k;
    sh.getRange(kr,1).setValue(kpis[k][0]).setFontWeight('bold');
    sh.getRange(kr,2).setFormula(kpis[k][1]).setFontWeight('bold').setNumberFormat('#,##0.00');
  }
  // Conditional formatting for net income: green if positive, red if negative
  var netIncomeRange=sh.getRange('B12');
  var rulePos=SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThan(0).setBackground(C.lGrn).setRanges([netIncomeRange]).build();
  var ruleNeg=SpreadsheetApp.newConditionalFormatRule().whenNumberLessThan(0).setBackground(C.lRed).setRanges([netIncomeRange]).build();
  sh.setConditionalFormatRules([rulePos,ruleNeg]);

  // ═ Section 3: Smart Alerts ═
  sh.getRange(16,1).setValue('🚨 تنبيهات ذكية').setFontSize(14).setFontWeight('bold');
  sh.getRange(16,1,1,8).merge().setBackground(C.dan).setFontColor(C.wht).setHorizontalAlignment('center');

  sh.getRange(17,1).setValue('رصيد الصندوق').setFontWeight('bold');
  sh.getRange(17,2).setFormula('=IF(B4<0,"⚠️ عجز في الصندوق!","✅ طبيعي")');
  sh.getRange(18,1).setValue('عهد مفتوحة').setFontWeight('bold');
  sh.getRange(18,2).setFormula("=IF(COUNTIF('"+SN.custody+"'!E:E,\"⚠️*\")>0,\"⚠️ يوجد عهد غير مسواة\",\"✅ لا توجد\")");
  sh.getRange(19,1).setValue('مشاريع خاسرة').setFontWeight('bold');
  sh.getRange(19,2).setFormula("=IF(COUNTIF('"+SN.projects+"'!H:H,\"❌*\")>0,\"⚠️ يوجد مشاريع خاسرة\",\"✅ كل المشاريع رابحة\")");

  // ═ Section 4: Monthly Comparison ═
  sh.getRange(21,1).setValue('📊 مقارنة شهرية').setFontSize(14).setFontWeight('bold');
  sh.getRange(21,1,1,8).merge().setBackground('#6A1B9A').setFontColor(C.wht).setHorizontalAlignment('center');
  var months=['يناير','فبراير','مارس','أبريل','مايو','يونيو','يوليو','أغسطس','سبتمبر','أكتوبر','نوفمبر','ديسمبر'];
  sh.getRange(22,1).setValue('الشهر').setFontWeight('bold');
  sh.getRange(22,2).setValue('إيرادات').setFontWeight('bold');
  sh.getRange(22,3).setValue('مصروفات').setFontWeight('bold');
  sh.getRange(22,4).setValue('صافي').setFontWeight('bold');
  for(var m=0;m<12;m++){
    var mr=23+m;
    sh.getRange(mr,1).setValue(months[m]);
    sh.getRange(mr,2).setFormula('=SUMPRODUCT(('+e+'J$4:J$1003="'+months[m]+'")*('+e+'C$4:C$1003="وارد")*'+e+'D$4:D$1003)').setNumberFormat('#,##0');
    sh.getRange(mr,3).setFormula('=SUMPRODUCT(('+e+'J$4:J$1003="'+months[m]+'")*('+e+'C$4:C$1003="صادر")*'+e+'D$4:D$1003)').setNumberFormat('#,##0');
    sh.getRange(mr,4).setFormula('=B'+mr+'-C'+mr).setNumberFormat('#,##0');
  }

  [200,150,150,120].forEach(function(w,i){sh.setColumnWidth(i+1,w);});
  sh.setFrozenRows(2);sh.setTabColor(C.pri);
}

// ═══ PROTECTION ═══
function protectQayd(){
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var me=Session.getEffectiveUser();
  var cnt=0;
  ss.getSheets().forEach(function(sh){
    var name=sh.getName();
    if(name===SN.entry||name===SN.settings) return;
    sh.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(function(p){p.remove();});
    var p=sh.protect().setDescription('للقراءة فقط — '+name);
    p.addEditor(me);
    p.getEditors().forEach(function(e){if(e.getEmail()!==me.getEmail())p.removeEditor(e);});
    cnt++;
  });
  // Protect formulas in entry sheet
  var entry=ss.getSheetByName(SN.entry);
  if(entry){
    var p=entry.protect().setDescription('حماية معادلات الإدخال');
    p.addEditor(me);
    var unp=[];
    [1,2,3,4,5,6,7,8].forEach(function(c){unp.push(entry.getRange(4,c,1000,1));});
    p.setUnprotectedRanges(unp);
    p.getEditors().forEach(function(e){if(e.getEmail()!==me.getEmail())p.removeEditor(e);});
  }
  SpreadsheetApp.getUi().alert('✅ تم حماية '+cnt+' شيت + معادلات سجل القيود');
}

