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

  sh.setColumnWidth(1,50);sh.setColumnWidth(2,220);sh.setColumnWidth(3,160);
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
    {l:'💰 الصندوق',f:"=SUMIFS(D4:D1003,F4:F1003,\"صندوق\",C4:C1003,\"وارد\")+SUMIFS(D4:D1003,F4:F1003,\"صندوق\",C4:C1003,\"مردود\")-SUMIFS(D4:D1003,F4:F1003,\"صندوق\",C4:C1003,\"صادر\")-SUMIFS(D4:D1003,F4:F1003,\"صندوق\",C4:C1003,\"تحويل\")"},
    {l:'🏦 الراجحي',f:"=SUMIFS(D4:D1003,F4:F1003,\"بنك الراجحي\",C4:C1003,\"وارد\")+SUMIFS(D4:D1003,F4:F1003,\"بنك الراجحي\",C4:C1003,\"مردود\")-SUMIFS(D4:D1003,F4:F1003,\"بنك الراجحي\",C4:C1003,\"صادر\")-SUMIFS(D4:D1003,F4:F1003,\"بنك الراجحي\",C4:C1003,\"تحويل\")"},
    {l:'🏦 الأهلي',f:"=SUMIFS(D4:D1003,F4:F1003,\"بنك الأهلي\",C4:C1003,\"وارد\")+SUMIFS(D4:D1003,F4:F1003,\"بنك الأهلي\",C4:C1003,\"مردود\")-SUMIFS(D4:D1003,F4:F1003,\"بنك الأهلي\",C4:C1003,\"صادر\")-SUMIFS(D4:D1003,F4:F1003,\"بنك الأهلي\",C4:C1003,\"تحويل\")"},
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

  // ═══ SIDE SUMMARY — Visible beside entry area (Essam's style) ═══
  // Placed in columns L-O (12-15), always visible without scrolling
  var sc=12; // start column (L)
  sh.getRange(2,sc).setValue('📋 ملخص لحظي').setFontSize(12).setFontWeight('bold');
  sh.getRange(2,sc,1,4).merge().setBackground(C.pri).setFontColor(C.wht).setHorizontalAlignment('center');

  sh.getRange(3,sc,1,4).setValues([['التصنيف','صادر','وارد','الصافي']]).setFontWeight('bold').setBackground(C.hdr);
  sh.setColumnWidth(sc,180);sh.setColumnWidth(sc+1,100);sh.setColumnWidth(sc+2,100);sh.setColumnWidth(sc+3,100);

  var cats=['مصروفات إدارية عامة','محروقات (بنزين+زيت+ديزل)','وجبات ومشروبات','مشتريات عامة',
    'رواتب ويوميات وسلف','صيانة سيارات','إيجارات','مبيعات نقدية','مبيعات آجلة','إيرادات مشاريع'];
  for(var ci=0;ci<cats.length;ci++){
    var cr=4+ci;
    sh.getRange(cr,sc).setValue(cats[ci]);
    sh.getRange(cr,sc+1).setFormula('=SUMIFS(D$4:D$1003,E$4:E$1003,"'+cats[ci]+'",C$4:C$1003,"صادر")').setNumberFormat('#,##0.00');
    sh.getRange(cr,sc+2).setFormula('=SUMIFS(D$4:D$1003,E$4:E$1003,"'+cats[ci]+'",C$4:C$1003,"وارد")').setNumberFormat('#,##0.00');
    sh.getRange(cr,sc+3).setFormula('=N'+cr+'-M'+cr).setNumberFormat('#,##0.00');
  }
  var totR=4+cats.length;
  sh.getRange(totR,sc).setValue('الإجمالي الكلي').setFontWeight('bold').setBackground(C.lYlw);
  sh.getRange(totR,sc+1).setFormula('=SUM(M4:M'+(totR-1)+')').setFontWeight('bold').setNumberFormat('#,##0.00');
  sh.getRange(totR,sc+2).setFormula('=SUM(N4:N'+(totR-1)+')').setFontWeight('bold').setNumberFormat('#,##0.00');
  sh.getRange(totR,sc+3).setFormula('=N'+totR+'-M'+totR).setFontWeight('bold').setNumberFormat('#,##0.00').setFontSize(12);
}
