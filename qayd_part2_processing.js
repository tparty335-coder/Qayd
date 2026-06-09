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
