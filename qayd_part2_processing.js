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
  // Running balance — O(n) cumulative, handles وارد/مردود/صادر/تحويل
  sh.getRange(3,8).setFormula('=IF(A3="","",IF(OR(C3="وارد",C3="مردود"),D3,-D3))');
  batch_(sh,4,502,8,function(r){
    return '=IF(A'+r+'="","",H'+(r-1)+'+IF(OR(C'+r+'="وارد",C'+r+'="مردود"),D'+r+',-D'+r+'))';
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
  sh.getRange(3,8).setFormula('=IF(A3="","",IF(OR(C3="وارد",C3="مردود"),D3,-D3))');
  batch_(sh,4,502,8,function(r){
    return '=IF(A'+r+'="","",H'+(r-1)+'+IF(OR(C'+r+'="وارد",C'+r+'="مردود"),D'+r+',-D'+r+'))';
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
  var projRange=ss.getRangeByName('rng_Projects');
  var projs=projRange?projRange.getValues().map(function(r){return r[0];}).filter(function(v){return v!=='';}):[];
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
  // Total row
  if(projs.length>0){
    var tr=3+projs.length;
    sh.getRange(tr,1).setValue('الإجمالي').setFontWeight('bold').setBackground(C.lYlw);
    sh.getRange(tr,2).setFormula('=SUM(B3:B'+(tr-1)+')').setFontWeight('bold').setNumberFormat('#,##0.00');
    sh.getRange(tr,3).setFormula('=SUM(C3:C'+(tr-1)+')').setFontWeight('bold').setNumberFormat('#,##0.00');
    sh.getRange(tr,4).setFormula('=B'+tr+'-C'+tr).setFontWeight('bold').setNumberFormat('#,##0.00').setFontSize(12);
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
  var empRange=ss.getRangeByName('rng_Employees');
  var emps=empRange?empRange.getValues().map(function(r){return r[0];}).filter(function(v){return v!=='';}):[];
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
  var entRange=ss.getRangeByName('rng_Entities');
  var ents=entRange?entRange.getValues().map(function(r){return r[0];}).filter(function(v){return v!=='';}):[];
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
