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
      sh.getRange(item.r,2).setFormula(
        '=SUMPRODUCT(('+e+'J$4:J$1003='+mo+')*('+e+'C$4:C$1003="'+item.type+'")*'+e+'D$4:D$1003)'+(item.neg?'*-1':'')
      );
    }
    if(item.cat){
      sh.getRange(item.r,2).setFormula(
        '=SUMPRODUCT(('+e+'J$4:J$1003='+mo+')*('+e+'E$4:E$1003="'+item.cat+'")*('+e+'C$4:C$1003="'+item.type+'")*'+e+'D$4:D$1003)'
      );
    }
    if(item.formula) sh.getRange(item.r,2).setFormula(item.formula);

    // Accrual column = same for now (can be customized)
    sh.getRange(item.r,3).setFormula('=B'+item.r);
    sh.getRange(item.r,4).setFormula('=B'+item.r+'-C'+item.r);
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
  sh.getRange(12,2).setFontSize(16).setBackground(sh.getRange(12,2).getFormula().indexOf('-')>-1?C.lRed:C.lGrn);

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
