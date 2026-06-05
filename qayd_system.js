// â•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گ
// Qayd â€” ظ‚ظٹط¯ ظˆط§ط­ط¯ â†’ ظٹط¸ظ‡ط± ظپظٹ ظƒظ„ ظ…ظƒط§ظ†
// Part 1: Core + Settings + Entry Log
// â•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گ

var C = {pri:'#1B3A5C',acc:'#E8A838',suc:'#2E7D32',dan:'#C62828',wht:'#FFFFFF',drk:'#263238',bg:'#F5F5F5',
  hdr:'#BDD6EE',hdr2:'#FFF2CC',hdr3:'#D9D2E9',lGrn:'#E8F5E9',lRed:'#FFEBEE',lYlw:'#FFF9C4'};

var SN = {
  entry:'ط³ط¬ظ„ ط§ظ„ظ‚ظٹظˆط¯',settings:'ط§ظ„ط¥ط¹ط¯ط§ط¯ط§طھ',
  cash:'ط§ظ„طµظ†ط¯ظˆظ‚',rajhi:'ط¨ظ†ظƒ ط§ظ„ط±ط§ط¬ط­ظٹ',ahli:'ط¨ظ†ظƒ ط§ظ„ط£ظ‡ظ„ظٹ',
  projects:'ط§ظ„ظ…ط´ط§ط±ظٹط¹',custody:'ط§ظ„ط¹ظ‡ط¯',interco:'ط§ظ„طھظ‚ط§طµ',
  incomeW:'ظ‚ط§ط¦ظ…ط© ط§ظ„ط¯ط®ظ„ ط§ظ„ط£ط³ط¨ظˆط¹ظٹط©',incomeM:'ظ‚ط§ط¦ظ…ط© ط§ظ„ط¯ط®ظ„ ط§ظ„ط´ظ‡ط±ظٹط©',
  dash:'ظ„ظˆط­ط© ط§ظ„طھط­ظƒظ…',reports:'ط§ظ„طھظ‚ط§ط±ظٹط±'
};

// â•گâ•گâ•گ HELPERS â•گâ•گâ•گ
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

// â•گâ•گâ•گ MENU â•گâ•گâ•گ
function onOpen(){
  SpreadsheetApp.getUi().createMenu('âڑ، Qayd')
    .addItem('ًں”§ ط¨ظ†ط§ط، ط§ظ„ظ†ط¸ط§ظ…','setupQayd')
    .addItem('ًں”’ ط­ظ…ط§ظٹط© ط§ظ„ظ…ط¹ط§ط¯ظ„ط§طھ','protectQayd')
    .addToUi();
}

// â•گâ•گâ•گ MAIN SETUP â•گâ•گâ•گ
function setupQayd(){
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var ui=SpreadsheetApp.getUi();
  var r=ui.alert('âڑ، Qayd','ط³ظٹطھظ… ط¨ظ†ط§ط، ط§ظ„ظ†ط¸ط§ظ… ط§ظ„ظƒط§ظ…ظ„. ظ‡ظ„ طھط±ظٹط¯ ط§ظ„ظ…طھط§ط¨ط¹ط©طں',ui.ButtonSet.YES_NO);
  if(r!==ui.Button.YES)return;

  buildSettings_(ss);SpreadsheetApp.flush();
  buildEntryLog_(ss);SpreadsheetApp.flush();
  buildCashSheet_(ss);SpreadsheetApp.flush();
  buildBankSheet_(ss,SN.rajhi,'ط§ظ„ط±ط§ط¬ط­ظٹ');SpreadsheetApp.flush();
  buildBankSheet_(ss,SN.ahli,'ط§ظ„ط£ظ‡ظ„ظٹ');SpreadsheetApp.flush();
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
  ui.alert('âœ… طھظ… ط¨ظ†ط§ط، Qayd ط¨ظ†ط¬ط§ط­!');
}

// â•گâ•گâ•گ SETTINGS â•گâ•گâ•گ
function buildSettings_(ss){
  var sh=getOrCreate_(ss,SN.settings);
  sh.clear();sh.setRightToLeft(true);

  sh.getRange(1,1).setValue('ط¥ط¹ط¯ط§ط¯ط§طھ Qayd').setFontSize(16).setFontWeight('bold');
  sh.getRange(1,1,1,4).merge().setBackground(C.pri).setFontColor(C.wht).setHorizontalAlignment('center');

  // Company info
  sh.getRange(3,1).setValue('ط§ط³ظ… ط§ظ„ط´ط±ظƒط©:').setFontWeight('bold');
  sh.getRange(3,2).setValue('ط´ط±ظƒط© ط¥ط¹ظ…ط§ط± ظ„ط§ظ†ط¯');
  sh.getRange(4,1).setValue('ط§ظ„ط³ظ†ط© ط§ظ„ظ…ط§ظ„ظٹط©:').setFontWeight('bold');
  sh.getRange(4,2).setValue(new Date().getFullYear());

  // Expense categories
  sh.getRange(6,1).setValue('طھطµظ†ظٹظپط§طھ ط§ظ„ظ…طµط±ظˆظپط§طھ').setFontWeight('bold').setBackground(C.hdr);
  sh.getRange(6,1,1,2).merge();
  var cats=['ظ…طµط±ظˆظپط§طھ ط¥ط¯ط§ط±ظٹط© ط¹ط§ظ…ط©','ظ…ط­ط±ظˆظ‚ط§طھ (ط¨ظ†ط²ظٹظ†+ط²ظٹطھ+ط¯ظٹط²ظ„)','ظˆط¬ط¨ط§طھ ظˆظ…ط´ط±ظˆط¨ط§طھ','ظ…ط´طھط±ظٹط§طھ ط¹ط§ظ…ط©',
    'ط±ظˆط§طھط¨ ظˆظٹظˆظ…ظٹط§طھ ظˆط³ظ„ظپ','طµظٹط§ظ†ط© ط³ظٹط§ط±ط§طھ','ط£طµظˆظ„ ط«ط§ط¨طھط©','ط¥ظٹط¬ط§ط±ط§طھ','ظƒظ‡ط±ط¨ط§ط، ظˆظ…ط§ط،','ط§طھطµط§ظ„ط§طھ ظˆط¥ظ†طھط±ظ†طھ',
    'ظ…طµط±ظˆظپط§طھ طھط³ظˆظٹظ‚ظٹط©','ط±ط³ظˆظ… ط­ظƒظˆظ…ظٹط©','ظ…طµط±ظˆظپط§طھ ظ†ظ‚ظ„','ظ…طµط±ظˆظپط§طھ ظ…طھظ†ظˆط¹ط©'];
  for(var i=0;i<cats.length;i++) sh.getRange(7+i,1).setValue(i+1), sh.getRange(7+i,2).setValue(cats[i]);

  // Revenue categories
  var rr=7+cats.length+1;
  sh.getRange(rr,1).setValue('طھطµظ†ظٹظپط§طھ ط§ظ„ط¥ظٹط±ط§ط¯ط§طھ').setFontWeight('bold').setBackground(C.lGrn);
  sh.getRange(rr,1,1,2).merge();
  var revs=['ظ…ط¨ظٹط¹ط§طھ ظ†ظ‚ط¯ظٹط©','ظ…ط¨ظٹط¹ط§طھ ط¢ط¬ظ„ط©','ط¥ظٹط±ط§ط¯ط§طھ ظ…ط´ط§ط±ظٹط¹','ط¥ظٹط±ط§ط¯ط§طھ ط®ط¯ظ…ط§طھ','ط¥ظٹط±ط§ط¯ط§طھ ط£ط®ط±ظ‰'];
  for(var j=0;j<revs.length;j++) sh.getRange(rr+1+j,1).setValue(j+1), sh.getRange(rr+1+j,2).setValue(revs[j]);

  // Projects
  var pr=rr+revs.length+2;
  sh.getRange(pr,1).setValue('ط§ظ„ظ…ط´ط§ط±ظٹط¹').setFontWeight('bold').setBackground(C.hdr2);
  sh.getRange(pr,1,1,3).merge();
  sh.getRange(pr+1,1,1,3).setValues([['ظƒظˆط¯','ط§ط³ظ… ط§ظ„ظ…ط´ط±ظˆط¹','ط§ظ„ط¹ظ…ظٹظ„']]).setFontWeight('bold');
  var projs=[['P01','ظ…ط´ط±ظˆط¹ ط§ظ„ط³ط¬ظˆظ†','ظˆط²ط§ط±ط© ط§ظ„ط¯ط§ط®ظ„ظٹط©'],['P02','ظ…ط´ط±ظˆط¹ ط§ظ„ظ…ط²ط§ط­ظ…ظٹط©','ط§ظ„ظ‚ط·ط§ط¹ ط§ظ„ط®ط§طµ'],
    ['P03','ظپظ„ظ„ ط¬ط¯ط©','ط´ط±ظƒط© ط§ظ„ط¨ظ†ط§ط،'],['P04','ظ…ط´ط±ظˆط¹ ط§ظ„ط±ظ…ط§ظ„','ط¯ط±ط± ط§ظ„ط¹ظˆط§طµظ…']];
  sh.getRange(pr+2,1,projs.length,3).setValues(projs);

  // Employees
  var er=pr+projs.length+3;
  sh.getRange(er,1).setValue('ط§ظ„ظ…ظˆط¸ظپظˆظ†').setFontWeight('bold').setBackground(C.hdr3);
  sh.getRange(er,1,1,2).merge();
  var emps=['ط³ط¹ط¯ظٹط©','ط·ط§ط±ظ‚','ظ…ط­ظ…ط¯ ظپظ‡ظٹظ…','ط£ط­ظ…ط¯','ط¹ط¨ط¯ط§ظ„ظ„ظ‡'];
  for(var e=0;e<emps.length;e++) sh.getRange(er+1+e,1).setValue(e+1), sh.getRange(er+1+e,2).setValue(emps[e]);

  // Intercompany entities
  var ir=er+emps.length+2;
  sh.getRange(ir,1).setValue('ظƒظٹط§ظ†ط§طھ ط§ظ„طھظ‚ط§طµ').setFontWeight('bold').setBackground(C.acc).setFontColor(C.wht);
  sh.getRange(ir,1,1,2).merge();
  var ents=['ط¥ط¹ظ…ط§ط±','ظ…ط³ط§ط±','ط§ظ„ظپط±ط¹ ط§ظ„ط±ظٹط§ط¶'];
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

// â•گâ•گâ•گ ENTRY LOG â€” The Single Input Sheet â•گâ•گâ•گ
// This is the ONLY sheet Essam types into. Everything else auto-populates.
function buildEntryLog_(ss){
  var sh=getOrCreate_(ss,SN.entry);
  sh.clear();
  if(sh.getMaxRows()<1010) sh.insertRowsAfter(sh.getMaxRows(),1010-sh.getMaxRows());
  sh.getRange(1,1,sh.getMaxRows(),sh.getMaxColumns()).clearDataValidations();
  sh.setRightToLeft(true);

  // Title
  sh.getRange(1,1).setValue('âڑ، Qayd â€” ط³ط¬ظ„ ط§ظ„ظ‚ظٹظˆط¯ ط§ظ„ظٹظˆظ…ظٹ').setFontSize(16).setFontWeight('bold');
  sh.getRange(1,1,1,10).merge().setBackground(C.pri).setFontColor(C.wht).setHorizontalAlignment('center');

  // Live balance bar (row 2) â€” instant feedback while entering
  var balItems=[
    {l:'ًں’° ط§ظ„طµظ†ط¯ظˆظ‚',f:"=SUMIFS(D4:D1003,F4:F1003,\"طµظ†ط¯ظˆظ‚\",C4:C1003,\"ظˆط§ط±ط¯\")-SUMIFS(D4:D1003,F4:F1003,\"طµظ†ط¯ظˆظ‚\",C4:C1003,\"طµط§ط¯ط±\")"},
    {l:'ًںڈ¦ ط§ظ„ط±ط§ط¬ط­ظٹ',f:"=SUMIFS(D4:D1003,F4:F1003,\"ط¨ظ†ظƒ ط§ظ„ط±ط§ط¬ط­ظٹ\",C4:C1003,\"ظˆط§ط±ط¯\")-SUMIFS(D4:D1003,F4:F1003,\"ط¨ظ†ظƒ ط§ظ„ط±ط§ط¬ط­ظٹ\",C4:C1003,\"طµط§ط¯ط±\")"},
    {l:'ًںڈ¦ ط§ظ„ط£ظ‡ظ„ظٹ',f:"=SUMIFS(D4:D1003,F4:F1003,\"ط¨ظ†ظƒ ط§ظ„ط£ظ‡ظ„ظٹ\",C4:C1003,\"ظˆط§ط±ط¯\")-SUMIFS(D4:D1003,F4:F1003,\"ط¨ظ†ظƒ ط§ظ„ط£ظ‡ظ„ظٹ\",C4:C1003,\"طµط§ط¯ط±\")"},
    {l:'ًں“ٹ ط§ظ„ط¥ط¬ظ…ط§ظ„ظٹ',f:"=B2+D2+F2"}
  ];
  for(var b=0;b<balItems.length;b++){
    var bc=1+b*2;
    sh.getRange(2,bc).setValue(balItems[b].l).setFontWeight('bold').setFontSize(11).setBackground(C.lYlw);
    sh.getRange(2,bc+1).setFormula(balItems[b].f).setFontWeight('bold').setFontSize(13).setNumberFormat('#,##0.00').setBackground(C.lYlw);
  }
  // Alert indicator
  sh.getRange(2,9).setValue('âڑ،').setFontSize(14).setBackground(C.acc).setFontColor(C.wht).setHorizontalAlignment('center');
  sh.getRange(2,10).setFormula('=IF(H2<0,"âڑ ï¸ڈ ط¹ط¬ط²!","âœ… ظ…طھظˆط§ط²ظ†")').setFontWeight('bold').setFontSize(11).setBackground(C.acc).setFontColor(C.wht);

  // Column headers (row 3) â€” only 8 input columns
  var headers=['ط§ظ„طھط§ط±ظٹط®','ط§ظ„ط¨ظٹط§ظ†','ظ†ظˆط¹ ط§ظ„ط­ط±ظƒط©','ط§ظ„ظ…ط¨ظ„ط؛','ط§ظ„طھطµظ†ظٹظپ','ط§ظ„ظ…طµط¯ط±','ط§ظ„ظ…ط´ط±ظˆط¹','ط§ظ„ط·ط±ظپ ط§ظ„ظ…ظ‚ط§ط¨ظ„','#','ط§ظ„ط´ظ‡ط±'];
  sh.getRange(3,1,1,10).setValues([headers]);
  fmtH_(sh,3,10,C.drk);
  sh.setColumnWidths(1,10,[110,250,80,120,180,120,130,150,40,80]);

  // Data validation dropdowns
  var DR=1000;
  addDV_(sh,4,3,DR,['ظˆط§ط±ط¯','طµط§ط¯ط±','ظ…ط±ط¯ظˆط¯','طھط­ظˆظٹظ„']);
  addDV_(sh,4,6,DR,['طµظ†ط¯ظˆظ‚','ط¨ظ†ظƒ ط§ظ„ط±ط§ط¬ط­ظٹ','ط¨ظ†ظƒ ط§ظ„ط£ظ‡ظ„ظٹ']);

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

  // â•گâ•گâ•گ EMBEDDED SUMMARY â€” Below data (like Essam's style) â•گâ•گâ•گ
  var sumRow=4+DR+2;
  sh.getRange(sumRow,1).setValue('ًں“‹ ظ…ظ„ط®طµ ط³ط±ظٹط¹').setFontSize(14).setFontWeight('bold');
  sh.getRange(sumRow,1,1,8).merge().setBackground(C.pri).setFontColor(C.wht).setHorizontalAlignment('center');

  var sumHeaders=['ط§ظ„طھطµظ†ظٹظپ','ط¥ط¬ظ…ط§ظ„ظٹ طµط§ط¯ط±','ط¥ط¬ظ…ط§ظ„ظٹ ظˆط§ط±ط¯','ط§ظ„طµط§ظپظٹ'];
  sh.getRange(sumRow+1,1,1,4).setValues([sumHeaders]).setFontWeight('bold').setBackground(C.hdr);

  // Auto summary per category
  var cats=['ظ…طµط±ظˆظپط§طھ ط¥ط¯ط§ط±ظٹط© ط¹ط§ظ…ط©','ظ…ط­ط±ظˆظ‚ط§طھ (ط¨ظ†ط²ظٹظ†+ط²ظٹطھ+ط¯ظٹط²ظ„)','ظˆط¬ط¨ط§طھ ظˆظ…ط´ط±ظˆط¨ط§طھ','ظ…ط´طھط±ظٹط§طھ ط¹ط§ظ…ط©',
    'ط±ظˆط§طھط¨ ظˆظٹظˆظ…ظٹط§طھ ظˆط³ظ„ظپ','طµظٹط§ظ†ط© ط³ظٹط§ط±ط§طھ','ط¥ظٹط¬ط§ط±ط§طھ','ظ…ط¨ظٹط¹ط§طھ ظ†ظ‚ط¯ظٹط©','ظ…ط¨ظٹط¹ط§طھ ط¢ط¬ظ„ط©','ط¥ظٹط±ط§ط¯ط§طھ ظ…ط´ط§ط±ظٹط¹'];
  for(var ci=0;ci<cats.length;ci++){
    var cr=sumRow+2+ci;
    sh.getRange(cr,1).setValue(cats[ci]);
    sh.getRange(cr,2).setFormula('=SUMIFS(D$4:D$1003,E$4:E$1003,"'+cats[ci]+'",C$4:C$1003,"طµط§ط¯ط±")').setNumberFormat('#,##0.00');
    sh.getRange(cr,3).setFormula('=SUMIFS(D$4:D$1003,E$4:E$1003,"'+cats[ci]+'",C$4:C$1003,"ظˆط§ط±ط¯")').setNumberFormat('#,##0.00');
    sh.getRange(cr,4).setFormula('=C'+cr+'-B'+cr).setNumberFormat('#,##0.00');
  }
  var totR=sumRow+2+cats.length;
  sh.getRange(totR,1).setValue('ط§ظ„ط¥ط¬ظ…ط§ظ„ظٹ ط§ظ„ظƒظ„ظٹ').setFontWeight('bold').setBackground(C.lYlw);
  sh.getRange(totR,2).setFormula('=SUM(B'+(sumRow+2)+':B'+(totR-1)+')').setFontWeight('bold').setNumberFormat('#,##0.00');
  sh.getRange(totR,3).setFormula('=SUM(C'+(sumRow+2)+':C'+(totR-1)+')').setFontWeight('bold').setNumberFormat('#,##0.00');
  sh.getRange(totR,4).setFormula('=C'+totR+'-B'+totR).setFontWeight('bold').setNumberFormat('#,##0.00').setFontSize(12);
}
// â•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گ
// Qayd Part 2: Auto-Processing Sheets (Hidden)
// All fed automatically from ط³ط¬ظ„ ط§ظ„ظ‚ظٹظˆط¯ via FILTER
// â•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گ

// â•گâ•گâ•گ CASH SHEET (auto from entry log where source=طµظ†ط¯ظˆظ‚) â•گâ•گâ•گ
function buildCashSheet_(ss){
  var sh=getOrCreate_(ss,SN.cash);sh.clear();sh.setRightToLeft(true);
  sh.getRange(1,1).setValue('ط­ط±ظƒط§طھ ط§ظ„طµظ†ط¯ظˆظ‚ (طھظ„ظ‚ط§ط¦ظٹ)').setFontSize(14).setFontWeight('bold');
  sh.getRange(1,1,1,8).merge().setBackground(C.pri).setFontColor(C.wht).setHorizontalAlignment('center');
  sh.getRange(2,1,1,8).setValues([['ط§ظ„طھط§ط±ظٹط®','ط§ظ„ط¨ظٹط§ظ†','ظ†ظˆط¹','ط§ظ„ظ…ط¨ظ„ط؛','ط§ظ„طھطµظ†ظٹظپ','ط§ظ„ظ…ط´ط±ظˆط¹','ط§ظ„ط·ط±ظپ','ط§ظ„ط±طµظٹط¯']]);
  fmtH_(sh,2,8,C.drk);
  // FILTER from entry log
  var e="'"+SN.entry+"'!";
  sh.getRange(3,1).setFormula(
    '=IFERROR(QUERY(FILTER({'+e+'A4:A1003,'+e+'B4:B1003,'+e+'C4:C1003,'+e+'D4:D1003,'+e+'E4:E1003,'+e+'G4:G1003,'+e+'H4:H1003},'+e+'F4:F1003="طµظ†ط¯ظˆظ‚",'+e+'A4:A1003<>""),"",0),"")'
  );
  // Running balance
  batch_(sh,3,502,8,function(r){
    return '=IF(A'+r+'="","",SUMPRODUCT((A$3:A'+r+'<>"")*IF(C$3:C'+r+'="ظˆط§ط±ط¯",D$3:D'+r+',-D$3:D'+r+')))';
  });
  sh.getRange(3,1,500,1).setNumberFormat('yyyy-mm-dd');
  sh.getRange(3,4,500,1).setNumberFormat('#,##0.00');
  sh.getRange(3,8,500,1).setNumberFormat('#,##0.00');
  sh.setFrozenRows(2);sh.setTabColor('#4CAF50');
}

// â•گâ•گâ•گ BANK SHEET (reusable for Rajhi/Ahli) â•گâ•گâ•گ
function buildBankSheet_(ss,sheetName,bankLabel){
  var sh=getOrCreate_(ss,sheetName);sh.clear();sh.setRightToLeft(true);
  sh.getRange(1,1).setValue('ط­ط±ظƒط§طھ '+bankLabel+' (طھظ„ظ‚ط§ط¦ظٹ)').setFontSize(14).setFontWeight('bold');
  sh.getRange(1,1,1,8).merge().setBackground('#1565C0').setFontColor(C.wht).setHorizontalAlignment('center');
  sh.getRange(2,1,1,8).setValues([['ط§ظ„طھط§ط±ظٹط®','ط§ظ„ط¨ظٹط§ظ†','ظ†ظˆط¹','ط§ظ„ظ…ط¨ظ„ط؛','ط§ظ„طھطµظ†ظٹظپ','ط§ظ„ظ…ط´ط±ظˆط¹','ط§ظ„ط·ط±ظپ','ط§ظ„ط±طµظٹط¯']]);
  fmtH_(sh,2,8,C.drk);
  var e="'"+SN.entry+"'!";
  sh.getRange(3,1).setFormula(
    '=IFERROR(QUERY(FILTER({'+e+'A4:A1003,'+e+'B4:B1003,'+e+'C4:C1003,'+e+'D4:D1003,'+e+'E4:E1003,'+e+'G4:G1003,'+e+'H4:H1003},'+e+'F4:F1003="'+sheetName+'",'+e+'A4:A1003<>""),"",0),"")'
  );
  batch_(sh,3,502,8,function(r){
    return '=IF(A'+r+'="","",SUMPRODUCT((A$3:A'+r+'<>"")*IF(C$3:C'+r+'="ظˆط§ط±ط¯",D$3:D'+r+',-D$3:D'+r+')))';
  });
  sh.getRange(3,1,500,1).setNumberFormat('yyyy-mm-dd');
  sh.getRange(3,4,500,1).setNumberFormat('#,##0.00');
  sh.getRange(3,8,500,1).setNumberFormat('#,##0.00');
  sh.setFrozenRows(2);sh.setTabColor('#1565C0');
}

// â•گâ•گâ•گ PROJECTS SUMMARY (auto P&L per project) â•گâ•گâ•گ
function buildProjectsSummary_(ss){
  var sh=getOrCreate_(ss,SN.projects);sh.clear();sh.setRightToLeft(true);
  sh.getRange(1,1).setValue('ط±ط¨ط­ظٹط© ط§ظ„ظ…ط´ط§ط±ظٹط¹ (طھظ„ظ‚ط§ط¦ظٹ)').setFontSize(14).setFontWeight('bold');
  sh.getRange(1,1,1,8).merge().setBackground(C.acc).setFontColor(C.wht).setHorizontalAlignment('center');
  sh.getRange(2,1,1,8).setValues([['ط§ظ„ظ…ط´ط±ظˆط¹','ط¥ظٹط±ط§ط¯ط§طھ','ظ…طµط±ظˆظپط§طھ','طµط§ظپظٹ ط±ط¨ط­/ط®ط³ط§ط±ط©','ظ†ط³ط¨ط© ط§ظ„ط±ط¨ط­ %','ط¹ط¯ط¯ ط§ظ„ط­ط±ظƒط§طھ','ط¢ط®ط± ط­ط±ظƒط©','ط­ط§ظ„ط©']]);
  fmtH_(sh,2,8,C.drk);
  var e="'"+SN.entry+"'!";
  var projs=['ظ…ط´ط±ظˆط¹ ط§ظ„ط³ط¬ظˆظ†','ظ…ط´ط±ظˆط¹ ط§ظ„ظ…ط²ط§ط­ظ…ظٹط©','ظپظ„ظ„ ط¬ط¯ط©','ظ…ط´ط±ظˆط¹ ط§ظ„ط±ظ…ط§ظ„'];
  for(var i=0;i<projs.length;i++){
    var r=3+i,p=projs[i];
    sh.getRange(r,1).setValue(p);
    sh.getRange(r,2).setFormula('=SUMIFS('+e+'D$4:D$1003,'+e+'G$4:G$1003,"'+p+'",'+e+'C$4:C$1003,"ظˆط§ط±ط¯")').setNumberFormat('#,##0.00');
    sh.getRange(r,3).setFormula('=SUMIFS('+e+'D$4:D$1003,'+e+'G$4:G$1003,"'+p+'",'+e+'C$4:C$1003,"طµط§ط¯ط±")').setNumberFormat('#,##0.00');
    sh.getRange(r,4).setFormula('=B'+r+'-C'+r).setNumberFormat('#,##0.00');
    sh.getRange(r,5).setFormula('=IF(B'+r+'>0,ROUND(D'+r+'/B'+r+'*100,1),"")');
    sh.getRange(r,6).setFormula('=COUNTIF('+e+'G$4:G$1003,"'+p+'")');
    sh.getRange(r,7).setFormula('=IFERROR(INDEX('+e+'A$4:A$1003,MATCH(2,1/('+e+'G$4:G$1003="'+p+'"),1)),"")').setNumberFormat('yyyy-mm-dd');
    sh.getRange(r,8).setFormula('=IF(D'+r+'>0,"âœ… ط±ط¨ط­",IF(D'+r+'<0,"â‌Œ ط®ط³ط§ط±ط©","âڑھ ظ…طھط¹ط§ط¯ظ„"))');
  }
  sh.setFrozenRows(2);sh.setTabColor('#FF6F00');
}

// â•گâ•گâ•گ CUSTODY (auto per employee) â•گâ•گâ•گ
function buildCustodySheet_(ss){
  var sh=getOrCreate_(ss,SN.custody);sh.clear();sh.setRightToLeft(true);
  sh.getRange(1,1).setValue('طھطھط¨ط¹ ط§ظ„ط¹ظ‡ط¯ (طھظ„ظ‚ط§ط¦ظٹ)').setFontSize(14).setFontWeight('bold');
  sh.getRange(1,1,1,5).merge().setBackground('#7B1FA2').setFontColor(C.wht).setHorizontalAlignment('center');
  sh.getRange(2,1,1,5).setValues([['ط§ظ„ظ…ظˆط¸ظپ','ط¹ظ‡ط¯ ظ…ط³ظ„ظ…ط©','ط¹ظ‡ط¯ ظ…ط³ظˆط§ط©','ط±طµظٹط¯ ط¹ظ‡ط¯ط©','ط­ط§ظ„ط©']]);
  fmtH_(sh,2,5,C.drk);
  var e="'"+SN.entry+"'!";
  var emps=['ط³ط¹ط¯ظٹط©','ط·ط§ط±ظ‚','ظ…ط­ظ…ط¯ ظپظ‡ظٹظ…','ط£ط­ظ…ط¯','ط¹ط¨ط¯ط§ظ„ظ„ظ‡'];
  for(var i=0;i<emps.length;i++){
    var r=3+i,emp=emps[i];
    sh.getRange(r,1).setValue(emp);
    sh.getRange(r,2).setFormula('=SUMIFS('+e+'D$4:D$1003,'+e+'H$4:H$1003,"'+emp+'",'+e+'C$4:C$1003,"طµط§ط¯ط±")').setNumberFormat('#,##0.00');
    sh.getRange(r,3).setFormula('=SUMIFS('+e+'D$4:D$1003,'+e+'H$4:H$1003,"'+emp+'",'+e+'C$4:C$1003,"ظˆط§ط±ط¯")').setNumberFormat('#,##0.00');
    sh.getRange(r,4).setFormula('=B'+r+'-C'+r).setNumberFormat('#,##0.00');
    sh.getRange(r,5).setFormula('=IF(D'+r+'>0,"âڑ ï¸ڈ ط¹ظ‡ط¯ط© ظ…ظپطھظˆط­ط©",IF(D'+r+'=0,"âœ… ظ…ط³ظˆظ‘ط§ط©","ًں”„ ط¯ط§ط¦ظ†"))');
  }
  sh.setTabColor('#7B1FA2');
}

// â•گâ•گâ•گ INTERCOMPANY (Emaar â†” Masar auto-reconciliation) â•گâ•گâ•گ
function buildIntercoSheet_(ss){
  var sh=getOrCreate_(ss,SN.interco);sh.clear();sh.setRightToLeft(true);
  sh.getRange(1,1).setValue('ط­ط³ط§ط¨ط§طھ ط§ظ„طھظ‚ط§طµ ط¨ظٹظ† ط§ظ„ظƒظٹط§ظ†ط§طھ (طھظ„ظ‚ط§ط¦ظٹ)').setFontSize(14).setFontWeight('bold');
  sh.getRange(1,1,1,5).merge().setBackground('#E65100').setFontColor(C.wht).setHorizontalAlignment('center');
  sh.getRange(2,1,1,5).setValues([['ط§ظ„ظƒظٹط§ظ†','ظ„ظ‡ (ظˆط§ط±ط¯)','ط¹ظ„ظٹظ‡ (طµط§ط¯ط±)','ط§ظ„ط±طµظٹط¯ ط§ظ„طµط§ظپظٹ','ط§ظ„ط§طھط¬ط§ظ‡']]);
  fmtH_(sh,2,5,C.drk);
  var e="'"+SN.entry+"'!";
  var ents=['ط¥ط¹ظ…ط§ط±','ظ…ط³ط§ط±','ط§ظ„ظپط±ط¹ ط§ظ„ط±ظٹط§ط¶'];
  for(var i=0;i<ents.length;i++){
    var r=3+i,ent=ents[i];
    sh.getRange(r,1).setValue(ent);
    sh.getRange(r,2).setFormula('=SUMIFS('+e+'D$4:D$1003,'+e+'H$4:H$1003,"'+ent+'",'+e+'C$4:C$1003,"ظˆط§ط±ط¯")').setNumberFormat('#,##0.00');
    sh.getRange(r,3).setFormula('=SUMIFS('+e+'D$4:D$1003,'+e+'H$4:H$1003,"'+ent+'",'+e+'C$4:C$1003,"طµط§ط¯ط±")').setNumberFormat('#,##0.00');
    sh.getRange(r,4).setFormula('=B'+r+'-C'+r).setNumberFormat('#,##0.00');
    sh.getRange(r,5).setFormula('=IF(D'+r+'>0,"â¬…ï¸ڈ ظ„ظ†ط§ ط¹ظ†ط¯ظ‡",IF(D'+r+'<0,"â‍،ï¸ڈ ظ„ظ‡ظ… ط¹ظ†ط¯ظ†ط§","âڑھ ظ…طھط³ط§ظˆظٹ"))');
  }
  sh.setTabColor('#E65100');
}
// â•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گ
// Qayd Part 3: Income Statements + Dashboard + Protection
// â•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گâ•گ

// â•گâ•گâ•گ WEEKLY INCOME STATEMENT â•گâ•گâ•گ
function buildWeeklyIncome_(ss){
  var sh=getOrCreate_(ss,SN.incomeW);sh.clear();sh.setRightToLeft(true);
  sh.getRange(1,1).setValue('ظ‚ط§ط¦ظ…ط© ط§ظ„ط¯ط®ظ„ ط§ظ„ط£ط³ط¨ظˆط¹ظٹط©').setFontSize(16).setFontWeight('bold');
  sh.getRange(1,1,1,6).merge().setBackground(C.suc).setFontColor(C.wht).setHorizontalAlignment('center');

  // Week selector
  sh.getRange(2,1).setValue('ظ…ظ† طھط§ط±ظٹط®:').setFontWeight('bold');
  sh.getRange(2,3).setValue('ط¥ظ„ظ‰ طھط§ط±ظٹط®:').setFontWeight('bold');
  sh.getRange(2,2).setValue(new Date()).setNumberFormat('yyyy-mm-dd').setBackground(C.lYlw);
  sh.getRange(2,4).setValue(new Date()).setNumberFormat('yyyy-mm-dd').setBackground(C.lYlw);

  var e="'"+SN.entry+"'!";
  var from='$B$2',to='$D$2';

  // Income Statement Structure
  var rows=[
    {r:4,l:'ط§ظ„ظ…ط¨ظٹط¹ط§طھ ظˆط§ظ„ط¥ظٹط±ط§ط¯ط§طھ',bg:C.hdr,bold:true,isHeader:true},
    {r:5,l:'ظ…ط¨ظٹط¹ط§طھ ظ†ظ‚ط¯ظٹط©',cat:'ظ…ط¨ظٹط¹ط§طھ ظ†ظ‚ط¯ظٹط©',type:'ظˆط§ط±ط¯'},
    {r:6,l:'ظ…ط¨ظٹط¹ط§طھ ط¢ط¬ظ„ط©',cat:'ظ…ط¨ظٹط¹ط§طھ ط¢ط¬ظ„ط©',type:'ظˆط§ط±ط¯'},
    {r:7,l:'ط¥ظٹط±ط§ط¯ط§طھ ظ…ط´ط§ط±ظٹط¹',cat:'ط¥ظٹط±ط§ط¯ط§طھ ظ…ط´ط§ط±ظٹط¹',type:'ظˆط§ط±ط¯'},
    {r:8,l:'ط¥ظٹط±ط§ط¯ط§طھ ط®ط¯ظ…ط§طھ',cat:'ط¥ظٹط±ط§ط¯ط§طھ ط®ط¯ظ…ط§طھ',type:'ظˆط§ط±ط¯'},
    {r:9,l:'ظ…ط±ط¯ظˆط¯ط§طھ ط§ظ„ظ…ط¨ظٹط¹ط§طھ',cat:'ظ…ط¨ظٹط¹ط§طھ ظ†ظ‚ط¯ظٹط©',type:'ظ…ط±ط¯ظˆط¯',neg:true},
    {r:10,l:'طµط§ظپظٹ ط§ظ„ظ…ط¨ظٹط¹ط§طھ',bg:C.lGrn,bold:true,sum:'B5:B8',sub:'B9'},
    {r:12,l:'ط§ظ„ظ…ط´طھط±ظٹط§طھ',bg:C.hdr2,bold:true,isHeader:true},
    {r:13,l:'ظ…ط´طھط±ظٹط§طھ ط¹ط§ظ…ط©',cat:'ظ…ط´طھط±ظٹط§طھ ط¹ط§ظ…ط©',type:'طµط§ط¯ط±'},
    {r:14,l:'ظ…ط±ط¯ظˆط¯ط§طھ ط§ظ„ظ…ط´طھط±ظٹط§طھ',cat:'ظ…ط´طھط±ظٹط§طھ ط¹ط§ظ…ط©',type:'ظ…ط±ط¯ظˆط¯'},
    {r:15,l:'طµط§ظپظٹ ط§ظ„ظ…ط´طھط±ظٹط§طھ',bg:C.lYlw,bold:true,sum:'B13',sub:'B14'},
    {r:17,l:'ظ…ط¬ظ…ظ„ ط§ظ„ط±ط¨ط­',bg:'#C8E6C9',bold:true,big:true,formula:'=B10-B15'},
    {r:19,l:'ط§ظ„ظ…طµط±ظˆظپط§طھ ط§ظ„طھط´ط؛ظٹظ„ظٹط©',bg:C.hdr3,bold:true,isHeader:true},
    {r:20,l:'ظ…طµط±ظˆظپط§طھ ط¥ط¯ط§ط±ظٹط© ط¹ط§ظ…ط©',cat:'ظ…طµط±ظˆظپط§طھ ط¥ط¯ط§ط±ظٹط© ط¹ط§ظ…ط©',type:'طµط§ط¯ط±'},
    {r:21,l:'ظ…ط­ط±ظˆظ‚ط§طھ',cat:'ظ…ط­ط±ظˆظ‚ط§طھ (ط¨ظ†ط²ظٹظ†+ط²ظٹطھ+ط¯ظٹط²ظ„)',type:'طµط§ط¯ط±'},
    {r:22,l:'ظˆط¬ط¨ط§طھ ظˆظ…ط´ط±ظˆط¨ط§طھ',cat:'ظˆط¬ط¨ط§طھ ظˆظ…ط´ط±ظˆط¨ط§طھ',type:'طµط§ط¯ط±'},
    {r:23,l:'ط±ظˆط§طھط¨ ظˆظٹظˆظ…ظٹط§طھ ظˆط³ظ„ظپ',cat:'ط±ظˆط§طھط¨ ظˆظٹظˆظ…ظٹط§طھ ظˆط³ظ„ظپ',type:'طµط§ط¯ط±'},
    {r:24,l:'طµظٹط§ظ†ط© ط³ظٹط§ط±ط§طھ',cat:'طµظٹط§ظ†ط© ط³ظٹط§ط±ط§طھ',type:'طµط§ط¯ط±'},
    {r:25,l:'ط¥ظٹط¬ط§ط±ط§طھ',cat:'ط¥ظٹط¬ط§ط±ط§طھ',type:'طµط§ط¯ط±'},
    {r:26,l:'ظƒظ‡ط±ط¨ط§ط، ظˆظ…ط§ط،',cat:'ظƒظ‡ط±ط¨ط§ط، ظˆظ…ط§ط،',type:'طµط§ط¯ط±'},
    {r:27,l:'ظ…طµط±ظˆظپط§طھ طھط³ظˆظٹظ‚ظٹط©',cat:'ظ…طµط±ظˆظپط§طھ طھط³ظˆظٹظ‚ظٹط©',type:'طµط§ط¯ط±'},
    {r:28,l:'ظ…طµط±ظˆظپط§طھ ظ…طھظ†ظˆط¹ط©',cat:'ظ…طµط±ظˆظپط§طھ ظ…طھظ†ظˆط¹ط©',type:'طµط§ط¯ط±'},
    {r:29,l:'ط¥ط¬ظ…ط§ظ„ظٹ ط§ظ„ظ…طµط±ظˆظپط§طھ',bg:C.lRed,bold:true,sum:'B20:B28'},
    {r:31,l:'طµط§ظپظٹ ط§ظ„ط¯ط®ظ„',bg:'#1B5E20',color:C.wht,bold:true,big:true,formula:'=B17-B29'},
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

// â•گâ•گâ•گ MONTHLY INCOME STATEMENT (Cash + Accrual side by side) â•گâ•گâ•گ
function buildMonthlyIncome_(ss){
  var sh=getOrCreate_(ss,SN.incomeM);sh.clear();sh.setRightToLeft(true);
  sh.getRange(1,1).setValue('ظ‚ط§ط¦ظ…ط© ط§ظ„ط¯ط®ظ„ ط§ظ„ط´ظ‡ط±ظٹط© â€” ظ†ظ‚ط¯ظٹ ظˆط§ط³طھط­ظ‚ط§ظ‚').setFontSize(14).setFontWeight('bold');
  sh.getRange(1,1,1,4).merge().setBackground(C.suc).setFontColor(C.wht).setHorizontalAlignment('center');

  sh.getRange(2,1).setValue('ط§ظ„ط´ظ‡ط±:').setFontWeight('bold');
  sh.getRange(2,2).setValue('ظٹظ†ط§ظٹط±').setBackground(C.lYlw);
  addDV_(sh,2,2,1,['ظٹظ†ط§ظٹط±','ظپط¨ط±ط§ظٹط±','ظ…ط§ط±ط³','ط£ط¨ط±ظٹظ„','ظ…ط§ظٹظˆ','ظٹظˆظ†ظٹظˆ','ظٹظˆظ„ظٹظˆ','ط£ط؛ط³ط·ط³','ط³ط¨طھظ…ط¨ط±','ط£ظƒطھظˆط¨ط±','ظ†ظˆظپظ…ط¨ط±','ط¯ظٹط³ظ…ط¨ط±']);

  sh.getRange(3,1).setValue('ط§ظ„ط¨ظٹط§ظ†').setFontWeight('bold');
  sh.getRange(3,2).setValue('ط£ط³ط§ط³ ظ†ظ‚ط¯ظٹ').setFontWeight('bold').setBackground('#A5D6A7');
  sh.getRange(3,3).setValue('ط£ط³ط§ط³ ط§ط³طھط­ظ‚ط§ظ‚').setFontWeight('bold').setBackground('#BBDEFB');
  sh.getRange(3,4).setValue('ط§ظ„ظپط±ظ‚').setFontWeight('bold').setBackground(C.lYlw);

  var e="'"+SN.entry+"'!";
  var mo='$B$2';
  var items=[
    {r:5,l:'ط¥ط¬ظ…ط§ظ„ظٹ ط§ظ„ظ…ط¨ظٹط¹ط§طھ',type:'ظˆط§ط±ط¯'},
    {r:6,l:'ظ…ط±ط¯ظˆط¯ط§طھ ط§ظ„ظ…ط¨ظٹط¹ط§طھ',type:'ظ…ط±ط¯ظˆط¯',neg:true},
    {r:7,l:'طµط§ظپظٹ ط§ظ„ظ…ط¨ظٹط¹ط§طھ',formula:'=B5-B6',bold:true,bg:C.lGrn},
    {r:9,l:'ظ…ط´طھط±ظٹط§طھ',type:'طµط§ط¯ط±',cat:'ظ…ط´طھط±ظٹط§طھ ط¹ط§ظ…ط©'},
    {r:10,l:'ط±ظˆط§طھط¨',type:'طµط§ط¯ط±',cat:'ط±ظˆط§طھط¨ ظˆظٹظˆظ…ظٹط§طھ ظˆط³ظ„ظپ'},
    {r:11,l:'ظ…ط­ط±ظˆظ‚ط§طھ',type:'طµط§ط¯ط±',cat:'ظ…ط­ط±ظˆظ‚ط§طھ (ط¨ظ†ط²ظٹظ†+ط²ظٹطھ+ط¯ظٹط²ظ„)'},
    {r:12,l:'ط¥ظٹط¬ط§ط±ط§طھ',type:'طµط§ط¯ط±',cat:'ط¥ظٹط¬ط§ط±ط§طھ'},
    {r:13,l:'ظ…طµط±ظˆظپط§طھ ط¥ط¯ط§ط±ظٹط©',type:'طµط§ط¯ط±',cat:'ظ…طµط±ظˆظپط§طھ ط¥ط¯ط§ط±ظٹط© ط¹ط§ظ…ط©'},
    {r:14,l:'ظ…طµط±ظˆظپط§طھ ط£ط®ط±ظ‰',type:'طµط§ط¯ط±',other:true},
    {r:15,l:'ط¥ط¬ظ…ط§ظ„ظٹ ط§ظ„ظ…طµط±ظˆظپط§طھ',formula:'=SUM(B9:B14)',bold:true,bg:C.lRed},
    {r:17,l:'طµط§ظپظٹ ط§ظ„ط¯ط®ظ„',formula:'=B7-B15',bold:true,big:true,bg:'#1B5E20',color:C.wht},
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

  sh.setColumnWidths(1,4,[200,130,130,100]);
  sh.setFrozenRows(3);sh.setTabColor('#1B5E20');
}

// â•گâ•گâ•گ DASHBOARD â•گâ•گâ•گ
function buildDashboard_(ss){
  var sh=getOrCreate_(ss,SN.dash);sh.clear();sh.setRightToLeft(true);
  sh.getRange(1,1).setValue('âڑ، Qayd â€” ظ„ظˆط­ط© ط§ظ„طھط­ظƒظ…').setFontSize(18).setFontWeight('bold');
  sh.getRange(1,1,1,8).merge().setBackground(C.pri).setFontColor(C.wht).setHorizontalAlignment('center');

  var e="'"+SN.entry+"'!";

  // â•گ Section 1: Live Balances â•گ
  sh.getRange(3,1).setValue('ًں’° ط§ظ„ط£ط±طµط¯ط© ط§ظ„ظ„ط­ط¸ظٹط©').setFontSize(14).setFontWeight('bold');
  sh.getRange(3,1,1,8).merge().setBackground(C.acc).setFontColor(C.wht).setHorizontalAlignment('center');

  var bals=[['ط§ظ„طµظ†ط¯ظˆظ‚','طµظ†ط¯ظˆظ‚'],['ط¨ظ†ظƒ ط§ظ„ط±ط§ط¬ط­ظٹ','ط¨ظ†ظƒ ط§ظ„ط±ط§ط¬ط­ظٹ'],['ط¨ظ†ظƒ ط§ظ„ط£ظ‡ظ„ظٹ','ط¨ظ†ظƒ ط§ظ„ط£ظ‡ظ„ظٹ']];
  for(var i=0;i<bals.length;i++){
    var r=4+i;
    sh.getRange(r,1).setValue(bals[i][0]).setFontWeight('bold').setFontSize(12);
    sh.getRange(r,2).setFormula(
      '=SUMIFS('+e+'D$4:D$1003,'+e+'F$4:F$1003,"'+bals[i][1]+'",'+e+'C$4:C$1003,"ظˆط§ط±ط¯")-SUMIFS('+e+'D$4:D$1003,'+e+'F$4:F$1003,"'+bals[i][1]+'",'+e+'C$4:C$1003,"طµط§ط¯ط±")'
    ).setFontSize(14).setFontWeight('bold').setNumberFormat('#,##0.00');
  }
  sh.getRange(7,1).setValue('ًں“ٹ ط§ظ„ط±طµظٹط¯ ط§ظ„ظƒظ„ظٹ').setFontWeight('bold').setFontSize(13).setBackground(C.lYlw);
  sh.getRange(7,2).setFormula('=SUM(B4:B6)').setFontWeight('bold').setFontSize(16).setNumberFormat('#,##0.00').setBackground(C.lYlw);

  // â•گ Section 2: Quick KPIs â•گ
  sh.getRange(9,1).setValue('ًں“ˆ ظ…ط¤ط´ط±ط§طھ ط§ظ„ط£ط¯ط§ط،').setFontSize(14).setFontWeight('bold');
  sh.getRange(9,1,1,8).merge().setBackground('#1565C0').setFontColor(C.wht).setHorizontalAlignment('center');

  var kpis=[
    ['ط¥ط¬ظ…ط§ظ„ظٹ ط§ظ„ط¥ظٹط±ط§ط¯ط§طھ','=SUMIF('+e+'C$4:C$1003,"ظˆط§ط±ط¯",'+e+'D$4:D$1003)'],
    ['ط¥ط¬ظ…ط§ظ„ظٹ ط§ظ„ظ…طµط±ظˆظپط§طھ','=SUMIF('+e+'C$4:C$1003,"طµط§ط¯ط±",'+e+'D$4:D$1003)'],
    ['طµط§ظپظٹ ط§ظ„ط¯ط®ظ„','=B10-B11'],
    ['ط¹ط¯ط¯ ط§ظ„ظ‚ظٹظˆط¯','=COUNTA('+e+'A$4:A$1003)'],
    ['ظ…طھظˆط³ط· ظ‚ظٹظ…ط© ط§ظ„ظ‚ظٹط¯','=IF(B13>0,ROUND((B10+B11)/B13,0),0)'],
  ];
  for(var k=0;k<kpis.length;k++){
    var kr=10+k;
    sh.getRange(kr,1).setValue(kpis[k][0]).setFontWeight('bold');
    sh.getRange(kr,2).setFormula(kpis[k][1]).setFontWeight('bold').setNumberFormat('#,##0.00');
  }
  sh.getRange(12,2).setFontSize(16).setBackground(sh.getRange(12,2).getFormula().indexOf('-')>-1?C.lRed:C.lGrn);

  // â•گ Section 3: Smart Alerts â•گ
  sh.getRange(16,1).setValue('ًںڑ¨ طھظ†ط¨ظٹظ‡ط§طھ ط°ظƒظٹط©').setFontSize(14).setFontWeight('bold');
  sh.getRange(16,1,1,8).merge().setBackground(C.dan).setFontColor(C.wht).setHorizontalAlignment('center');

  sh.getRange(17,1).setValue('ط±طµظٹط¯ ط§ظ„طµظ†ط¯ظˆظ‚').setFontWeight('bold');
  sh.getRange(17,2).setFormula('=IF(B4<0,"âڑ ï¸ڈ ط¹ط¬ط² ظپظٹ ط§ظ„طµظ†ط¯ظˆظ‚!","âœ… ط·ط¨ظٹط¹ظٹ")');
  sh.getRange(18,1).setValue('ط¹ظ‡ط¯ ظ…ظپطھظˆط­ط©').setFontWeight('bold');
  sh.getRange(18,2).setFormula("=IF(COUNTIF('"+SN.custody+"'!E:E,\"âڑ ï¸ڈ*\")>0,\"âڑ ï¸ڈ ظٹظˆط¬ط¯ ط¹ظ‡ط¯ ط؛ظٹط± ظ…ط³ظˆط§ط©\",\"âœ… ظ„ط§ طھظˆط¬ط¯\")");
  sh.getRange(19,1).setValue('ظ…ط´ط§ط±ظٹط¹ ط®ط§ط³ط±ط©').setFontWeight('bold');
  sh.getRange(19,2).setFormula("=IF(COUNTIF('"+SN.projects+"'!H:H,\"â‌Œ*\")>0,\"âڑ ï¸ڈ ظٹظˆط¬ط¯ ظ…ط´ط§ط±ظٹط¹ ط®ط§ط³ط±ط©\",\"âœ… ظƒظ„ ط§ظ„ظ…ط´ط§ط±ظٹط¹ ط±ط§ط¨ط­ط©\")");

  // â•گ Section 4: Monthly Comparison â•گ
  sh.getRange(21,1).setValue('ًں“ٹ ظ…ظ‚ط§ط±ظ†ط© ط´ظ‡ط±ظٹط©').setFontSize(14).setFontWeight('bold');
  sh.getRange(21,1,1,8).merge().setBackground('#6A1B9A').setFontColor(C.wht).setHorizontalAlignment('center');
  var months=['ظٹظ†ط§ظٹط±','ظپط¨ط±ط§ظٹط±','ظ…ط§ط±ط³','ط£ط¨ط±ظٹظ„','ظ…ط§ظٹظˆ','ظٹظˆظ†ظٹظˆ','ظٹظˆظ„ظٹظˆ','ط£ط؛ط³ط·ط³','ط³ط¨طھظ…ط¨ط±','ط£ظƒطھظˆط¨ط±','ظ†ظˆظپظ…ط¨ط±','ط¯ظٹط³ظ…ط¨ط±'];
  sh.getRange(22,1).setValue('ط§ظ„ط´ظ‡ط±').setFontWeight('bold');
  sh.getRange(22,2).setValue('ط¥ظٹط±ط§ط¯ط§طھ').setFontWeight('bold');
  sh.getRange(22,3).setValue('ظ…طµط±ظˆظپط§طھ').setFontWeight('bold');
  sh.getRange(22,4).setValue('طµط§ظپظٹ').setFontWeight('bold');
  for(var m=0;m<12;m++){
    var mr=23+m;
    sh.getRange(mr,1).setValue(months[m]);
    sh.getRange(mr,2).setFormula('=SUMPRODUCT(('+e+'J$4:J$1003="'+months[m]+'")*('+e+'C$4:C$1003="ظˆط§ط±ط¯")*'+e+'D$4:D$1003)').setNumberFormat('#,##0');
    sh.getRange(mr,3).setFormula('=SUMPRODUCT(('+e+'J$4:J$1003="'+months[m]+'")*('+e+'C$4:C$1003="طµط§ط¯ط±")*'+e+'D$4:D$1003)').setNumberFormat('#,##0');
    sh.getRange(mr,4).setFormula('=B'+mr+'-C'+mr).setNumberFormat('#,##0');
  }

  sh.setColumnWidths(1,4,[200,150,150,120]);
  sh.setFrozenRows(2);sh.setTabColor(C.pri);
}

// â•گâ•گâ•گ PROTECTION â•گâ•گâ•گ
function protectQayd(){
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var me=Session.getEffectiveUser();
  var cnt=0;
  ss.getSheets().forEach(function(sh){
    var name=sh.getName();
    if(name===SN.entry||name===SN.settings) return;
    sh.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(function(p){p.remove();});
    var p=sh.protect().setDescription('ظ„ظ„ظ‚ط±ط§ط،ط© ظپظ‚ط· â€” '+name);
    p.addEditor(me);
    p.getEditors().forEach(function(e){if(e.getEmail()!==me.getEmail())p.removeEditor(e);});
    cnt++;
  });
  // Protect formulas in entry sheet
  var entry=ss.getSheetByName(SN.entry);
  if(entry){
    var p=entry.protect().setDescription('ط­ظ…ط§ظٹط© ظ…ط¹ط§ط¯ظ„ط§طھ ط§ظ„ط¥ط¯ط®ط§ظ„');
    p.addEditor(me);
    var unp=[];
    [1,2,3,4,5,6,7,8].forEach(function(c){unp.push(entry.getRange(4,c,1000,1));});
    p.setUnprotectedRanges(unp);
    p.getEditors().forEach(function(e){if(e.getEmail()!==me.getEmail())p.removeEditor(e);});
  }
  SpreadsheetApp.getUi().alert('âœ… طھظ… ط­ظ…ط§ظٹط© '+cnt+' ط´ظٹطھ + ظ…ط¹ط§ط¯ظ„ط§طھ ط³ط¬ظ„ ط§ظ„ظ‚ظٹظˆط¯');
}
