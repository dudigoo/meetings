var permanent="קבוע";

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('More')
      .addItem('Worker recur', 'aprj.flagPermanentMeetings')
      .addItem('Copy to recur', 'cpToRecur')
      .addItem('Update current DOW sheets', 'aprj.updateDayOfWeekSheetsMain')
      .addItem('---', '***')
      .addItem('Update current date sheet', 'aprj.updateShibCurrSheet')
      .addItem('Student details', 'studentDetails')
      .addItem('Report absence', 'ReportAbsence')
      .addItem('Copy to recur & update DOWs', 'cpToRecurAndDOWs')
      .addItem('---', '***')
      .addItem('Update next 5 days', 'aprj.update5ShibSheetsFrom2Main')
      .addItem('Update all sheets from next day', 'aprj.updateShibSheetsFrom2Main')
      .addItem('Update **all** sheets', 'aprj.updateShibSheetsMain')
      .addItem('Color current sheet window rows', 'aprj.ColorCurrentSheetWindowRows')
      .addItem('Email error report', 'aprj.findAllGroupsSchedMistakesMain')
      .addItem('Update allDays sheet', 'aprj.updateAllDatesSheetMain')
      .addItem('Update blocked students', 'aprj.updateShibBlockedStudentListsMain')
      .addToUi();
}

function studentDetails(){
  aprj.collectParams();
  let sh=SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  let val = sh.getSelection().getCurrentCell().getValue();
  let stu_ar=aprj.getStuAr(val);
  Logger.log('stu details:'+stu_ar);
  let msg;
  if (!stu_ar){
    msg='Student not found: '+val+'. Select a student cell';
  } else {
    msg='Student: '+val+ ' Grade: '+stu_ar[0]+' Group: '+stu_ar[3]+'\nMobile: '+stu_ar[4]+'\nMath: '+stu_ar[15]+'\nEnglish: '+stu_ar[16]+'\nLang: '+stu_ar[17]+'\nNote: '+stu_ar[18];
    let query='select * where B starts with "'+stu_ar[0]+stu_ar[3]+'"';
    let shnm='מדריכי פנימיה';
    let wrkr=aprj.querySheet(query, aprj.gp.wrkrs_ss_id,shnm,1);
    Logger.log('wrkr:'+JSON.stringify( wrkr));
    for (let i=0;i<wrkr.length;i++){
      msg= msg+ '\n'+wrkr[i][0]+' : '+wrkr[i][3] + ' : '+wrkr[i][2];
      Logger.log('loop:'+wrkr[i][0]+' : '+wrkr[i][3] + ' : '+wrkr[i][2]);
    }
    Logger.log('msg:'+msg);
  }
  SpreadsheetApp.getUi().alert(msg);
}


function wrkrRecur(){
  aprj.collectParams();
  let sh  = SpreadsheetApp.getActiveSheet();
  let dow=aprj.dowmap[aprj.getDtObjFromTabNm(SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName()).getDay()];
  let selection  = sh.getSelection();
  let row=selection.getActiveRangeList().getRanges()[0].getRow();
  //let msg=sh.getRange(row,3,1,1).getValue();
  //let msg=ScriptApp.getService().getUrl();
  //SpreadsheetApp.getUi().alert(msg);
  //return;
  let qry="select B,C,D where A='"+dow+"' and B='" +frtm+"' and D='"+wrkr+"'"; 
  let vals = aprj.querySheet(qry, aprj.getShibutzSS(), 'recur', 1);
  if (! vals){
    msg='Teacher not found for this date/time';
  } else {
    msg='Teacher found for this date/time';
  }
  SpreadsheetApp.getUi().alert(msg);
}

function ReportAbsence(){
  aprj.collectParams();
  let res=iterateSelectedRanges('absence');
  let rows2push=res[0];
  let errs=res[1];
  aprj.appendRows2Maakav(rows2push);
  if (errs.length){
    SpreadsheetApp.getUi().alert(errs.join('\n'));
  }  
}


function cpToRecur(){
  aprj.collectParams();
  let this_shnm=SpreadsheetApp.getActiveSheet().getName();
  let dt=aprj.getDtStrFromShNm(this_shnm);
  let dt1= aprj.getDtObj(dt);
  if (isNaN(dt1)){
    let ui = SpreadsheetApp.getUi();
    ui.alert('Can not copy from this sheet to recur');
    return;
  }
  aprj.shib_cur_sheet_dow=aprj.dowmap[dt1.getDay()];
  let res=iterateSelectedRanges('cp2recur');
  let rows2push=res[0];
  let errs=res[1];
  let torow=aprj.getRecurSh().getLastRow()+1;
  //Logger.log('rows2push ='+JSON.stringify(rows2push));
  //Logger.log('ar='+ar+' torow='+torow+' ar.length='+ar.length);
  aprj.getRecurSh().getRange(torow,1,rows2push.length,rows2push[0].length).setValues(rows2push);
  return this_shnm;
}


function cpToRecurAndDOWs(){
  let this_shnm=cpToRecur();
  if (this_shnm){
    aprj.updateDOWSheetsMain(this_shnm);
  }
}

function iterateSelectedRanges(type){
  let sh=SpreadsheetApp.getActiveSheet();

  let red='#ea4335';

  let errs=[];
  let selection = sh.getSelection();
  let ranges =  selection.getActiveRangeList().getRanges();
  let rows2push=[];
  let offset = (sh.getName() == 'history') ? 1 : 0;
  for (let i = 0; i < ranges.length; i++) {
    let rngcols=ranges[i].getNumColumns()
    let rng_ar=ranges[i].getValues();
    let r1=ranges[i].getRow();
    let rng_errs=[]
    //Logger.log('selected Range: ' + ranges[i].getA1Notation() +' r1='+r1);
    for (let r=0; r< ranges[i].getNumRows(); r++) {
      let selrow_ar=sh.getRange(r1+r,1,1,17).getValues();
      if (type=='absence'){
        doRngRowAbs(sh,rngcols,rng_ar,r,offset,rows2push,rng_errs,selrow_ar);
        if (rng_errs.length){
          errs=errs.concat(rng_errs);
        } else {
          ranges[i].setFontColor(red);
        }
      } else {//cp2recur
        doRngRowRecur(sh,r,r1,rows2push,selrow_ar);
        if (rng_errs.length){
          errs=errs.concat(rng_errs);
        }
      }
    }
  }
  return [rows2push,errs];
}

function doRngRowRecur(sh,r,r1,rows2push,selrow_ar){
  //Logger.log('selrow_ar='+JSON.stringify(selrow_ar));
  let vals=[aprj.shib_cur_sheet_dow].concat(selrow_ar[0].slice(0,3)).concat(selrow_ar[0].slice(4,15)).concat([permanent]);
  rows2push.push(vals);
  sh.getRange(r1+r,16).setValue(permanent);
  sh.getRange(r1+r,6).setValue(aprj.setPermanentComment(sh.getRange(r1+r,6).getValue()));

}

function doRngRowAbs(sh,rngcols,rng_ar,r,offset,rows2push,errs,selrow_ar){
      //Logger.log(' row='+selrow_ar);
      let acti = 'חיסור ידני משיבוץ';
      let atd = 'לא הגיע';
      let subj=selrow_ar[0][offset+4];
      let teac=selrow_ar[0][offset+2];
      let lvl=selrow_ar[0][offset+6];
      let frtm=selrow_ar[0][offset+0];
      let totm=selrow_ar[0][offset+1];
      let dt = (sh.getName() == 'history') ? selrow[0][0] : aprj.getDtObjFromTabNm(sh.getName());
      for (let c=0; c< rngcols; c++) {
        //Logger.log('c='+c+' r='+r);
        //Logger.log('selrow_ar='+selrow_ar);
        let stu=rng_ar[r][c];
        if (stu && dt && subj && teac) {
          let ar= [dt, subj, acti, '', teac, stu, lvl, '', '=ROW()', 1, atd, '', '', frtm, totm];
          rows2push.push(ar);
        } else {
          let msg='missing info. pupil='+stu+' ';
          msg += (teac ? '' : ' no teacher '); 
          msg += (dt ? '' : ' no date '); 
          msg += (stu ? '' : ' no pupil '); 
          msg += (subj ? '' : ' no subject '); 
          errs.push(msg);
          Logger.log('missing info. kid='+stu+' teacher='+teac+' subject='+subj+' dt='+dt);
          Logger.log(' row='+selrow_ar);
        }
      }
}

function sortRecurByTeacher() {
  aprj.collectParams();
  let sh = aprj.getRecurSh();
  let lrow=sh.getLastRow()
  sh.getRange(2,1,lrow-1,23).sort([{column: 4, ascending: true}, {column: 1, ascending: true}, {column: 2, ascending: true}]);
}

function sortRecurByTime() {
  aprj.collectParams();
  let sh = aprj.getRecurSh();
  let lrow=sh.getLastRow()
  sh.getRange(2,1,lrow-1,23).sort([{column: 1, ascending: true}, {column: 2, ascending: true}, {column: 4, ascending: true}]);
}


function onEdit(e){
  //var ss=SpreadsheetApp.getActiveSheet();
  var ss = e.source;
  var sh=ss.getActiveSheet();
  //Logger.log('start onEdit');
  if (['template','lists'].includes(sh.getName())){
    return;
  }
  var aCell = sh.getActiveCell();
  var aColumn = aCell.getColumn();
  var aRow = aCell.getRow();
  
  if(aRow<2 || aRow >sh.getLastRow()){
    return;
  }
  //Logger.log('start row:'+aRow + ' col:'+aColumn);
  if (aColumn == 7){
    //Logger.log('sh='+sh.getName());
    //Logger.log('grade selected row='+aRow);
    var sourceRange = ss.getRangeByName(aCell.getValue());
    var rule = SpreadsheetApp.newDataValidation().requireValueInRange(sourceRange, true).build();
    var trange = sh.getRange(aRow, 9, 1, 7);
    trange.setDataValidation(rule);
    //var trange = sh.getRange(aRow, aColumn + 2);
    //Logger.log('end');
  } 

}
