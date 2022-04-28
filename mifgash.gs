var dowmap={'0':'א', '1':'ב', '2':'ג', '3':'ד', '4':'ה','5':'ו','6':'ז'};
var dmy_fmt='y';

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('More')
      .addItem('Copy to recur', 'cpToRecur')
      .addItem('Sort recur by teacher', 'sortRecurByTeacher')
      .addItem('Sort recur by time', 'sortRecurByTime')
      .addItem('Report absence', 'ReportAbsence')
      .addItem('Filter selected teacher', 'filterTeacher')
      .addItem('Update daily sheets', 'aprj.updateShibSheets')
      .addToUi();
}

function fmt_dmy_date(dt){
  let res=dt.replace(/^(\d+)[\/\.](\d+)[\/\.](\d+)/, "$2/$1/$3");
  return res;
}

function filterTeacher(){
  let sh=SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  let val = sh.getSelection().getCurrentCell().getValue();
  let rng=sh.getRange('A:Q');
  let filt=rng.createFilter();
  let fc=SpreadsheetApp.newFilterCriteria().whenTextEqualTo(val);
  filt.setColumnFilterCriteria(3,fc);
}

function ReportAbsence(){
  aprj.collectParams();
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
      doRngRow(sh,rngcols,rng_ar,r,r1,offset,rows2push,rng_errs);
    }
    if (rng_errs.length){
      errs=errs.concat(rng_errs);
    } else {
      ranges[i].setFontColor(red);
    }
  }
  aprj.appendRows2Maakav(rows2push);
  if (errs.length){
    SpreadsheetApp.getUi().alert(errs.join('\n'));
  }
}

function doRngRow(sh,rngcols,rng_ar,r,r1,offset,rows2push,errs){
      let selrow_ar=sh.getRange(r1+r,1,1,17).getValues();
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

function cpToRecur(){
  let sh=SpreadsheetApp.getActiveSheet();
  let dt=sh.getName().replace(/^\d./,'');
  if (dmy_fmt){
    //Logger.log('bdt='+dt);
    dt=fmt_dmy_date(dt);
    //Logger.log('edt='+dt);
  }
  let dt1= new Date(dt);
  if (isNaN(dt1) || sh.getTabColor() == '#1129e9'){
    let ui = SpreadsheetApp.getUi();
    ui.alert('Can not copy from this sheet: '+sh.getName());
    return;
  }
  //Logger.log('dt='+dt+' dt1='+dt1+'dt1.getDate()='+dt1.getDate());
  let row=sh.getActiveRange().getRow();
  let vals=sh.getRange(row,1,1,15).getValues()[0];
  //Logger.log('row='+row+' vals='+vals);
  let pvals=[dowmap[dt1.getDay()]];
  let rcd=addToRecur(pvals.concat(vals.slice(0,3)).concat(vals.slice(4).concat(['קבוע'])));
  sh.getRange(row,16,1,1).setValue('קבוע');
}

function getRecurSh(){
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName('recur');
}

function addToRecur(flds_ar){
  let rsh=getRecurSh();
  let lrow=rsh.getLastRow();
  //Logger.log('lrow='+lrow+' arr len='+flds_ar.length);
  //Logger.log('flds_ar='+flds_ar);
  rsh.insertRowAfter(lrow);
  rsh.getRange(lrow+1,1,1,16).setValues([flds_ar]);
}

function sortRecurByTeacher() {
  let sh = getRecurSh();
  let lrow=sh.getLastRow()
  sh.getRange(2,1,lrow-1,16).sort([{column: 4, ascending: true}, {column: 1, ascending: true}, {column: 2, ascending: true}]);
}

function sortRecurByTime() {
  let sh = getRecurSh();
  let lrow=sh.getLastRow()
  sh.getRange(2,1,lrow-1,16).sort([{column: 1, ascending: true}, {column: 2, ascending: true}, {column: 4, ascending: true}]);
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
