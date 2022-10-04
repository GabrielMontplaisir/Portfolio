function getSheetbyId(id) {
  return SpreadsheetApp.getActive().getSheets().filter(
    function(s) {return s.getSheetId().toString() === id;}
  )[0];
}

// function getID() {return PropertiesService.getScriptProperties().getProperty("IDCol")}
function getComment() {return PropertiesService.getScriptProperties().getProperty('commentCol')}

// function updateID() {
//   var currentid = SpreadsheetApp.getCurrentCell().getValue();
//   PropertiesService.getScriptProperties().setProperty('IDCol', currentid);
//   return currentid
// }

function updateComment() {
  var currentComment = SpreadsheetApp.getCurrentCell().getValue();
  PropertiesService.getScriptProperties().setProperty('commentCol', currentComment);
  return currentComment
}

function getDocProps() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName('Document Properties');
  let uObj=PropertiesService.getDocumentProperties().getProperties();
  let keys = Object.keys(uObj);
  sh.clearContents();
  let a=[['Key','Value']];
  keys.forEach(k => {a.push([k,uObj[k]]);});
  sh.getRange(1,1,a.length, a[0].length).setValues(a);
  ss.toast('Document Properties generated.')
}

function delDocProps() {
  PropertiesService.getDocumentProperties().deleteAllProperties();
}