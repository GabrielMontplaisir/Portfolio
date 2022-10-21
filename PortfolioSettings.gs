function getComment() {return PropertiesService.getScriptProperties().getProperty('commentCol')}

function updateComment() {
  var currentComment = SpreadsheetApp.getCurrentCell().getValue();
  PropertiesService.getScriptProperties().setProperty('commentCol', currentComment.toString().toLowerCase());
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