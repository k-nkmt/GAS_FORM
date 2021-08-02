  //シート読み取り
var forms     = SpreadsheetApp.getActive().getSheetByName('forms');
var formData = forms.getDataRange().getValues();
formData.shift();

var fieldlist = SpreadsheetApp.getActive().getSheetByName('fieldlist');
var fieldData = fieldlist.getDataRange().getValues();
// fieldData.shift();

var codelist= SpreadsheetApp.getActive().getSheetByName('codelist');
var codeData = codelist.getDataRange().getValues();
codeData.shift();

function onOpen() {
  var ui = SpreadsheetApp.getUi()
  var menu = ui.createMenu("スクリプト実行"); 
  menu.addItem("フォーム生成","makeForm");
  menu.addToUi();
}

//フォーム作成
function makeForm() {
  var folder  = DriveApp.createFolder("form_master");

  //フォーム名と説明文を取得し、フォームを作成
  for(var i = 0 ; i < formData.length ; i++){
    var form = FormApp.create(formData[i][1]);
    form.setDescription(formData[i][2]);
    // form.setLimitOneResponsePerUser(TRUE);
    makeField(form= form,name=formData[i][0]);
    var file=DriveApp.getFileById(form.getId());
    file.moveTo(folder);
  }
}

//フィールド作成
 function makeField(form,name){

  var col = fieldData[0].flat();
  var formfields = fieldData.filter(function(e){return e[0] === name}) ;

  for(var i = 0 ; i < formfields.length ; i++){
    if (formfields[i][col.indexOf("image")]) {
    var img = UrlFetchApp.fetch(formfields[i][col.indexOf("image")]);
        form.addImageItem()
            .setImage(img);
          }

    switch (formfields[i][col.indexOf("type")]){
      case "num":
        var textValidation = FormApp.createTextValidation()
          .requireNumber()
          .build();
        var item = form.addTextItem();
        item.setValidation(textValidation);
      break;
      case "char":
        var item = form.addTextItem();
        if (formfields[i][col.indexOf("regex")]) {
        var textValidation = FormApp.createTextValidation()
          .requireTextMatchesPattern(formfields[i][col.indexOf("regex")])
          .setHelpText(formfields[i][col.indexOf("message")])
          .build();
        item.setValidation(textValidation);
        }
      break;
      case "email":
        var item = form.addTextItem();
        var textValidation = FormApp.createTextValidation()
          .requireTextIsEmail()
          .build();
        item.setValidation(textValidation);
      break;
      case "text":
        var item = form.addParagraphTextItem();
      break;
      case "date":
        var item = form.addDateItem();
      break;
      case "time":
        var item = form.addTimeItem();
      break;
      case "dur":
        var item = form.addDurationItem();
      break;
      case "select":
        var item = form.addListItem();
        makeChoice(item,formfields[i][col.indexOf("codelist")]);
      break;
      case "radio":
        var item = form.addMultipleChoiceItem();
        makeChoice(item,formfields[i][col.indexOf("codelist")]);
      break;
      case "checkbox":
        var item = form.addCheckboxItem();
        makeChoice(item,formfields[i][col.indexOf("codelist")]);
      break;
      case "scale":
        var item = form.addScaleItem();
        makeBounds(item,formfields[i][col.indexOf("codelist")]);
      break;
      case "head":
        var item = form.addSectionHeaderItem();
      break;
      case "page":
        var item = form.addPageBreakItem();
      break;
     }
    
    item.setTitle(formfields[i][col.indexOf("question")]);
    item.setHelpText(formfields[i][col.indexOf("description")]);
    if (!(["head","page"].includes(formfields[i][col.indexOf("type")]))){
        item.setRequired(formfields[i][col.indexOf("required")]);
        }
  }
}

//スケール作成
function makeBounds(item,name){
  var codes = codeData.filter(function(e){return e[0] === name}) ;
  item.setBounds(codes[0][1], codes[1][1])
      .setLabels(codes[0][2],codes[1][2]);
}

//選択肢作成
function makeChoice(item,name){
  var codes = codeData.filter(function(e){return e[0] === name}) ;
  var codeChoice = [];

  for(var i = 0 ; i < codes.length ; i++){
    switch(codes[i][2]){
      case "SUBMIT":
        codeChoice.push(item.createChoice(codes[i][1], FormApp.PageNavigationType.SUBMIT)) ;
      break;
      case "CONTINUE":
        codeChoice.push(item.createChoice(codes[i][1], FormApp.PageNavigationType.CONTINUE)) ;
      break;
      default:
        codeChoice.push(item.createChoice(codes[i][1])) ;
      break;
    }
  }
  item.setChoices(codeChoice);
}