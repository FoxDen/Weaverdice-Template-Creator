/**
@NotOnlyCurrentDoc
*/

function onInstall(e){
    onOpen(e);
    
    var userProperties = PropertiesService.getUserProperties();
    userProperties.setProperty('hFont', 'Arial');
    userProperties.setProperty('tFont', 'Arial');
    userProperties.setProperty('city', 'Default');
    userProperties.setProperty('month','January');
    userProperties.setProperty('day','01');
    userProperties.setProperty('year','2009');
    userProperties.setProperty('area','USA');
    userProperties.setProperty('color','#ffffff');
    userProperties.setProperty('folderName','Weaverdice RP Campaign' );
    userProperties.setProperty('folderSession','Sessions');
}

function onOpen(e){
  var ui = DocumentApp.getUi();
  ui.createAddonMenu()
    .addItem('Create Campaign Doc','createCampaignDoc')
    .addSubMenu(DocumentApp.getUi().createMenu('Main Page')
                .addItem('Add Section','addSection'))
    .addSubMenu(DocumentApp.getUi().createMenu('Timeline')
                .addItem('Add Regular Log',"norm")
                .addItem('Add Secret Log','sec')
                .addItem('Add Public Log','pub')
                .addItem('Add Public Event','addPubE')
                .addItem('Add PHO Thread','pho')
                .addSeparator()
                .addItem('New Date','date'))
    .addSubMenu(DocumentApp.getUi().createMenu('Important Docs/Know Where You Live')
                .addItem('Add Group','addGrp')
                .addItem('Add PRT','addPRT')
                .addItem('Add Person','pers'))
    .addItem('Customizations','customizationsSideBar')
    .addSeparator()
    .addItem('Info','infoSideBar')
    .addToUi();
}

//city, realCity, year, month, headingF, textF, currPlayers, quePlayers, cityEvent, image, area.
function generate(a,b,c,d,e,f,g,h,i,j,k){
  var body = DocumentApp.getActiveDocument().getBody();
  var userProperties = PropertiesService.getUserProperties();
  if(e!=''&&f!=''){
  userProperties.setProperty('hFont', e);
  userProperties.setProperty('tFont', f);
  }
  
  if(j!=''){
    url = j;
  } else{
    url = "http://www.psdgraphics.com/file/city-skyline-silhouette.jpg";
  }
  userProperties.setProperty('city', a);
  userProperties.setProperty('month',d);
  userProperties.setProperty('year', c);
  userProperties.setProperty('area', k);
  
  var docB = createDoc(a, 'main');
  var docBody = docB.getBody();
  var docBUrl = docB.getUrl();
  var docBID = docB.getId();
  var docW = createDoc(a, 'where');
  var docWhere = docW.getBody();
  var docWUrl = docW.getUrl();
  var docWID = docW.getId();
  var docS = createDoc(a, 'sessions');
  var docSess = docS.getBody();
  var docSUrl = docS.getUrl();
  var docSID = docS.getId();
  
  docBody.setMarginLeft(40).setMarginRight(40);
  docWhere.setMarginLeft(40).setMarginRight(40);
  docSess.setMarginLeft(40).setMarginRight(40);
  
   
  //-- Files and Folders --//
  var currFile = DocumentApp.getActiveDocument().getId();
  var thisFile = DriveApp.getFileById(currFile);
  var parentFolder = thisFile.getParents().next().getId();
  
  DriveApp.getFolderById(parentFolder).addFile(DriveApp.getFileById(docBID));
  DriveApp.getFolderById(parentFolder).addFile(DriveApp.getFileById(docWID));
  DriveApp.getFolderById(parentFolder).addFile(DriveApp.getFileById(docSID));
 
  
  //-- Create 'Where you Live' doc --//
  
  docWhere.editAsText().setFontFamily(userProperties.getProperty('tFont'));
  var heading = docWhere.insertParagraph(0,'Parahumans Online');
  heading.setHeading(DocumentApp.ParagraphHeading.HEADING1);
  var subheading = docWhere.insertParagraph(1,'Know where you live: '+a);
  subheading.setHeading(DocumentApp.ParagraphHeading.SUBTITLE);
  heading.editAsText().setBold(true).setFontFamily(userProperties.getProperty('hFont'));
  subheading.editAsText().setBold(false).setFontFamily(userProperties.getProperty('tFont'));

  var wikiInsert = docWhere.insertParagraph(2,"Welcome to the Parahumans Online wiki for "+a+", "+ k +
  ", and its surroundings. Please contact the local moderation team if you'd like to contribute. "+
  "You'll require a verified PHO account in the "+ k +" subforum, and edits you make to the wiki will need to "+
  "be approved by a moderator until you've made enough valid contributions.");
  
  var modStaff = docWhere.appendParagraph('Wiki Moderation Staff:\n');
  modStaff.editAsText().setBold(0,21,true).setFontFamily(userProperties.getProperty('tFont'));
  
   addGeneratedSection('Groups','Know where you live. Only publicly available information, '+
  'not guaranteed to be accurate.',docWhere,5);

 //-- Create 'Sessions Log' doc --//
  docSess.editAsText().setFontFamily(userProperties.getProperty('tFont'));
  addGeneratedSection('Timeline','The story so far. All information in logs not marked as public information '+
  'are to be treated as meta-knowledge.',docSess,0);
  docSess.appendParagraph('Key').editAsText().setBold(0,2,true);
  var tableR0, tableR1,tableR2,tableR3;
  createKey('#999999','Session:', 'Regular Session.', 'Treat as meta-knowledge.', '#dddddd', tableR0, docSess);
  createKey('#6fa8dc','Digression:','A written exchange, speech, or relevant document.', 'Treat as meta-knowledge.', '#cfe2f3', tableR1, docSess);
  createKey('#ffd966','Public Information:','Events or information that most characters would be at least broadly aware of.','','#fff2cc', tableR2, docSess);
  createKey('#000000','Obscured:', 'Session is private, or lost to the sands of time.', '', '#434343', tableR3, docSess);
  
  var tableOfContents = docSess.appendParagraph('\nTable of Contents\n');
  var arc1 = 'Arc 1';
  var cell = [[arc1]];
  var tableArc = docSess.appendTable(cell);
  var tableCell = tableArc.getRow(0).getCell(0);
  tableCell.getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  tableCell.setVerticalAlignment(DocumentApp.VerticalAlignment.CENTER);
  tableCell.setWidth(50);
  
  var arc = docSess.appendParagraph('\n'+arc1); 
  
   //-- Create 'Main Player Info' doc --//
  docBody.editAsText().setFontFamily(userProperties.getProperty('tFont'));
  var imageInsert = UrlFetchApp.fetch(url);
  docBody.insertImage(0, imageInsert).setHeight(200).setWidth(708);
  var date = docBody.insertParagraph(1,d+' '+c);
  date.setAlignment(DocumentApp.HorizontalAlignment.RIGHT).setLineSpacing(0);
  date.editAsText().setItalic(true).setFontSize(9).setForegroundColor('#878787');
  
  var cityBlurb = docBody.insertParagraph(2,i+'\n');
  cityBlurb.editAsText().setFontSize(11).setForegroundColor('#000000').setItalic(false);

  addGeneratedSection('Links','A few useful documents', docBody,3);
  var map =  'Map of '+a;
  var map2 = 'Non-exhaustive and not necessarily to scale.';
  var maps = docBody.insertParagraph(4,map+'\n'+map2+'\n');
  
  var know = 'Know Where You Live';
  var know2 = 'Publically available and editable knowledge regarding the capes of '+a+'. Not guaranteed to be accurate.';
  var knows = docBody.insertParagraph(5,know+'\n'+know2+'\n');
  var knowLink = knows.findText(know).getElement().asText().setLinkUrl(0, know.length, docWUrl);
  
  var pho = a+' PHO Subforum';
  var pho2 = 'A place for players and spectators to post and discuss events within the city.';
  var phos = docBody.insertParagraph(6,pho+'\n'+pho2+'\n');
  
  var sesshLog = 'Session Logs';
  var sesshLog2 = 'The story thus far. All information in logs not marked as public information are to be treated as meta-knowledge.';
  var logs = docBody.insertParagraph(7,sesshLog+'\n'+sesshLog2);
  var logsLink = logs.findText(sesshLog).getElement().asText().setLinkUrl(0, sesshLog.length, docSUrl);
  
  addGeneratedSection('Players','Those currently in the fray', docBody,8);
  var current = docBody.insertParagraph(9,'Current').setHeading(DocumentApp.ParagraphHeading.HEADING4);
  var headerStyle = {};
  headerStyle[DocumentApp.Attribute.BACKGROUND_COLOR] = '#000000';
  headerStyle[DocumentApp.Attribute.FOREGROUND_COLOR] = '#ffffff';  
  headerStyle[DocumentApp.Attribute.BOLD] = true;
  headerStyle[DocumentApp.Attribute.ITALIC] = false;
  headerStyle[DocumentApp.Attribute.FONT_FAMILY] = userProperties.getProperty('hFont');
  var cellStyle = {};
  cellStyle[DocumentApp.Attribute.FOREGROUND_COLOR] = '#000000';
  cellStyle[DocumentApp.Attribute.BOLD] = false;
  cellStyle[DocumentApp.Attribute.FONT_FAMILY] = userProperties.getProperty('tFont');
  var paraStyle = {};
  paraStyle[DocumentApp.Attribute.SPACING_AFTER] = 0;
  paraStyle[DocumentApp.Attribute.LINE_SPACING] = .5;
  var table = docBody.insertTable(10);
  table.setBorderColor('#666666');
    for(var i = 0;i<=g;i++){
      var tr = table.appendTableRow();
      for(var k = 0; k<5;k++){
        var td = tr.appendTableCell('');
        var para = td.getChild(0).asParagraph();
        para.setAttributes(paraStyle);
        if(i==0){
          td.setAttributes(headerStyle);
          switch(k){
            case 0:
              td.setText('Player');
              break;
            case 1:
              td.setText('Character');
              break;
            case 2:
              td.setText('Fealty');
              break;
            case 3:
              td.setText('Notes');
              break;
            case 4:
              td.setText('✓');
              td.setWidth(15);
            default:
              break;
          }
        } else{ 
            td.setAttributes(cellStyle).setPaddingTop(1).setPaddingBottom(1);      
         }
      }
    }  
  var dead = docBody.insertParagraph(11,'Dead and Gone').setHeading(DocumentApp.ParagraphHeading.HEADING4);
  var deadCells = [['Player', 'Character','Status','Notes']];
  var deadTable = docBody.insertTable(12, deadCells);
  deadTable.setBorderColor('#666666');
  for(var i = 0;i<4;i++){
    var td = deadTable.getCell(0,i);
    td.setAttributes(headerStyle);
    td.getChild(0).asParagraph().setAttributes(paraStyle);
  }

  addGeneratedSection('Queue','Those eager to join the fray', docBody,13);
  var queue = docBody.insertTable(14);
  queue.setBorderColor('#666666');
  for(var i = 0;i<=h;i++){
    var tr = queue.appendTableRow()
    for(var k = 0; k<2;k++){
      var td = tr.appendTableCell('');
      var para = td.getChild(0).asParagraph();
      para.setAttributes(paraStyle);
      if(i==0){
        td.setAttributes(headerStyle);
        switch(k){
          case 0:
            td.setText('Player');
            break;
          case 1:
            td.setText('Queue');
            break;
          default:
            break;
        }
      } else{
          td.setAttributes(cellStyle).setPaddingTop(1).setPaddingBottom(1); 
      }
    }
  }
 

}

function createKey(borderColor, secondCellTextB, secondCellTextR,secondCellTextI,backgroundColor, innerTable, docBody){
  var userProperties = PropertiesService.getUserProperties();
  var endingLength = secondCellTextB.length+secondCellTextR.length+secondCellTextI.length;
  innerTable = docBody.appendTable().setBorderColor(borderColor).appendTableRow();
  innerTable.appendTableCell(['']);
  innerTable.appendTableCell([secondCellTextB+' '+secondCellTextR+' '+secondCellTextI]);
  var cell01 = innerTable.getCell(0).setWidth(20).setPaddingTop(1).setPaddingBottom(1);
  cell01.setBackgroundColor(backgroundColor);
  cell01.getChild(0).asParagraph().setLineSpacing(.5);
  var cell02 = innerTable.getCell(1).setPaddingTop(1).setPaddingBottom(1);
  cell02.editAsText().setBold(0,secondCellTextB.length-1,true);
  if(secondCellTextI.length>1){
    var italic = cell02.findText(secondCellTextI);
    italic.getElement().asText().setItalic(secondCellTextB.length+secondCellTextR.length,endingLength,true);
  }
}

function createDoc(nameOfCity, type){
  var document;
  if(type == 'where'){
    document = DocumentApp.create('Know Where You Live - '+nameOfCity);
  } else if (type == 'main'){
    document = DocumentApp.create('#WD'+nameOfCity + ' - Player Info');
  } else if (type == 'session'){
    document = DocumentApp.create('#WD'+nameOfCity + ' - ?.?');
  } else {
    document = DocumentApp.create('#WD'+nameOfCity + ' - Session Logs');
  }
return document;
}

function addGeneratedSection(section, subtitle, doc, position){  
  var userProperties = PropertiesService.getUserProperties();
  var cursor = DocumentApp.getActiveDocument().getCursor();
  var element = cursor.getElement();
  var parent = element.getParent();
  var cell = [
   [section+'\n'+subtitle]
   ];
  var table;
  if(position==-1){
   if(cursor){
    if(element){
      table = doc.insertTable(parent.getChildIndex(element)+1, cell).setBorderColor('#434343');
    } else{
      DocumentApp.getUi().alert("Element cannot be inserted.");
      }
    } else{
    DocumentApp.getUi().alert("Unable to find cursor.");
      }
  } else{
    table = doc.insertTable(position,cell);
  }
  
  var t = table.findText(section);
  table.setBorderColor('#d2d2d2');
  t.getElement().asText().setBold(0,section.length-1,true).setFontSize(12).setFontFamily(userProperties.getProperty('hFont'));
  var h = table.findText(subtitle);
  h.getElement().asText().setItalic(0,subtitle.length-1,true).setFontSize(10).setFontFamily(userProperties.getProperty('tFont')).setForegroundColor('#666666');
}

function addLog(typeOfLog){
  var userProperties = PropertiesService.getUserProperties();
  var cursor = DocumentApp.getActiveDocument().getCursor();
  var body = DocumentApp.getActiveDocument().getBody();
  var myTable = [['merge!', 'summary'],['merge!','players']];
  var element = cursor.getElement();
  var parent = element.getParent();
  if(cursor){
    if(element){
      var newTable = body.insertTable(parent.getChildIndex(element)+1, myTable).setBorderColor('#434343');
    } else{
      DocumentApp.getUi().alert("Element cannot be inserted.");
    }
  } else {
    DocumentApp.getUi().alert("Unable to find cursor.");
  }
  var row = newTable.getRow(0);
  row.editAsText().setFontFamily(0,13,userProperties.getProperty('tFont'));
  var cell = row.getCell(0).setWidth(50); //merge
  cell.setVerticalAlignment(DocumentApp.VerticalAlignment.CENTER);
  cell.getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  cell.editAsText().setBold(0,5,true).setItalic(false);
  
  var row2 = newTable.getRow(1);
  row2.editAsText().setFontFamily(0,13,userProperties.getProperty('tFont'));
  var cell2 = row2.getCell(0).setWidth(50) //merge
  cell2.getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  cell2.editAsText().setBold(0,5,true).setItalic(false).setFontFamily(0,5,userProperties.getProperty('tFont'));
  
  var sessionLog = createDoc(userProperties.getProperty('city'),'session');
  var sessionLogId = sessionLog.getId();
  var sessionLogUrl = sessionLog.getUrl();
  
    var currFile = DocumentApp.getActiveDocument().getId();
    var thisFile = DriveApp.getFileById(currFile);
    var parentFolder = thisFile.getParents().next().getId();
    DriveApp.getFolderById(parentFolder).addFile(DriveApp.getFileById(sessionLogId));
    
  var returnedLog;
  cell.getChild(0).asText().setLinkUrl(0, 5, sessionLogUrl);
  cell2.getChild(0).asText().setLinkUrl(0, 5, sessionLogUrl);
  
    switch(typeOfLog)
    {
      case "sec":
        cell.setBackgroundColor('#434343');
        cell2.setBackgroundColor('#434343');
        returnedLog = createSession(sessionLog,false);
        break;
      case "pub":
        cell.setBackgroundColor('#fff2cc');
        cell2.setBackgroundColor('#fff2cc');
        returnedLog = createSession(sessionLog,false);
        break;
      case "pho":
        cell.setBackgroundColor('#cfe2f3');
        cell2.setBackgroundColor('#cfe2f3');
        returnedLog = createSession(sessionLog,true);
        break;
      default:
        cell.setBackgroundColor('#dddddd');
        cell2.setBackgroundColor('#dddddd');
        returnedLog = createSession(sessionLog,false);
    }
  
    
}

function createSession(session, isThread){
  var userProperties = PropertiesService.getUserProperties();
  var sessBody = session.getBody();
  if(isThread == false){
  var heading = sessBody.insertParagraph(0,'#WD'+userProperties.getProperty('city')+' - x.x');
  heading.setHeading(DocumentApp.ParagraphHeading.HEADING2);
  heading.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  
  var quote = sessBody.insertParagraph(1,'"quote"');
  quote.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  quote.editAsText().setItalic(0,6,true);
  
  var tableSumm = sessBody.insertTable(2);

  var tellerUrl = 'https://tellergram.github.io/log_formatter.html';
  var row = tableSumm.appendTableRow();
  row.appendTableCell(tellerUrl+'\n\n\n\n');
  var cell = row.getCell(0).getChild(0);
  cell.asParagraph().setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  row.getCell(0).setBackgroundColor('#cccccc').editAsText().setLinkUrl(0,tellerUrl.length-1,tellerUrl);
  
  sessBody.insertHorizontalRule(3);
  var cells = [['← Previous',userProperties.getProperty('month')+' '+userProperties.getProperty('day')+', '+userProperties.getProperty('year'), 'Next →']];
  var navTable = sessBody.insertTable(4,cells).setBorderWidth(0);
  var navRow = navTable.getRow(0);
  var cell1 = navRow.getCell(0);
  var cell2 = navRow.getCell(1).getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  var cell3 = navRow.getCell(2).getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
  sessBody.insertHorizontalRule(5);
  sessBody.insertHorizontalRule(6);
  var sessTable = sessBody.insertTable(7).setBorderWidth(0).appendTableRow().appendTableCell('Sessions');
  var sess = sessTable.getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  
  sessBody.insertHorizontalRule(8);  
  sessBody.editAsText().setFontFamily(userProperties.getProperty('tFont'));
  } else {
    sessBody.editAsText().setBackgroundColor('#073763').setFontFamily(userProperties.getProperty('tFont')).setForegroundColor('#ffffff').setBold(true);
    sessBody.insertParagraph(0,"► Topic: Insert name").setIndentFirstLine(30);
    sessBody.insertParagraph(1,"In: Boards ► Places ► America ► "+userProperties.getProperty('area')+" ► "+userProperties.getProperty('city')).setIndentFirstLine(30);
   
    var cells = [['','']]
    var table = sessBody.appendTable(cells);
   
    var cellTable1 = table.getCell(0, 0);
    cellTable1.setWidth(100).setBackgroundColor('#a4c2f4');
    cellTable1.getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.CENTER).setForegroundColor('#000000').setBold(false);
    var cellTable2 = table.getCell(0, 1);
    cellTable2.setWidth(400).setBackgroundColor('#cfe2f3');
    cellTable2.getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.RIGHT).setForegroundColor('#000000').setBold(false).setItalic(true);
    }  
  
}

function addPubE(){  
  var userProperties = PropertiesService.getUserProperties();
  var body = DocumentApp.getActiveDocument().getBody();
  var cell = [
   ['🛈' + '   \t News']
   ];
  var table = body.appendTable(cell);
  table.setBorderColor('#f1c232');
  var cell = table.getRow(0).getCell(0).setBackgroundColor('#fff2cc');
  cell.setVerticalAlignment(DocumentApp.VerticalAlignment.CENTER);
  cell.editAsText().setBold(false).setItalic(false).editAsText().setFontFamily(userProperties.getProperty('tFont'));
}

function addGrp(){
  var userProperties = PropertiesService.getUserProperties();
  var body = DocumentApp.getActiveDocument().getBody();
  var grp = body.appendParagraph('Group');
  grp.setHeading(DocumentApp.ParagraphHeading.HEADING2);
  grp.editAsText().setFontFamily(userProperties.getProperty('tFont'));
  var info = body.appendParagraph('Information about said group here.');
  info.editAsText().setItalic(0,32,true).setFontFamily(userProperties.getProperty('tFont'));
  
  addPers(false);  
}

function addPRT(){
  var userProperties = PropertiesService.getUserProperties();
  var body = DocumentApp.getActiveDocument().getBody();
  var prtHead = body.appendParagraph('PRT Dept -- : Leadership').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  prtHead.editAsText().setFontFamily(userProperties.getProperty('tFont'));
  addPers(true);
  var prtPrt = body.appendParagraph('PRT Dept -- : Protectorate').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  prtHead.editAsText().setFontFamily(userProperties.getProperty('tFont'));
  addPers(true);
  var prtWards = body.appendParagraph('PRT Dept -- : Wards').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  prtWards.editAsText().setFontFamily(userProperties.getProperty('tFont'));
  addPers(true);  
}

function addPers(isPrt){
  var userProperties = PropertiesService.getUserProperties();
  var body = DocumentApp.getActiveDocument().getBody();
  var cell = [['name','powers\ndescription'],['','information about them as a person']];
  var table = body.appendTable(cell);
  var cell1 = table.getCell(0,0);
  cell1.setVerticalAlignment(DocumentApp.VerticalAlignment.CENTER);
  cell1.getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  cell1.editAsText().setBold(0,3,true).setFontSize(11).editAsText().setFontFamily(userProperties.getProperty('tFont'));
  
  cell1.setWidth(50);
  var cell12 = table.getCell(1,1); //merge
  var cell11 = table.getCell(1,0); //merge
  cell12.getChild(0).editAsText().setFontFamily(userProperties.getProperty('tFont'));
  
  var cell03 = table.getCell(0,1);  
  cell03.getChild(0).editAsText().setFontFamily(userProperties.getProperty('tFont'));
  var pow = table.findText('powers');
  pow.getElement().asText().setBold(0,5,true);
  var description = table.findText('description').getElement().asText().setItalic(0,10,true).setBold(false);
  var cell02 = table.getCell(0,1);
  cell02.setBackgroundColor('#dddddd');
  
  if(isPrt == true){
    cell1.setBackgroundColor('#efefef');
  } else {
    cell1.setBackgroundColor(userProperties.getProperty('color'));
  }
}
//-- nHFont,hTFont,colorType,city,month,day,year,area --//
function changeThings(a,b,c,d,e,f,g,h){
  var userProperties = PropertiesService.getUserProperties();
  if(a!=''){
    userProperties.setProperty('hFont', a);
  }
  if(b!=''){
    userProperties.setProperty('tFont', b);
  }
  if(c!=''){
    userProperties.setProperty('color', c);
  }
  if(d!=''){
    userProperties.setProperty('city', d);
  }
  if(e!=''){
    userProperties.setProperty('month', e);
  }
  if(f!=''){
    userProperties.setProperty('day', f);
  }
  if(g!=''){
    userProperties.setProperty('year', g);
  }
  if(h!=''){
    userProperties.setProperty('area', h);
  }
  
}
//-- Menu-specific Functions --//
function addSection(){
  var userProperties = PropertiesService.getUserProperties();
  var body = DocumentApp.getActiveDocument().getBody();
  addGeneratedSection("section","subtitle",body,-1)
}
function pers(){
  addPers(false);
}
function sec(){
  addLog('sec');
}
function pub(){
  addLog('pub');
}
function pho(){
  addLog('pho');
}
function norm(){
  addLog('norm');
}
function date(){
  var userProperties = PropertiesService.getUserProperties();
  var body = DocumentApp.getActiveDocument().getBody();
  var date = body.appendParagraph(
    userProperties.getProperty('month')+' '+
    userProperties.getProperty('day')+', '+
    userProperties.getProperty('year'));
  var dateText = date.editAsText();
  dateText.setBold(0,dateText.getText().length-1,true).setFontFamily(userProperties.getProperty('tFont'));
}
//-- End menu functions --//

//-- Menu sidebars --//
function customizationsSideBar(){
  var html = HtmlService.createHtmlOutputFromFile('custos')    
    .setTitle('Customizations')
    .setWidth(600);
  DocumentApp.getUi()
     .showSidebar(html);
}

function infoSideBar(){
  var html = HtmlService.createHtmlOutputFromFile('infoCusto')
    .setTitle('Information About This Add-On')
    .setWidth(350);
  DocumentApp.getUi()
    .showSidebar(html)
}

function createCampaignDoc(){
  var html = HtmlService.createHtmlOutputFromFile('creationPage')    
    .setTitle('Document Generator')
    .setWidth(600);
  DocumentApp.getUi()
     .showSidebar(html);
}

