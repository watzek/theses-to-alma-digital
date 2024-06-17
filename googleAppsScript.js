// Alma Digital Sub-Collection Codes
let codes={};
codes["Art History"]="81244791060001844";
codes["Asian Studies"]="81256481040001844";
codes["Biology"]="81256481090001844";
codes["Biochemistry / Molecular Biology"]="81256481090001844";
codes["Chemistry"]="81257650760001844";
codes["Classics"]="81257631000001844";
codes["Computer Science"]="81257630990001844";
codes["Economics"]="81257630980001844";
codes["English"]="81257630970001844";
codes["Environmental Studies"]="81257630960001844";
codes["French"]="81257630950001844";
codes["Gender Studies"]="81257630940001844";
codes["German"]="81257630930001844";
codes["History"]="81257630920001844";
codes["International Affairs"]="81257630910001844";
codes["Mathematics"]="81257630900001844";
codes["Music"]="81257630890001844";
codes["Philosophy"]="81257630880001844";
codes["Physics"]="81257630870001844";
codes["Political Science"]="81257630860001844";
codes["Psychology"]="81257630850001844";
codes["Religious Studies"]="81257630840001844";
codes["Rhetoric and Media Studies"]="81257630830001844";
codes["Sociology and Anthropology"]="81257630820001844";
codes["Spanish"]="81257630810001844";
codes["Theatre"]="81257630800001844";
codes["World Languages and Literatures"]="81257630790001844";




//creates menu
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createMenu('Thesis Functions')
  .addItem('Prepare metadata for xml', 'prepareForXml')
  .addSeparator()
  .addItem('Generate Marc XML file', 'generateXml')
  .addSeparator()
  .addItem('Download Zip for Alma Ingest', 'downloadFolderAsZip')
  .addToUi();

}


function prepareForXml(){
  // get all sheet values
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var values = sheet.getDataRange().getValues();

  // add column headers
  addHeaders(sheet);

  // loop through rows
  values.forEach( function(row, index) {
    if(index>0){ //skips header row

      var firstName=row[2];
      var lastName=row[3];
      var year=row[6];
      var fileUrl=row[9];
      var visibility=row[10];
      var dept=row[11];
      var advisor=row[13];

      updateFiles(sheet, index, dept, firstName, lastName, fileUrl, year);

      setAccessPolicy(sheet, index, visibility);

      setCellCollectionCode(sheet, index, dept);

      setFaculty(sheet, index, advisor);


    }
  });

}


function addHeaders(sheet){
    //16 is P
  //Set column header
  var newFileHeader = sheet.getRange(1,16);
  newFileHeader.setValue("renamedFile");
  //17 is Q
  //Set column header
  var accessPolicyHeader = sheet.getRange(1,17);
  accessPolicyHeader.setValue("accessPolicy");
  //18 is R
  //Set column header
  var collectionCodeHeader = sheet.getRange(1,18);
  collectionCodeHeader.setValue("collectionCode");

  //19 is S
  //Set column header
  var facultyHeader = sheet.getRange(1,19);
  facultyHeader.setValue("faculty");

}

function setAccessPolicy(sheet, index, visibility){

      if(visibility=="LC - only"){var newVal="CAS access";}
      if(visibility=="Public"){var newVal="Open to the world";}

      // update spreadsheet with new file name    
      var cell = sheet.getRange(index+1,17);
      cell.setValue(newVal);
}

function setCellCollectionCode(sheet, index, dept){
      collectionCode=codes[dept];
      Logger.log(collectionCode);
      //update spreadsheet with collection code    
      var cell = sheet.getRange(index+1,18);
      cell.setValue(collectionCode);

}

function setFaculty(sheet, index, advisor){
  //example: Bostian, Moriah - mbbostian@lclark.edu
  var pieces=advisor.split(" - ");
  var name=pieces[0];
  var namePieces=name.split(", ");
  var facultyName=namePieces[1]+" "+namePieces[0];
  var cell = sheet.getRange(index+1,19);
  cell.setValue(facultyName);

}

function updateFiles(sheet, index, dept, firstName, lastName, fileUrl, year){

      var fi=firstName[0].toLowerCase();
      var ln = lastName.replace(/\s/g, '').toLowerCase();
      var dept=dept.replace(/\s/g, '').toLowerCase();

      //naming convention: holzer-k-Biology-2008-0001.pdf
      var fileName=ln+"-"+fi+"-"+dept+"-"+year+"-0001.pdf";

      var filePieces=fileUrl.split("?id=");
      var fileId=filePieces[1];

      // update spreadsheet with new file name    
      var cell = sheet.getRange(index+1,16);
      cell.setValue(fileName);

      //rename the file in G-Drive
      DriveApp.getFileById(fileId).setName(fileName);
      getContainingFolderId(fileId);

      Logger.log(index);
      Logger.log(fileName); // column index as 4


}

function getEight(){
  //format control field 008
  var yy=new Date().toLocaleString("en-US", {year: '2-digit'});
  var mm=new Date().toLocaleString("en-US", {month: '2-digit'});
  var dd=new Date().toLocaleString("en-US", {day: '2-digit'});
  var yyyy=new Date().getFullYear();
  var yymmdd=yy+mm+dd;
  Logger.log(yymmdd);
  Logger.log(yyyy);

  var ohOhEight=yymmdd+"s"+yyyy+"    xx a     bm   000 0 eng d";
  Logger.log(ohOhEight)
  return ohOhEight;


}

function generateXml() {
  // Define variables for each element's value
  var leader = "02470ctm a2200469 a 4500";
  var controlfield008=getEight();
  //var controlfield008 = "151009s2019    xx a     bm   000 0 eng d";
  var controlfield001 = "";

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var values = sheet.getDataRange().getValues();


  // Create XML structure
  var xml = '<?xml version="1.0" encoding="UTF-8" ?>\n';
  xml += '<marc:collection xmlns:marc="http://www.loc.gov/MARC21/slim" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:schemaLocation="http://www.loc.gov/MARC21/slim http://www.loc.gov/standards/marcxml...MARC21slim.xsd">\n';

  // loop through rows to create marc record for each row
  values.forEach( function(row, index) {
    if(index>0){ //skips header row

      var firstName=row[2];
      var lastName=row[3];
      var author=lastName+", "+firstName
      var year=row[6];
      var title=row[7].replace(/&/g, '&amp;');
      var abstract=row[8].replace(/&/g, '&amp;');
      var advisorName=row[18];
      var advisor="Advisor: "+advisorName;
      var fileUrl=row[9];
      var visibility=row[10];
      var dept=row[11].replace(/&/g, '&amp;');
      var degree=row[12];
      //var advisor=row[13];

      var filePath=row[15];
      var accessPolicy=row[16];
      var collectionCode=row[17];


      xml += '  <!-- BEGIN One Bib, One Rep, One File -->\n';
      xml += '  <marc:record>\n';
      xml += `    <marc:leader>${leader}</marc:leader>\n`;
      xml += `    <marc:controlfield tag="008">${controlfield008}</marc:controlfield>\n`;
      xml += `    <marc:controlfield tag="001">${controlfield001}</marc:controlfield>\n`;
      xml += '    <marc:datafield tag="100" ind1="1" ind2=" ">\n';
      xml += `      <marc:subfield code="a">${author}</marc:subfield>\n`;
      xml += '    </marc:datafield>\n';
      xml += '    <marc:datafield tag="245" ind1="1" ind2="0">\n';
      xml += `      <marc:subfield code="a">${title}</marc:subfield>\n`;
      xml += '    </marc:datafield>\n';
      xml += '    <marc:datafield tag="500" ind1="1" ind2="0">\n'; 
      xml += `      <marc:subfield code="a">${advisor}</marc:subfield>\n`;
      xml += '    </marc:datafield>\n';
      xml += '    <marc:datafield tag="502" ind1="1" ind2="0">\n'; 
      xml += `      <marc:subfield code="b">B.A.</marc:subfield>\n`;
      xml += `      <marc:subfield code="c">Lewis &amp; Clark College</marc:subfield>\n`;
      xml += `      <marc:subfield code="d">${year}</marc:subfield>\n`;
      xml += `.     <marc:subfield code="g">${dept}</marc:subfield>`
      xml += '    </marc:datafield>\n';
      xml += '    <marc:datafield tag="520" ind1=" " ind2=" ">\n';
      xml += `      <marc:subfield code="a">${abstract}</marc:subfield>\n`;
      xml += '    </marc:datafield>\n';
      xml += '    <marc:datafield tag="540" ind1=" " ind2=" ">\n';
      xml += `      <marc:subfield code="a">${accessPolicy}</marc:subfield>\n`;
      xml += '    </marc:datafield>';
      xml += '    <marc:datafield tag="655" ind1="0" ind2=" ">\n'; 
      xml += `      <marc:subfield code="a">Academic Thesis</marc:subfield>\n`;
      xml += '    </marc:datafield>\n';
      xml += '    <marc:datafield tag="787" ind1="0" ind2=" ">\n';
      xml += `      <marc:subfield code="w">${collectionCode}</marc:subfield>\n`;
      xml += '    </marc:datafield>\n';
      xml += '    <marc:datafield tag="856" ind1="1" ind2="2">\n';
      xml += `      <marc:subfield code="u">${filePath}</marc:subfield>\n`;
      xml += `      <marc:subfield code="y">View Thesis</marc:subfield>\n`;
      xml += '    </marc:datafield>\n';
      xml += '  </marc:record>\n';
      xml += '  <!-- END One Bib, One Rep, One File -->\n';
      

    }
  });

    xml += '</marc:collection>';






/*
  var author = "Kennedy, Dakota";
  var title = "Synthesis and Investigation of Surface-Bound Gold and Silver Nanoparticles";
  var advisor = "Advisor: Anne Bentley";
  var degree = "B.A.";
  var institution = "Lewis & Clark";
  var year = "2019";
  var abstract = "Abstract: Recent advances in the understanding of nanomaterials have led to a rapid expansion in the number of applications for these materials. For gold and silver nanomaterials, a variety of uses have been found in art, catalysis, and electronics but a large and rapidly increasing number of applications are being found for these precious metals as nanomaterials in biomedicine and diagnostic biosensing. As the use of these particles becomes more and more prevalent, the amount of nanoparticle-containing waste that is generated is rising. Despite this, little is known about how these particles interact with biological environments after being discarded, and more specifically, how these particles can interact with each other. By observing changes in localized surface plasmon resonance (LSPR) behavior, it is known that some sort of interaction occurs between aqueous gold and silver nanoparticles under ambient conditions but exactly what and how it occurs is a mystery. Here, I report that the LSPR behavior that is observed in a solution composed of both aqueous gold and silver nanoparticles is incongruous with the behavior shown when one of the nanoparticle species is immobilized on a substrate in the absence of a capping agent. I also report a simple and novel synthetic strategy for the electrochemical growth and immobilization of gold nanoparticles on a substrate and furthermore, report a similar, facile synthesis for the electrochemical growth and immobilization of silver-containing nanoparticles on a substrate.";
  var academicTheses = "Academic theses.";
  var recordId = "81257650760001844";
  var filePath = "kennedy-d-chemistry-2019-0001.pdf";

*/



  // Log the XML (or do something else with it, like saving to a file)
  Logger.log(xml);

  var folderId=getFolderId();
  Logger.log(xml);

  // Get the folder by ID
  var folder = DriveApp.getFolderById(folderId);

  // Create the XML file in the specified folder
  var fileName = 'output.xml';
  var file = folder.createFile(fileName, xml, MimeType.PLAIN_TEXT);
  
}









function getFolderId(){
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var values = sheet.getDataRange().getValues();
  var cell = sheet.getRange(2,10);
  fileUrl=cell.getValue();
  Logger.log(fileUrl);
  var filePieces=fileUrl.split("?id=");
  var fileId=filePieces[1];

      
      //DriveApp.getFileById(fileId).setName(fileName);
  folderId=getContainingFolderId(fileId);
  return folderId;

}





function getCollectionCode(){

  // get all sheet values
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var values = sheet.getDataRange().getValues();
  //18 is R
  //Set column header
  var newHeader = sheet.getRange(1,18);
  newHeader.setValue("collectionCode");
    values.forEach( function(row, index) {
    if(index>0){ //skips header row



      var dept=row[11];
      collectionCode=codes[dept];
      Logger.log(collectionCode);



      //update spreadsheet with collection code    
      var cell = sheet.getRange(index+1,18);
      cell.setValue(collectionCode);


    }
  });


}





  //update access rights to match Alma Access Rights Policies
function setAccess(){

  // get all sheet values
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var values = sheet.getDataRange().getValues();
  //17 is Q
  //Set column header
  var newHeader = sheet.getRange(1,17);
  newHeader.setValue("accessPolicy");

  values.forEach( function(row, index) {
    if(index>0){ //skips header row
      //naming convention: holzer-k-Biology-2008-0001.pdf


      var visibility=row[10];
      if(visibility=="LC - only"){var newVal="CAS access";}
      if(visibility=="Public"){var newVal="Open to the world";}


      // update spreadsheet with new file name    
      var cell = sheet.getRange(index+1,17);
      cell.setValue(newVal);


    }
  });




}


function renameFiles(){

  // get all sheet values
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var values = sheet.getDataRange().getValues();
  //16 is P
  //Set column header
  var newHeader = sheet.getRange(1,16);
  newHeader.setValue("renamedFile");


  values.forEach( function(row, index) {
    if(index>0){ //skips header row
      

      
      var firstName=row[2];
      var lastName=row[3];
      var year=row[6];
      var fileUrl=row[9];
      var visibility=row[10];
      var dept=row[11];

      var fi=firstName[0].toLowerCase();
      var ln = lastName.replace(/\s/g, '').toLowerCase();
      var dept=dept.replace(/\s/g, '').toLowerCase();

      //naming convention: holzer-k-Biology-2008-0001.pdf
      var fileName=ln+"-"+fi+"-"+dept+"-"+year+"-0001.pdf";

      var filePieces=fileUrl.split("?id=");
      var fileId=filePieces[1];

      // update spreadsheet with new file name    
      var cell = sheet.getRange(index+1,16);
      cell.setValue(fileName);

      //rename the file in G-Drive
      DriveApp.getFileById(fileId).setName(fileName);
      getContainingFolderId(fileId);

      Logger.log(index);
      Logger.log(fileName); // column index as 4
    }
  });



}

function getContainingFolderId(fileId) {
  // Get the file by ID
  var file = DriveApp.getFileById(fileId);
  
  // Get the parents of the file (files can have multiple parents)
  var parents = file.getParents();
  
  // Assuming the file has at least one parent, get the first parent folder
  if (parents.hasNext()) {
    var parent = parents.next();
    var parentId = parent.getId();
    Logger.log('Parent Folder ID: ' + parentId);
    return parentId;
  } else {
    Logger.log('The file is in the root folder or has no parent folders.');
    return null;
  }
}



function downloadFolderAsZip() {
  //var folderId = 'YOUR_FOLDER_ID';  // Replace with your folder ID
  var folderId=getFolderId();
  var folder = DriveApp.getFolderById(folderId);
  var zipBlob = createZipBlob(folder);
  
  // Create a temporary file in Drive
  var tempFile = DriveApp.createFile(zipBlob);
  
  // Set the permissions to make the file accessible
  tempFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  
  // Get the URL for downloading the file
  var downloadUrl = tempFile.getUrl();
  
  // Show the download link in an HTML dialog box
  var htmlOutput = HtmlService.createHtmlOutput('<p>Download the ZIP file from this link: <a href="' + downloadUrl + '" target="_blank">' + downloadUrl + '</a></p>')
      .setWidth(400)
      .setHeight(200);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Download ZIP');
}

function createZipBlob(folder) {
  var zip = Utilities.newBlob('');
  var files = folder.getFiles();
  var blobs = [];
  
  while (files.hasNext()) {
    var file = files.next();
    blobs.push(file.getBlob());
  }
  
  zip = Utilities.zip(blobs, folder.getName() + '.zip');
  return zip;
}






