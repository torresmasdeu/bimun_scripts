/** CREACIÓ DEL MENÚ */  
function open(){
    var ui = SpreadsheetApp.getUi();
    ui.createMenu('Committee allocations')
        .addItem('Make Committee Allocations','allocations')
            
        .addToUi();
}

/** OMPLIR CA SPREADSHEET AMB DDF DETAILS I MARCAR PAÏSOS AGAFATS A LIST OF COUNTRIES */  
var ui = SpreadsheetApp.getUi();
var ca_ss_id = '1hKG1Pgcmv4yiJIp1FlNquASFq14TgsqUXZO2Hi5fWQs';
var ca_ss = SpreadsheetApp.openById(ca_ss_id);
var ca_sheet = ca_ss.getSheets()[0];

var ddf_folder_id = '1EFnA_Gp80knz_UDPADpQ_ms6UwxMwCmm' //(My Drive > DOCUMENTS > DDF)
var ddf_folder = DriveApp.getFolderById(ddf_folder_id)
var contents = ddf_folder.getFiles()

var loc_id = '1R7sdiBryzruQrVM2wXgdyzCpsN8ILMOf7vsIaI9rUmk'
var loc_doc = DocumentApp.openById(loc_id)
var loc_body = loc_doc.getBody()
var tables = loc_body.getTables()

function allocations() {
  var school_name_result = '';
  var school_name = '1';

  /**Nom del cole que volem fer el PD */
  while (school_name != school_name_result) {
    var result = ui.prompt(
      'School Name',
      'Please paste the complete name of the school (just as it is stored in the "Schools" document):',
      ui.ButtonSet.OK);
    if (result.getSelectedButton() == ui.Button.OK) {
      school_name = result.getResponseText();
      var school_name_file = school_name.replaceAll(' ','_') +'_DDF'

      //look for school_name_file in the DDF folder
      while (contents.hasNext()){
        var file = contents.next();
        var file_name = file.getName();
        if (file_name == school_name_file){
          school_name_result = school_name;
          var ddf_file = file;
          break
        }
        else if (contents.hasNext() == false){
          ui.alert('Incorrect name. Please try again')
        }
      }
    }
    else {
      return
    }
  }

  var ddf_ss = SpreadsheetApp.open(ddf_file)
  var ddf_sheet = ddf_ss.getSheets()[0];

  var missing = [] //array that stores the name of those delegates whose CA could not be done

  for (var ddf_row = 5; ddf_row <= ddf_sheet.getLastRow(); ddf_row++){

    allocation: 
    if (ddf_sheet.getRange(ddf_row,8).getValue() != ''){
  
      for (var allocation_option = 8; allocation_option <=12; allocation_option ++){
        var name = ddf_sheet.getRange(ddf_row,1).getValue() + ' ' + ddf_sheet.getRange(ddf_row,2).getValue() //delegate's name
        var email = ddf_sheet.getRange(ddf_row,14).getValue() //delegate's email
        var country = ddf_sheet.getRange(ddf_row,allocation_option).getValue()
        var committee = ddf_sheet.getRange(ddf_row,allocation_option+1).getValue()

        if (country != ''){var country_list = ca_sheet.createTextFinder(country).findAll()} //look for country in excel
        else {missing = missing.concat(name); break allocation;} //no information found for this delegate (cell was empty)

        for (var r=0; r <country_list.length; r++){
          var country_row = country_list[r].getRow() //return row number where country is found

          if ((ca_sheet.getRange(country_row,5).getValue() == committee) && (ca_sheet.getRange(country_row,4).getValue() == '')){
            ca_sheet.getRange(country_row,1).setValue(name)
            ca_sheet.getRange(country_row,2).setValue(email)
            ca_sheet.getRange(country_row,4).setValue(school_name)

            var searchType = DocumentApp.ElementType.PARAGRAPH;
            var searchHeading = DocumentApp.ParagraphHeading.HEADING1;
            var searchResult = null;

            var table_index = 0

            //search the committee name in Country List doc until the paragraph containing it is found
            searchparagraph:
            while (searchResult = loc_body.findElement(searchType, searchResult)) {
              var par = searchResult.getElement().asParagraph();
              if (par.getHeading() == searchHeading){
                if (par.getText() == committee) {break searchparagraph;}
                table_index = table_index + 1 //table number
              }
            }

            var table = tables[table_index] //retrieve table for desire committee

            var loc_n_rows = table.getNumRows()
            var loc_n_col = table.getRow(0).getNumCells()

            //go through each cell from the table and look for the country name
            findcell:
            for (var loc_row = 0; loc_row < loc_n_rows; loc_row ++){
              for (var loc_col = 0; loc_col < loc_n_col; loc_col ++){
                if (table.getCell(loc_row,loc_col).getText() == country) {
                  table.getCell(loc_row,loc_col).setBackgroundColor('#ea9999')
                  break findcell;
                }
              }
            }
            break allocation;
          }
        }
        allocation_option = allocation_option + 1

        if (allocation_option == 13){missing = missing.concat(name)} //variable containing names of delegates with no CA
      }
    }
  }
}

/** POSAR TOTES LES CELLS EN BLANC, I EN GROC ELS PAÏSOS DIFÍCILS*/  

function allwhite() {
    var loc_id = '1R7sdiBryzruQrVM2wXgdyzCpsN8ILMOf7vsIaI9rUmk'
    var loc_doc = DocumentApp.openById(loc_id)
    var loc_body = loc_doc.getBody()
    var tables = loc_body.getTables()
  
    for (i=0; i<tables.length;i++){
      table = tables[i]
  
      if (i==0){}
  
      var loc_n_rows = table.getNumRows()
      var loc_n_col = table.getRow(0).getNumCells()
  
      for (var loc_row = 0; loc_row < loc_n_rows; loc_row ++){
        for (var loc_col = 0; loc_col < loc_n_col; loc_col ++){
          if (i==0){
            table.getCell(loc_row,loc_col).setBackgroundColor('#f9f97c')
          }
          else {
            if (
              table.getCell(loc_row,loc_col).getText()=='UK'||
              table.getCell(loc_row,loc_col).getText()=='USA'||
              table.getCell(loc_row,loc_col).getText()=='France'||
              table.getCell(loc_row,loc_col).getText()=='Russian Federation'||
              table.getCell(loc_row,loc_col).getText()=='China'){
                table.getCell(loc_row,loc_col).setBackgroundColor('#f9f97c')
            }
            else {
              table.getCell(loc_row,loc_col).setBackgroundColor('#FFFFFF')
            }
          }
        }   
      }
    } 
}