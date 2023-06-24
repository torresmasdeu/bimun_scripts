var ui = SpreadsheetApp.getUi();
var schools_ss_id = '1b2BWKdnBoTC7DqPE0OjbDH6-2Q3_sF3oM1fRhNm7i8s'
var schools_ss = SpreadsheetApp.openById(schools_ss_id);
var allinc_sheet = schools_ss.getSheetByName('all inc');
var allinc_p = schools_ss.getSheetByName('all inc payment');
var standard_sheet = schools_ss.getSheetByName('standard');
var standard_p = schools_ss.getSheetByName('standard payment');
var pd_folder_id = '1_bYVHa3k84dBOSsfs1xBwaJo5zR56fCN' //(My Drive > DOCUMENTS > Payments > Payment documents)
var pd_folder = DriveApp.getFolderById(pd_folder_id);
var ood_folder_id = '1NP2jk7a0u23jogiYxrqRv65Ux1y8oags' //(My Drive > DOCUMENTS > Payments > Payment documents > Out-of-Date PD)
var ood_folder = DriveApp.getFolderById(ood_folder_id)

var aiID = '1pgqtheNLbCzgwgClIPVlx-_gml2hmACy' //AI information PDF document

var allinc_ddf_template_id = '1ckMUlakSnHOWsxukeTSg0HzryEvlZ9mRSdQiaGDscYU'
var allinc_ddf_template = SpreadsheetApp.openById(allinc_ddf_template_id)
var standard_ddf_template_id = '1TYcomq-pTgls7uUAPx-0B6mYWfuKM-HsnOyVd0htdUM'
var standard_ddf_template = SpreadsheetApp.openById(standard_ddf_template_id)
var ddf_folder_id = '1EFnA_Gp80knz_UDPADpQ_ms6UwxMwCmm' //(My Drive > DOCUMENTS > DDF)

var signature = Gmail.Users.Settings.SendAs.get("me", "bimun@barcelonamun.com").signature;

function get_cells_and_vals(){
  //Function that looks for the desired 
  /**Nom del cole que volem fer el PD*/
  while (true) {
    var result = ui.prompt(
      'School Name',
      'Please paste the complete name of the school (just as it is stored in this document):',
      ui.ButtonSet.OK);
    if (result.getSelectedButton() == ui.Button.OK) {
      var school_name = result.getResponseText();
      
      
      for (var s = 0; s < s_sheets.length; s++){
        var s_sheet = s_sheets[s]
        for (var r = 2; r <= s_sheet.getLastRow();r++) {
          
          if (school_name == s_sheet.getRange(r,2).getValue()){
            var school_row = r;
          }
            var school_number = s_sheet.getRange(school_row,1).getValue();
            var teacher_name = s_sheet.getRange(school_row,3).getValue();  
            var emailaddress = s_sheet.getRange(school_row,4).getValue();
            var n_delegates = s_sheet.getRange(school_row,5).getValue();
            var n_teachers = s_sheet.getRange(school_row,6).getValue();

            if (s_sheet == allinc_sheet){
              var n_triples_teach = allinc_sheet.getRange(school_row,7).getValue();
              var n_doubles_teach = allinc_sheet.getRange(school_row,8).getValue();
              var n_singles_teach = allinc_sheet.getRange(school_row,9).getValue();
              var n_triples_del = allinc_sheet.getRange(school_row,10).getValue();
              var n_doubles_del = allinc_sheet.getRange(school_row,11).getValue();
              var n_singles_del = allinc_sheet.getRange(school_row,12).getValue();
              var extranights = allinc_sheet.getRange(school_row,13).getValue();
              var social = allinc_sheet.getRange(school_row,14).getValue();
              var pd_number_cell = allinc_sheet.getRange(school_row,15);
              var pd_cell = allinc_sheet.getRange(school_row,16);
              var pdrec_cell = allinc_sheet.getRange(school_row,17);
              var DDFrec_cell = allinc_sheet.getRange(school_row,18);
              var CA_cell = allinc_sheet.getRange(school_row,19);
              var CAsent_cell = allinc_sheet.getRange(school_row,20);
              var obs_cell = allinc_sheet.getRange(school_row,21);
            }
            else{ //standard
              var social = standard_sheet.getRange(school_row,7).getValue();
              var pd_number_cell = standard_sheet.getRange(school_row,8).getValue();
            }






            return {school_name, school_row, app_sheet}
          }
        }
      }
    }
    else {
      return '';
    }
  }
}

function hola(){
  let {school_name, school_row, app_sheet} = get_cells()
  console.log(school_name, school_row, app_sheet)
  if (school_name == undefined || school_row == undefined){return}

  //assignar valors de cada cel·la a cada variable 
  var school_number = allinc_sheet.getRange(school_row,1).getValue();
  var teacher_name = allinc_sheet.getRange(school_row,3).getValue();  
  var emailaddress = allinc_sheet.getRange(school_row,4).getValue();
  var n_delegates = allinc_sheet.getRange(school_row,5).getValue();
  var n_teachers = allinc_sheet.getRange(school_row,6).getValue();
  var n_triples_teach = allinc_sheet.getRange(school_row,7).getValue();
  var n_doubles_teach = allinc_sheet.getRange(school_row,8).getValue();
  var n_singles_teach = allinc_sheet.getRange(school_row,9).getValue();
  var n_triples_del = allinc_sheet.getRange(school_row,10).getValue();
  var n_doubles_del = allinc_sheet.getRange(school_row,11).getValue();
  var n_singles_del = allinc_sheet.getRange(school_row,12).getValue();
  var extranights = allinc_sheet.getRange(school_row,13).getValue();
  var social = allinc_sheet.getRange(school_row,14).getValue();
  var pd_number = allinc_sheet.getRange(school_row,15).getValue();
  var duedate = String(allinc_p.getRange('H12').getDisplayValue());
  
  if (teacher_name.toString() == ''||emailaddress.toString() == ''||n_delegates.toString() == ''||n_teachers.toString() == ''||n_triples_teach.toString() == ''||n_doubles_teach.toString() == ''||n_singles_teach.toString() == ''||n_triples_del.toString() == ''||n_doubles_del.toString() == ''||n_singles_del.toString() == ''||extranights.toString() == ''||social.toString() == '' || (teacher_name.toString().split(' ') [0] != 'Mr' && teacher_name.toString().split(' ') [0] != 'Ms' && teacher_name.toString().split(' ') [0] != 'Mrs')){
    for (var e = 3; e<15; e++){
      if ((allinc_sheet.getRange(school_row,e).getValue().toString()=='')){
        if (7<=e && e<=9){
          missing = 'Number of teacher rooms: ' + allinc_sheet.getRange(2,e).getValue();
        }
        else if (10<=e && e<=12){
          missing = 'Number of delegate rooms: ' + allinc_sheet.getRange(2,e).getValue();
        }
        else{
          missing = allinc_sheet.getRange(1,e).getValue();
        }
        ui.alert('You are missing "'+ missing +'" information. Please fill it in and try again.');
      }
    }
    if (teacher_name.toString().split(' ') [0] != 'Mr' && teacher_name.toString().split(' ') [0] != 'Ms' && teacher_name.toString().split(' ') [0] != 'Mrs'){
        ui.alert("You are missing information on the teacher's title (Mr, Ms or Mrs). Please fill it in and try again.");
      }
    return
  }

  /**Comprovar que el nombre de participants correspon al nombre d'habitacions reservada*/
  var n_triples = n_triples_del + n_triples_teach;
  var n_doubles = n_doubles_del + n_doubles_teach;
  var n_singles = n_singles_del + n_singles_teach;

  if (((n_triples*3)+(n_doubles*2)+n_singles) != (n_teachers + n_delegates)){
    var res = ui.alert(
      'Number of rooms',
      'The number of booked rooms does not correspond to the number of participants. Do you want to send an email specifying some changes?',
      ui.ButtonSet.OK);
    if (res == ui.Button.OK){
      ai_pd_wrongrooms ();
    }
    return
  }

  //determinar Version number
  if (pd_number == ''){
    var version_n = '01'
  }
  else{
    var pd_number_digits = (""+pd_number).split(""); //passar de string a array (0,X,2,3,Y,Y)
    console.log (++(pd_number_digits[1])) //afegir 1 al segon element de l'array (X, que és el número de versió)
    var version_n = pd_number_digits [0] + pd_number_digits [1]

    //canviar antic pdf document a la carpeta OOD
    var old_pd_number = pd_number;
    var old_pdfName = school_name.replaceAll(' ','_') + '_' +old_pd_number;
    const old_file = pd_folder.getFilesByName(old_pdfName+'.pdf').next();
    old_file.moveTo(ood_folder);
  }

  /**Formatejar el Payment document*/
  //Omplir Payment document amb les dades de allinc sheet
  allinc_p.getRange('J9').setValue(school_number);
  allinc_p.getRange('B12').setValue(version_n);
  pd_number = allinc_p.getRange('H9').getValue();

  for (var w = 1; w<1000000; w++){
    if (allinc_p.getRange('J9').getValue() == school_number && allinc_p.getRange('B12').getValue()==version_n){
      break
    }
  };  

  var pdfName = school_name.replaceAll(' ','_') + '_' +pd_number;

  /**Generar el Payment document*/
  MakePDF(allinc_p,pdfName)
    
  /**Enviar email amb PDF attached */
  SendEmail('AI_PD_email',teacher_name, school_name, n_delegates, n_teachers,'','','','','','','','','','', n_triples, n_doubles, n_singles, n_triples_del, n_doubles_del, n_singles_del, n_triples_teach, n_doubles_teach, n_singles_teach, extranights, duedate, social, version_n, school_row, emailaddress, pdfName,'')  

  ui.alert('Check the "Drafts" section of GMAIL and change the default signature for yours. If all is correct, send the email.')

  allinc_sheet.getRange(school_row,15).setValue(pd_number);
  allinc_sheet.getRange(school_row,16).setValue('Yes');
  allinc_sheet.getRange(school_row,7,1,6).setBackground('#b7e1cd');

  var date = String(allinc_p.getRange('D12').getDisplayValue());
  allinc_sheet.getRange(school_row,21).setValue('Payment document '+pd_number+ '. Sent on '+ date);

  return

  function ai_pd_wrongrooms () {
    var real_n_participants = n_teachers + n_delegates;
    var real_n_teachers = n_teachers;
    var real_n_delegates = n_delegates;

    var app_n_participants = (n_triples*3)+(n_doubles*2)+n_singles;
    var app_n_teachers = (n_triples_teach*3)+(n_doubles_teach*2)+n_singles_teach;
    var app_n_delegates = (n_triples_del*3)+(n_doubles_del*2)+n_singles_del;

    var app_n_triples = n_triples;
    var app_n_doubles = n_doubles;
    var app_n_singles = n_singles;

    var app_n_triples_teach = n_triples_teach;
    var app_n_doubles_teach = n_doubles_teach;
    var app_n_singles_teach = n_singles_teach;

    var app_n_triples_del = n_triples_del;
    var app_n_doubles_del = n_doubles_del;
    var app_n_singles_del = n_singles_del; 

    if (app_n_teachers != n_teachers){
      if (app_n_teachers < real_n_teachers){
        var dif_t = real_n_teachers - app_n_teachers;
        if (dif_t == 1){
          if (n_doubles_teach > 0){
            n_triples_teach = n_triples_teach + 1; //+3
            n_doubles_teach = n_doubles_teach - 1; //-2
          }
          else if (n_triples_teach > 0){
            n_triples_teach = n_triples_teach - 1; //-3
            n_doubles_teach = n_doubles_teach + 2; //+4
          }
          else {
            n_singles_teach = n_singles_teach + 1;
          }
        }
        else if (dif_t == 2){
          n_doubles_teach = n_doubles_teach + 1;
        }
        else if (dif_t == 3){
          n_triples_teach = n_triples_teach + 1;
        }
        else if (dif_t == 4){
          if (n_doubles_teach > 0){
            n_triples_teach = n_triples_teach + 2; //+6
            n_doubles_teach = n_doubles_teach - 1; //-2
          }
          else {
            n_doubles_teach = n_doubles_teach + 2
          }
        }
      }

      else {
        var dif_t = app_n_teachers - real_n_teachers;
        if (dif_t == 1){
          if (n_triples_teach > 0) {
            n_triples_teach = n_triples_teach - 1; //-3
            n_doubles_teach = n_doubles_teach + 1; //+2
          }
          else if (n_doubles_teach > 0) {
            n_triples_teach = n_triples_teach + 1; //+3
            n_doubles_teach = n_doubles_teach - 2; //-4
          }
        }
        else if (dif_t == 2){
          if (n_doubles_teach > 0){
            n_doubles_teach = n_doubles_teach - 1;
          }
          else if (n_triples_teach >= 2) {
            n_triples_teach = n_triples_teach -2; //-6
            n_doubles_teach = n_doubles_teach + 2; //+4
          }
          else if (n_triples_teach > 0) {
            n_triples_teach = n_triples_teach - 1; //-3
            n_singles_teach = n_singles_teach + 1; //+1
          }
          else {
            n_singles_teach = n_singles_teach - 2;
          }
        }
        else if (dif_t == 3){
          if (n_triples_teach > 0){
            n_triples_teach = n_triples_teach - 1;
          }
          else if (n_doubles_teach >= 3){
          n_doubles_teach = n_doubles_teach - 3; //-6
          n_triples_teach = n_triples_teach + 1; //+3
          }
          else if (n_doubles_teach >= 2){
          n_doubles_teach = n_doubles_teach - 2; //-4
          n_singles_teach = n_singles_teach + 1; //+1
          }
        }
        else if (dif_t == 4){
          if (n_doubles_teach >= 2){
            n_doubles_teach = n_doubles_teach - 2;
          }
          else if (n_triples_teach >= 3){
            n_triples_teach = n_triples_teach - 3; //-6
            n_doubles_teach = n_doubles_teach + 1; //+2
          }
        }
      }
    }

    if (app_n_delegates != n_delegates){
      if (app_n_delegates < real_n_delegates){
        var dif_d = real_n_delegates - app_n_delegates;
        if (dif_d == 1){
          if (n_doubles_del > 0){
            n_triples_del = n_triples_del + 1; //+3
            n_doubles_del = n_doubles_del - 1; //-2
          }
          else if (n_triples_del > 0){
            n_triples_del = n_triples_del - 1; //-3
            n_doubles_del = n_doubles_del + 2; //+4
          }
          else {
            n_singles_del = n_singles_del + 1;
          }
        }
        else if (dif_d == 2){
          n_doubles_del = n_doubles_del + 1;
        }
        else if (dif_d == 3){
          n_triples_del = n_triples_del + 1;
        }
        else if (dif_d == 4){
          if (n_doubles_del > 0){
            n_triples_del = n_triples_del + 2; //+6
            n_doubles_del = n_doubles_del - 1; //-2
          }
          else {
            n_doubles_del = n_doubles_del + 2
          }
        }
      }

      else {
        var dif_d = app_n_delegates - real_n_delegates;
        if (dif_d == 1){
          if (n_triples_del > 0) {
            n_triples_del = n_triples_del - 1; //-3
            n_doubles_del = n_doubles_del + 1; //+2
          }
          else if (n_doubles_del > 0) {
            n_triples_del = n_triples_del + 1; //+3
            n_doubles_del = n_doubles_del - 2; //-4
          }
        }
        else if (dif_d == 2){
          if (n_doubles_del > 0){
            n_doubles_del = n_doubles_del - 1;
          }
          else if (n_triples_del >= 2) {
            n_triples_del = n_triples_del -2; //-6
            n_doubles_del = n_doubles_del + 2; //+4
          }
          else if (n_triples_del > 0) {
            n_triples_del = n_triples_del - 1; //-3
            n_singles_del = n_singles_del + 1; //+1
          }
          else {
            n_singles_del = n_singles_del - 2;
          }
        }
        else if (dif_d == 3){
          if (n_triples_del > 0){
            n_triples_del = n_triples_del - 1;
          }
          else if (n_doubles_del >= 3){
          n_doubles_del = n_doubles_del - 3; //-6
          n_triples_del = n_triples_del + 1; //+3
          }
          else if (n_doubles_del >= 2){
          n_doubles_del = n_doubles_del - 2; //-4
          n_singles_del = n_singles_del + 1; //+1
          }
        }
        else if (dif_d == 4){
          if (n_doubles_del >= 2){
            n_doubles_del = n_doubles_del - 2;
          }
          else if (n_triples_del >= 3){
            n_triples_del = n_triples_del - 3; //-6
            n_doubles_del = n_doubles_del + 1; //+2
          }
        }
      }
    }

    n_triples = n_triples_del + n_triples_teach;
    n_doubles = n_doubles_del + n_doubles_teach;
    n_singles = n_singles_del + n_singles_teach;

    SendEmail('AI_wrongrooms_email', teacher_name, school_name, n_delegates, n_teachers, app_n_triples, app_n_doubles, app_n_singles, app_n_triples_del, app_n_doubles_del, app_n_singles_del, app_n_triples_teach, app_n_doubles_teach, app_n_singles_teach, app_n_participants, n_triples, n_doubles, n_singles, n_triples_del, n_doubles_del, n_singles_del, n_triples_teach, n_doubles_teach, n_singles_teach,'','','','','', emailaddress,'','')

    ui.alert('Check the "Drafts" section of GMAIL and change the default signature for yours. If all is correct, send the email.')

    allinc_sheet.getRange(school_row,21).setValue('Incorrect room number. Suggested '+ n_triples +' triples, '+ n_doubles +' doubles and '+ n_singles +' singles')
    allinc_sheet.getRange(school_row,7,1,6).setBackground('#f4cccc');
  }
}

function std_pd(){
  var school_name_result = '';
  var school_name = '1';

  /**Nom del cole que volem fer el PD*/
  while (school_name != school_name_result) {
    var result = ui.prompt(
      'Standard School Name',
      'Please paste the complete name of the school (just as it is stored in this document):',
      ui.ButtonSet.OK);
    if (result.getSelectedButton() == ui.Button.OK) {
      school_name_result = result.getResponseText();
      for (var i = 2; i <= standard_sheet.getMaxRows();i++) {
        if (school_name_result == standard_sheet.getRange(i,2).getValue()){
          school_name = school_name_result;
          var school_row = i;
          break; //surt de l'if loop
        }
      }
    }
    else {
      return
    }
  }

  //assignar valors de cada cel·la a cada variable 
  var school_number = standard_sheet.getRange(school_row,1).getValue();
  var teacher_name = standard_sheet.getRange(school_row,3).getValue();  
  var emailaddress = standard_sheet.getRange(school_row,4).getValue();
  var n_delegates = standard_sheet.getRange(school_row,5).getValue();
  var n_teachers = standard_sheet.getRange(school_row,6).getValue();
  var social = standard_sheet.getRange(school_row,7).getValue();
  var pd_number = standard_sheet.getRange(school_row,8).getValue();
  var duedate = String(standard_p.getRange('H12').getDisplayValue());

  if (teacher_name.toString() == ''||emailaddress.toString() == ''||n_delegates.toString() == ''||n_teachers.toString() == ''||social.toString() == ''|| (teacher_name.toString().split(' ') [0] != 'Mr' && teacher_name.toString().split(' ') [0] != 'Ms' && teacher_name.toString().split(' ') [0] != 'Mrs')){
    for (var e = 3; e<8; e++){
      if ((standard_sheet.getRange(school_row,e).getValue().toString()=='')){
        missing = standard_sheet.getRange(1,e).getValue();
        ui.alert('You are missing "'+ missing +'" information. Please fill it in and try again.');
      }
    }
    if (teacher_name.toString().split(' ') [0] != 'Mr' && teacher_name.toString().split(' ') [0] != 'Ms' && teacher_name.toString().split(' ') [0] != 'Mrs') {
        ui.alert("You are missing information on the teacher's title (Mr, Ms or Mrs). Please fill it in and try again.");
    }
    return
  }

  //determinar Version number
  if (pd_number == ''){
    var version_n = '01'
  }
  else{
    var pd_number_digits = (""+pd_number).split(""); //passar de string a array (S,0,X,2,3,Y,Y)
    console.log (++(pd_number_digits[2])) //afegir 1 al tercer element de l'array (X, que és el número de versió)
    var version_n = pd_number_digits [1] + pd_number_digits [2]

    //canviar antic pdf document a la carpeta OOD
    var old_pd_number = pd_number;
    var old_pdfName = school_name.replaceAll(' ','_') + '_' +old_pd_number;
    const old_file = pd_folder.getFilesByName(old_pdfName+'.pdf').next();
    old_file.moveTo(ood_folder);
  }

  /**Formatejar el Payment document*/
  //Omplir Payment document amb les dades de standard sheet
  standard_p.getRange('J9').setValue(school_number);
  standard_p.getRange('B12').setValue(version_n);
  pd_number = standard_p.getRange('H9').getValue();

  for (var w = 1; w<1000000; w++){
    if (standard_p.getRange('J9').getValue() == school_number && standard_p.getRange('B12').getValue()==version_n){
      break
    }
  };  

  var pdfName = school_name.replaceAll(' ','_') + '_' +pd_number;

  /**Generar el Payment document*/  
  MakePDF(standard_p,pdfName)

  /**Enviar email amb PDF attached */
  SendEmail('STD_PD_email', teacher_name, school_name, n_delegates, n_teachers,'','','','','','','','','','','','','','','','','','','','', duedate, social, version_n,'', emailaddress, pdfName,'')

  ui.alert('Check the "Drafts" section of GMAIL and change the default signature for yours. If all is correct, send the email.')

  standard_sheet.getRange(school_row,8).setValue(pd_number);
  standard_sheet.getRange(school_row,9).setValue('Yes');

  var date = String(standard_p.getRange('D12').getDisplayValue());
  standard_sheet.getRange(school_row,14).setValue('Payment document '+pd_number+ '. Sent on '+ date);

  return
}

function ai_ddf(){
  var school_name_result = '';
  var school_name = '1';

  /**Nom del cole que volem fer la DDF*/
  while (school_name != school_name_result) {
    var result = ui.prompt(
      'All inclusive School Name',
      'Please paste the complete name of the school (just as it is stored in this document):',
      ui.ButtonSet.OK);
    if (result.getSelectedButton() == ui.Button.OK) {
      school_name_result = result.getResponseText();
      for (var i = 3; i <= allinc_sheet.getMaxRows();i++) {
        if (school_name_result == allinc_sheet.getRange(i,2).getValue()){
          school_name = school_name_result;
          var school_row = i;
          break; //surt de l'if loop
        }
      }
    }
    else {
      return
    }
  }

  //assignar valors de cada cel·la a cada variable 
  var teacher_name = allinc_sheet.getRange(school_row,3).getValue();  
  var emailaddress = allinc_sheet.getRange(school_row,4).getValue();
  var n_delegates = allinc_sheet.getRange(school_row,5).getValue();
  var n_teachers = allinc_sheet.getRange(school_row,6).getValue();
  var n_triples_teach = allinc_sheet.getRange(school_row,7).getValue();
  var n_doubles_teach = allinc_sheet.getRange(school_row,8).getValue();
  var n_singles_teach = allinc_sheet.getRange(school_row,9).getValue();
  var n_triples_del = allinc_sheet.getRange(school_row,10).getValue();
  var n_doubles_del = allinc_sheet.getRange(school_row,11).getValue();
  var n_singles_del = allinc_sheet.getRange(school_row,12).getValue();
  
 /**Copy DDF document template and move it in 'DDF' folder*/
  var ddf_ss = allinc_ddf_template.copy(school_name.replaceAll(' ','_') +'_DDF')
  var ddf_sheet = ddf_ss.getSheets()[0]

  var ddf_ss_id = ddf_ss.getId()
  var ddf_ss_url = ddf_ss.getUrl()

  DriveApp.getFileById(ddf_ss_id).moveTo(DriveApp.getFolderById(ddf_folder_id))

  /** Personalise DDF document with info on the school */

  var room_number = 1
  var room_teacher = n_teachers

  var n_singles = n_singles_del + n_singles_teach
  var n_doubles = n_doubles_del + n_doubles_teach
  var n_triples = n_triples_del + n_triples_teach

  var room_type = [n_singles, n_doubles, n_triples]
  var last_row = ddf_sheet.getLastRow()

  for (var rt = 0; rt <= 2; rt ++){
    for (var n = 1; n <= room_type[rt]; n++){
      var number_rows = rt + 2

      ddf_sheet.insertRowsAfter(last_row,number_rows);

      last_row = last_row + number_rows
      var room_row = last_row - (rt + 1)
      var data_row = last_row - rt

      //Format "number of room" row
      ddf_sheet.getRange(room_row,1,1,7).merge().setBackground('#009fe3').protect();
      ddf_sheet.getRange(room_row,8,1,6).merge().setBackground('#9ccae8').protect();
      ddf_sheet.getRange(room_row,14,1,4).merge().setBackground('#9ccae8').protect();
          
      if (room_teacher>0) {
        if (rt == 0){var s = ''} else {var s = 's'}
        ddf_sheet.getRange(room_row,1,1,7).merge().setBackground('#009fe3').setValue('Room #'+room_number+ ' - Teacher' + s).setFontColor('white').protect();
        ddf_sheet.getRange(data_row,1,rt + 1,17).setBackground('white').setFontColor('#000b3d')
        ddf_sheet.getRange(data_row,8,rt + 1,6).setBackground('#434343').protect()
        room_teacher = room_teacher - 1 - rt
      }
      else {
        ddf_sheet.getRange(room_row,1,1,7).merge().setBackground('#009fe3').setValue('Room #'+room_number).setFontColor('white').protect();
        ddf_sheet.getRange(data_row,1,rt + 1,17).setBackground('white').setFontColor('#000b3d')

        for (var i=8; i<=13; i++){
          if ( i % 2 != 0){
            ddf_sheet.getRange(data_row,i,rt + 1)
            .setDataValidation(SpreadsheetApp.newDataValidation()
              .setAllowInvalid(false)
              .requireValueInList(['Security Council', 'ECOSOC', 'UNESCO', 'UNODC', 'UNHRC', 'WHO', 'UNWOMEN'], true)
              .build());
          }
          else {
            ddf_sheet.getRange(data_row,i,rt + 1)
            .setDataValidation(SpreadsheetApp.newDataValidation()
              .setAllowInvalid(false)
              .setHelpText('There is a typo in your country request. Please write it as specified in the List of Countries document.')
              .requireValueInList(['Afghanistan', 'Albania', 'Algeria', 'Angola', 'Argentina', 'Armenia', 'Australia', 'Austria', 'Azerbaijan', 'Bahrain', 'Bangladesh', 'Belarus', 'Belgium', 'Benin', 'Bhutan', 'Bolivia', 'Bosnia & Herzegovina', 'Botswana', 'Brazil', 'Bulgaria', 'Cambodia', 'Cameroon', 'Canada', 'Chad', 'Chile', 'China', 'Colombia', 'Congo', 'Côte d’Ivoire', 'Croatia', 'Cuba', 'Czech Republic', 'Czechia', 'Denmark', 'Djibouti', 'Dominican Republic', 'DPRK', 'DR Congo', 'Ecuador', 'Egypt', 'El Salvador', 'Eritrea', 'Eswatini', 'Ethiopia', 'Finland', 'France', 'Gabon', 'Gambia', 'Germany', 'Ghana', 'Greece', 'Grenada', 'Guatemala', 'Guinea', 'Guinea-Bisseau', 'Guyana', 'Haiti', 'Honduras', 'Hungary', 'Iceland', 'India', 'Indonesia', 'Irak', 'Iran', 'Ireland', 'Israel', 'Italy', 'Japan', 'Jordan', 'Kazakhstan', 'Kenya', 'Kiribati', 'Kuwait', 'Kyrgyzstan', 'Latvia', 'Libya', 'Lithuania', 'Luxembourg', 'Macedonia', 'Madagascar', 'Malawi', 'Malaysia', 'Maldives', 'Malta', 'Mauritius', 'Mexico', 'Micronesia', 'Moldova', 'Mongolia', 'Montenegro', 'Morocco', 'Mozambique', 'Myanmar', 'Nambia', 'Namibia', 'Nepal', 'Netherlands', 'New Zealand', 'Nicaragua', 'Niger', 'Nigeria', 'Norway', 'Oman', 'Pakistan', 'Panama', 'Papua New Guinea', 'Paraguay', 'Peru', 'Philippines', 'Poland', 'Puerto Rico', 'Qatar', 'Republic of Korea', 'Romania', 'Russian Federation', 'Rwanda', 'Saint Lucia', 'Saudi Arabia', 'Senegal', 'Serbia', 'Slovakia', 'Slovenia', 'Somalia', 'South Africa', 'South Sudan', 'Spain', 'Sweden', 'Switzerland', 'Syria', 'Tajikistan', 'Tanzania', 'Thailand', 'Timor-Leste', 'Togo', 'Trinidad & Tobago', 'Tunisia', 'Turkey', 'Turkmenistan', 'UAE', 'Uganda', 'UK', 'Ukraine', 'Uruguay', 'USA', 'Uzbekistan', 'Venezuela', 'Viet Nam', 'Yemen', 'Zimbabwe'], false)
              .build());
          }
        }
        
        ddf_sheet.getRange(data_row,15,rt + 1).setBackground('#434343').protect()
        
        ddf_sheet.getRange(data_row,7,rt + 1).setBorder(null, null, null, true, null, null, '#000b3d', SpreadsheetApp.BorderStyle.SOLID_MEDIUM)
        ddf_sheet.getRange(data_row,9,rt + 1).setBorder(null, null, null, true, null, null, '#000b3d', SpreadsheetApp.BorderStyle.SOLID)
        ddf_sheet.getRange(data_row,11,rt + 1).setBorder(null, null, null, true, null, null, '#000b3d', SpreadsheetApp.BorderStyle.SOLID)
        ddf_sheet.getRange(data_row,13,rt + 1).setBorder(null, null, null, true, null, null, '#000b3d', SpreadsheetApp.BorderStyle.SOLID_MEDIUM)

      }
      room_number = room_number + 1
    }
  }
  ddf_sheet.setFrozenRows(3)

  ddf_ss.addEditor(emailaddress) //add teacher as editor

  var protections = ddf_sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE)
  
  for (var i = 0; i < protections.length; i++) {
    var protection = protections[i]
    protection.removeEditors([emailaddress]);
  }

  SendEmail('AI_DDF_email', teacher_name, school_name,'','','','','','','','','','','','', n_triples, n_doubles, n_singles, n_triples_del, n_doubles_del, n_singles_del, n_triples_teach, n_doubles_teach, n_singles_teach,'','','','','', emailaddress,'', ddf_ss_url)

  ui.alert('Check the "Drafts" section of GMAIL and change the default signature for yours. If all is correct, send the email.')

  allinc_sheet.getRange(school_row,17).setValue('Yes');
  allinc_sheet.getRange(school_row,21).setValue('DDF sent');
}

function std_ddf(){

  var school_name_result = '';
  var school_name = '1';

  /**Nom del cole que volem fer la DDF */
  while (school_name != school_name_result) {
    var result = ui.prompt(
      'Standard School Name',
      'Please paste the complete name of the school (just as it is stored in this document):',
      ui.ButtonSet.OK);
    if (result.getSelectedButton() == ui.Button.OK) {
      school_name_result = result.getResponseText();
      for (var i = 2; i <= standard_sheet.getMaxRows();i++) {
        if (school_name_result == standard_sheet.getRange(i,2).getValue()){
          school_name = school_name_result;
          var school_row = i;
          break; //surt de l'if loop
        }
      }
    }
    else {
      return
    }
  }
  
  //assignar valors de cada cel·la a cada variable 
  var teacher_name = standard_sheet.getRange(school_row,3).getValue();  
  var emailaddress = standard_sheet.getRange(school_row,4).getValue();
  var n_delegates = standard_sheet.getRange(school_row,5).getValue();
  var n_teachers = standard_sheet.getRange(school_row,6).getValue();
  
 /**Copy DDF document template and move it in 'DDF' folder*/
  var ddf_ss = standard_ddf_template.copy(school_name.replaceAll(' ','_') +'_DDF')
  var ddf_sheet = ddf_ss.getSheets()[0]

  var ddf_ss_id = ddf_ss.getId()
  var ddf_ss_url = ddf_ss.getUrl()

  DriveApp.getFileById(ddf_ss_id).moveTo(DriveApp.getFolderById(ddf_folder_id))

  /** Personalise DDF document with info on the school */

  var last_col = ddf_sheet.getLastColumn()
  
  //insert and edit delegate rows
  ddf_sheet.insertRowsAfter(7,n_delegates)
  ddf_sheet.getRange(8,1,n_delegates,last_col)
    .setBorder(true, true, true, true, null, null, '#000b3d', SpreadsheetApp.BorderStyle.SOLID_MEDIUM)
    .setBorder(null, null, null, null, true, true, '#000b3d', SpreadsheetApp.BorderStyle.SOLID)
    .applyRowBanding(SpreadsheetApp.BandingTheme.BLUE,false,false)
  ddf_sheet.getRange(8,5,n_delegates).setBorder(null, null, null, true, null, null, '#000b3d', SpreadsheetApp.BorderStyle.SOLID_MEDIUM)
  ddf_sheet.getRange(8,7,n_delegates).setBorder(null, null, null, true, null, null, '#000b3d', SpreadsheetApp.BorderStyle.SOLID_MEDIUM)

  for (var i=8; i<=13; i++){
    if ( i % 2 != 0){
      ddf_sheet.getRange(8,i,n_delegates)
      .setBorder(null, false, null, null, null, null)
      .setDataValidation(SpreadsheetApp.newDataValidation()
        .setAllowInvalid(false)
        .requireValueInList(['Security Council', 'ECOSOC', 'UNESCO', 'UNODC', 'UNHRC', 'WHO', 'UNWOMEN'], true)
        .build());
    }
    else {
      ddf_sheet.getRange(8,i,n_delegates)
      .setDataValidation(SpreadsheetApp.newDataValidation()
        .setAllowInvalid(false)
        .setHelpText('There is a typo in your country request. Please write it as specified in the List of Countries document.')
        .requireValueInList(['Afghanistan', 'Albania', 'Algeria', 'Angola', 'Argentina', 'Armenia', 'Australia', 'Austria', 'Azerbaijan', 'Bahrain', 'Bangladesh', 'Belarus', 'Belgium', 'Benin', 'Bhutan', 'Bolivia', 'Bosnia & Herzegovina', 'Botswana', 'Brazil', 'Bulgaria', 'Cambodia', 'Cameroon', 'Canada', 'Chad', 'Chile', 'China', 'Colombia', 'Congo', 'Côte d’Ivoire', 'Croatia', 'Cuba', 'Czech Republic', 'Czechia', 'Denmark', 'Djibouti', 'Dominican Republic', 'DPRK', 'DR Congo', 'Ecuador', 'Egypt', 'El Salvador', 'Eritrea', 'Eswatini', 'Ethiopia', 'Finland', 'France', 'Gabon', 'Gambia', 'Germany', 'Ghana', 'Greece', 'Grenada', 'Guatemala', 'Guinea', 'Guinea-Bisseau', 'Guyana', 'Haiti', 'Honduras', 'Hungary', 'Iceland', 'India', 'Indonesia', 'Irak', 'Iran', 'Ireland', 'Israel', 'Italy', 'Japan', 'Jordan', 'Kazakhstan', 'Kenya', 'Kiribati', 'Kuwait', 'Kyrgyzstan', 'Latvia', 'Libya', 'Lithuania', 'Luxembourg', 'Macedonia', 'Madagascar', 'Malawi', 'Malaysia', 'Maldives', 'Malta', 'Mauritius', 'Mexico', 'Micronesia', 'Moldova', 'Mongolia', 'Montenegro', 'Morocco', 'Mozambique', 'Myanmar', 'Nambia', 'Namibia', 'Nepal', 'Netherlands', 'New Zealand', 'Nicaragua', 'Niger', 'Nigeria', 'Norway', 'Oman', 'Pakistan', 'Panama', 'Papua New Guinea', 'Paraguay', 'Peru', 'Philippines', 'Poland', 'Puerto Rico', 'Qatar', 'Republic of Korea', 'Romania', 'Russian Federation', 'Rwanda', 'Saint Lucia', 'Saudi Arabia', 'Senegal', 'Serbia', 'Slovakia', 'Slovenia', 'Somalia', 'South Africa', 'South Sudan', 'Spain', 'Sweden', 'Switzerland', 'Syria', 'Tajikistan', 'Tanzania', 'Thailand', 'Timor-Leste', 'Togo', 'Trinidad & Tobago', 'Tunisia', 'Turkey', 'Turkmenistan', 'UAE', 'Uganda', 'UK', 'Ukraine', 'Uruguay', 'USA', 'Uzbekistan', 'Venezuela', 'Viet Nam', 'Yemen', 'Zimbabwe'], false)
        .build());
    }
  }

  //insert and edit teacher rows
  if (n_teachers == 0) {
    ddf_sheet.deleteRows(1,3)
  }
  else {
    if (n_teachers == 1) {ddf_sheet.getRange(1,1).setValue('TEACHER')}
    
    ddf_sheet.insertRowsAfter(2,n_teachers)
    ddf_sheet.getRange(3,1,n_teachers,7)
    .setBackground('white')
    .setBorder(true, true, true, true, null, null, '#000b3d', SpreadsheetApp.BorderStyle.SOLID_MEDIUM)
    .setBorder(null, null, null, null, true, true, '#000b3d', SpreadsheetApp.BorderStyle.SOLID)

    ddf_sheet.getRange(1,8,n_teachers+2,6).merge()
      .setBackground('#f3f3f3')
      .setBorder(true, null, true, true, true, true, '#f3f3f3', SpreadsheetApp.BorderStyle.SOLID)
      .protect();

    ddf_sheet.getRange(3,5,n_teachers).setBorder(null, null, null, true, null, null, '#000b3d', SpreadsheetApp.BorderStyle.SOLID_MEDIUM)

  }

  ddf_ss.addEditor(emailaddress) //add teacher as editor

  var protections = ddf_sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE)
  
  for (var i = 0; i < protections.length; i++) {
    var protection = protections[i]
    protection.removeEditors([emailaddress]);
  }

  SendEmail('STD_DDF_email', teacher_name, school_name,'','','','','','','','','','','','','','','','','','','','','','','','','','', emailaddress,'', ddf_ss_url) 

  ui.alert('Check the "Drafts" section of GMAIL and change the default signature for yours. If all is correct, send the email.')

  standard_sheet.getRange(school_row,10).setValue('Yes');
  standard_sheet.getRange(school_row,14).setValue('DDF sent');
}

function SendEmail(doc, teacher_name, school_name, n_delegates, n_teachers, app_n_triples, app_n_doubles, app_n_singles, app_n_triples_del, app_n_doubles_del, app_n_singles_del, app_n_triples_teach, app_n_doubles_teach, app_n_singles_teach, app_n_participants, n_triples, n_doubles, n_singles, n_triples_del, n_doubles_del, n_singles_del, n_triples_teach, n_doubles_teach, n_singles_teach, extranights, duedate, social, version_n, school_row, emailaddress, pdfName, ddf_ss_url) {
  var body = HtmlService.createTemplateFromFile(doc);

  body.teacher_name = teacher_name.toString().split(' ')[0] + ' ' + teacher_name.toString().split(' ')[2];
  body.school_name = school_name;
  body.n_delegates = n_delegates;
  body.n_participants = n_delegates + n_teachers;

  if (n_teachers == 0) {body.n_teachers = 'no'}
  else{body.n_teachers = n_teachers};

  if (n_teachers == 1){body.teacher = 'teacher'}
  else {body.teacher = 'teachers'};

  var app_total_rooms = ''
  var app_del_rooms = ''
  var app_teach_rooms = ''

  if (app_n_triples != 0 && app_n_doubles != 0 && app_n_singles != 0){
    //total rooms
    if (app_n_triples >= 2){
    app_total_rooms = app_n_triples + ' triple rooms, '
    }
    else if (app_n_triples == 1) {
      app_total_rooms = app_n_triples + ' triple room, '
    }

    if (app_n_doubles >= 2){
      app_total_rooms = app_total_rooms + app_n_doubles + ' double rooms and '
    }
    else if (app_n_doubles == 1) {
      app_total_rooms = app_total_rooms +  app_n_doubles + ' double room and '
    }

    if (app_n_singles >= 2){
      app_total_rooms = app_total_rooms + app_n_singles + ' single rooms'
    }
    else if (app_n_singles == 1) {
      app_total_rooms = app_total_rooms + app_n_singles + ' single room'
    }

    //delegate rooms
    if (app_n_triples_del != 0 && app_n_doubles_del != 0 && app_n_singles_del != 0){
      if (app_n_triples_del >= 2){
        app_del_rooms = app_n_triples_del + ' triple rooms, '
      }
      else if (app_n_triples_del == 1) {
        app_del_rooms = app_n_triples_del + ' triple room, '
      }

      if (app_n_doubles_del >= 2){
        app_del_rooms = app_del_rooms + app_n_doubles_del + ' double rooms and '
      }
      else if (app_n_doubles_del == 1) {
        app_del_rooms = app_del_rooms + app_n_doubles_del + ' double room and '
      }

      if (app_n_singles_del >= 2){
        app_del_rooms = app_del_rooms + app_n_singles_del + ' single rooms'
      }
      else if (app_n_singles_del == 1) {
        app_del_rooms = app_del_rooms + app_n_singles_del + ' single room'
      }

    }
    else if (app_n_triples_del != 0 && app_n_doubles_del != 0 && app_n_singles_del == 0){
      if (app_n_triples_del >= 2){
      app_del_rooms = app_n_triples_del + ' triple rooms and '
      }
      else if (app_n_triples_del == 1) {
        app_del_rooms = app_n_triples_del + ' triple room and '
      }

      if (app_n_doubles_del >= 2){
        app_del_rooms = app_del_rooms + app_n_doubles_del + ' double rooms'
      }
      else if (app_n_doubles_del == 1) {
        app_del_rooms = app_del_rooms + app_n_doubles_del + ' double room'
      }
    }
    else if (app_n_triples_del != 0 && app_n_doubles_del == 0 && app_n_singles_del != 0){
      if (app_n_triples_del >= 2){
      app_del_rooms = app_n_triples_del + ' triple rooms and '
      }
      else if (app_n_triples_del == 1) {
        app_del_rooms = app_n_triples_del + ' triple room and '
      }

      if (app_n_singles_del >= 2){
        app_del_rooms = app_del_rooms + app_n_singles_del + ' single rooms'
      }
      else if (app_n_singles_del == 1) {
        app_del_rooms = app_del_rooms + app_n_singles_del + ' single room'
      }
    }
    else if (app_n_triples_del == 0 && app_n_doubles_del != 0 && app_n_singles_del != 0){
      if (app_n_doubles_del >= 2){
        app_del_rooms = app_n_doubles_del + ' double rooms and '
      }
      else if (app_n_doubles_del == 1) {
        app_del_rooms = app_n_doubles_del + ' double room and '
      }

      if (app_n_singles_del >= 2){
        app_del_rooms = app_del_rooms + app_n_singles_del + ' single rooms'
      }
      else if (app_n_singles_del == 1) {
        app_del_rooms = app_del_rooms + app_n_singles_del + ' single room'
      }
    }
    else if (app_n_triples_del != 0){
      if (app_n_triples_del >= 2){
      app_del_rooms = app_n_triples_del + ' triple rooms'
      }
      else if (app_n_triples_del == 1) {
        app_del_rooms = app_n_triples_del + ' triple room'
      }
    }
    else if (app_n_doubles_del != 0){
      if (app_n_doubles_del >= 2){
      app_del_rooms = app_n_doubles_del + ' double rooms'
      }
      else if (app_n_doubles_del == 1) {
        app_del_rooms = app_n_doubles_del + ' double room'
      }
    }
    else if (app_n_singles_del != 0){
      if (app_n_singles_del >= 2){
      app_del_rooms = app_n_singles_del + ' single rooms'
      }
      else if (app_n_singles_del == 1) {
        app_del_rooms = app_n_singles_del + ' single room'
      }
    }
    //teacher rooms
    if (app_n_triples_teach != 0 && app_n_doubles_teach != 0 && app_n_singles_teach != 0){
      if (app_n_triples_teach >= 2){
        app_teach_rooms = app_n_triples_teach + ' triple rooms, '
      }
      else if (app_n_triples_teach == 1) {
        app_teach_rooms = app_n_triples_teach + ' triple room, '
      }

      if (app_n_doubles_teach >= 2){
        app_teach_rooms = app_teach_rooms + app_n_doubles_teach + ' double rooms and '
      }
      else if (app_n_doubles_teach == 1) {
        app_teach_rooms = app_teach_rooms + app_n_doubles_teach + ' double room and '
      }

      if (app_n_singles_teach >= 2){
        app_teach_rooms = app_teach_rooms + app_n_singles_teach + ' single rooms'
      }
      else if (app_n_singles_teach == 1) {
        app_teach_rooms = app_teach_rooms + app_n_singles_teach + ' single room'
      }

    }
    else if (app_n_triples_teach != 0 && app_n_doubles_teach != 0 && app_n_singles_teach == 0){
      if (app_n_triples_teach >= 2){
      app_teach_rooms = app_n_triples_teach + ' triple rooms and '
      }
      else if (app_n_triples_teach == 1) {
        app_teach_rooms = app_n_triples_teach + ' triple room and '
      }

      if (app_n_doubles_teach >= 2){
        app_teach_rooms = app_teach_rooms + app_n_doubles_teach + ' double rooms'
      }
      else if (app_n_doubles_teach == 1) {
        app_teach_rooms = app_teach_rooms + app_n_doubles_teach + ' double room'
      }
    }
    else if (app_n_triples_teach != 0 && app_n_doubles_teach == 0 && app_n_singles_teach != 0){
      if (app_n_triples_teach >= 2){
      app_teach_rooms = app_n_triples_teach + ' triple rooms and '
      }
      else if (app_n_triples_teach == 1) {
        app_teach_rooms = app_n_triples_teach + ' triple room and '
      }

      if (app_n_singles_teach >= 2){
        app_teach_rooms = app_teach_rooms + app_n_singles_teach + ' single rooms'
      }
      else if (app_n_singles_teach == 1) {
        app_teach_rooms = app_teach_rooms + app_n_singles_teach + ' single room'
      }
    }
    else if (app_n_triples_teach == 0 && app_n_doubles_teach != 0 && app_n_singles_teach != 0){
      if (app_n_doubles_teach >= 2){
        app_teach_rooms = app_n_doubles_teach + ' double rooms and '
      }
      else if (app_n_doubles_teach == 1) {
        app_teach_rooms = app_n_doubles_teach + ' double room and '
      }

      if (app_n_singles_teach >= 2){
        app_teach_rooms = app_teach_rooms + app_n_singles_teach + ' single rooms'
      }
      else if (app_n_singles_teach == 1) {
        app_teach_rooms = app_teach_rooms + app_n_singles_teach + ' single room'
      }
    }
    else if (app_n_triples_teach != 0){
      if (app_n_triples_teach >= 2){
      app_teach_rooms = app_n_triples_teach + ' triple rooms'
      }
      else if (app_n_triples_teach == 1) {
        app_teach_rooms = app_n_triples_teach + ' triple room'
      }
    }
    else if (app_n_doubles_teach != 0){
      if (app_n_doubles_teach >= 2){
      app_teach_rooms = app_n_doubles_teach + ' double rooms'
      }
      else if (app_n_doubles_teach == 1) {
        app_teach_rooms = app_n_doubles_teach + ' double room'
      }
    }
    else if (app_n_singles_teach != 0){
      if (app_n_singles_teach >= 2){
      app_teach_rooms = app_n_singles_teach + ' single rooms'
      }
      else if (app_n_singles_teach == 1) {
        app_teach_rooms = app_n_singles_teach + ' single room'
      }
    }
  }

  else if (app_n_triples != 0 && app_n_doubles != 0 && app_n_singles == 0){
    if (app_n_triples >= 2){
    app_total_rooms = app_n_triples + ' triple rooms and '
    }
    else if (app_n_triples == 1) {
      app_total_rooms = app_n_triples + ' triple room and '
    }

    if (app_n_doubles >= 2){
      app_total_rooms = app_total_rooms + app_n_doubles + ' double rooms'
    }
    else if (app_n_doubles == 1) {
      app_total_rooms = app_total_rooms + app_n_doubles + ' double room'
    }

    //delegate rooms
    if (app_n_triples_del != 0 && app_n_doubles_del != 0 && app_n_singles_del == 0){
      if (app_n_triples_del >= 2){
      app_del_rooms = app_n_triples_del + ' triple rooms and '
      }
      else if (app_n_triples_del == 1) {
        app_del_rooms = app_n_triples_del + ' triple room and '
      }

      if (app_n_doubles_del >= 2){
        app_del_rooms = app_del_rooms + app_n_doubles_del + ' double rooms'
      }
      else if (app_n_doubles_del == 1) {
        app_del_rooms = app_del_rooms + app_n_doubles_del + ' double room'
      }
    }
    else if (app_n_triples_del != 0 && app_n_doubles_del == 0 && app_n_singles_del != 0){
      if (app_n_triples_del >= 2){
      app_del_rooms = app_n_triples_del + ' triple rooms and '
      }
      else if (app_n_triples_del == 1) {
        app_del_rooms = app_n_triples_del + ' triple room and '
      }

      if (app_n_singles_del >= 2){
        app_del_rooms = app_del_rooms + app_n_singles_del + ' single rooms'
      }
      else if (app_n_singles_del == 1) {
        app_del_rooms = app_del_rooms + app_n_singles_del + ' single room'
      }
    }
    else if (app_n_triples_del == 0 && app_n_doubles_del != 0 && app_n_singles_del != 0){
      if (app_n_doubles_del >= 2){
        app_del_rooms = app_n_doubles_del + ' double rooms and '
      }
      else if (app_n_doubles_del == 1) {
        app_del_rooms = app_n_doubles_del + ' double room and '
      }

      if (app_n_singles_del >= 2){
        app_del_rooms = app_del_rooms + app_n_singles_del + ' single rooms'
      }
      else if (app_n_singles_del == 1) {
        app_del_rooms = app_del_rooms + app_n_singles_del + ' single room'
      }
    }
    else if (app_n_triples_del != 0){
      if (app_n_triples_del >= 2){
      app_del_rooms = app_n_triples_del + ' triple rooms'
      }
      else if (app_n_triples_del == 1) {
        app_del_rooms = app_n_triples_del + ' triple room'
      }
    }
    else if (app_n_doubles_del != 0){
      if (app_n_doubles_del >= 2){
      app_del_rooms = app_n_doubles_del + ' double rooms'
      }
      else if (app_n_doubles_del == 1) {
        app_del_rooms = app_n_doubles_del + ' double room'
      }
    }
    else if (app_n_singles_del != 0){
      if (app_n_singles_del >= 2){
      app_del_rooms = app_n_singles_del + ' single rooms'
      }
      else if (app_n_singles_del == 1) {
        app_del_rooms = app_n_singles_del + ' single room'
      }
    }
    //teacher rooms
    if (app_n_triples_teach != 0 && app_n_doubles_teach != 0 && app_n_singles_teach == 0){
      if (app_n_triples_teach >= 2){
        app_teach_rooms = app_n_triples_teach + ' triple rooms and '
      }
      else if (app_n_triples_teach == 1) {
        app_teach_rooms = app_n_triples_teach + ' triple room and '
      }

      if (app_n_doubles_teach >= 2){
        app_teach_rooms = app_teach_rooms + app_n_doubles_teach + ' double rooms'
      }
      else if (app_n_doubles_teach == 1) {
        app_teach_rooms = app_teach_rooms + app_n_doubles_teach + ' double room'
      }
    }
    else if (app_n_triples_teach != 0 && app_n_doubles_teach == 0 && app_n_singles_teach != 0){
      if (app_n_triples_teach >= 2){
      app_teach_rooms = app_n_triples_teach + ' triple rooms and '
      }
      else if (app_n_triples_teach == 1) {
        app_teach_rooms = app_n_triples_teach + ' triple room and '
      }

      if (app_n_singles_teach >= 2){
        app_teach_rooms = app_teach_rooms + app_n_singles_teach + ' single rooms'
      }
      else if (app_n_singles_teach == 1) {
        app_teach_rooms = app_teach_rooms + app_n_singles_teach + ' single room'
      }
    }
    else if (app_n_triples_teach == 0 && app_n_doubles_teach != 0 && app_n_singles_teach != 0){
      if (app_n_doubles_teach >= 2){
        app_teach_rooms = app_n_doubles_teach + ' double rooms and '
      }
      else if (app_n_doubles_teach == 1) {
        app_teach_rooms = app_n_doubles_teach + ' double room and '
      }

      if (app_n_singles_teach >= 2){
        app_teach_rooms = app_teach_rooms + app_n_singles_teach + ' single rooms'
      }
      else if (app_n_singles_teach == 1) {
        app_teach_rooms = app_teach_rooms + app_n_singles_teach + ' single room'
      }
    }
    else if (app_n_triples_teach != 0){
      if (app_n_triples_teach >= 2){
      app_teach_rooms = app_n_triples_teach + ' triple rooms'
      }
      else if (app_n_triples_teach == 1) {
        app_teach_rooms = app_n_triples_teach + ' triple room'
      }
    }
    else if (app_n_doubles_teach != 0){
      if (app_n_doubles_teach >= 2){
      app_teach_rooms = app_n_doubles_teach + ' double rooms'
      }
      else if (app_n_doubles_teach == 1) {
        app_teach_rooms = app_n_doubles_teach + ' double room'
      }
    }
    else if (app_n_singles_teach != 0){
      if (app_n_singles_teach >= 2){
      app_teach_rooms = app_n_singles_teach + ' single rooms'
      }
      else if (app_n_singles_teach == 1) {
        app_teach_rooms = app_n_singles_teach + ' single room'
      }
    }
  }

  else if (app_n_triples != 0 && app_n_doubles == 0 && app_n_singles != 0){
    if (app_n_triples >= 2){
    app_total_rooms = app_n_triples + ' triple rooms and '
    }
    else if (app_n_triples == 1) {
      app_total_rooms = app_n_triples + ' triple room and '
    }

    if (app_n_singles >= 2){
      app_total_rooms = app_total_rooms + app_n_singles + ' single rooms'
    }
    else if (app_n_singles == 1) {
      app_total_rooms = app_total_rooms + app_n_singles + ' single room'
    }

    //delegate rooms
    if (app_n_triples_del != 0 && app_n_doubles_del == 0 && app_n_singles_del != 0){
      if (app_n_triples_del >= 2){
      app_del_rooms = app_n_triples_del + ' triple rooms and '
      }
      else if (app_n_triples_del == 1) {
        app_del_rooms = app_n_triples_del + ' triple room and '
      }

      if (app_n_singles_del >= 2){
        app_del_rooms = app_del_rooms + app_n_singles_del + ' single rooms'
      }
      else if (app_n_singles_del == 1) {
        app_del_rooms = app_del_rooms + app_n_singles_del + ' single room'
      }
    }
    else if (app_n_triples_del == 0 && app_n_doubles_del != 0 && app_n_singles_del != 0){
      if (app_n_doubles_del >= 2){
        app_del_rooms = app_n_doubles_del + ' double rooms and '
      }
      else if (app_n_doubles_del == 1) {
        app_del_rooms = app_n_doubles_del + ' double room and '
      }

      if (app_n_singles_del >= 2){
        app_del_rooms = app_del_rooms + app_n_singles_del + ' single rooms'
      }
      else if (app_n_singles_del == 1) {
        app_del_rooms = app_del_rooms + app_n_singles_del + ' single room'
      }
    }
    else if (app_n_triples_del != 0){
      if (app_n_triples_del >= 2){
      app_del_rooms = app_n_triples_del + ' triple rooms'
      }
      else if (app_n_triples_del == 1) {
        app_del_rooms = app_n_triples_del + ' triple room'
      }
    }
    else if (app_n_doubles_del != 0){
      if (app_n_doubles_del >= 2){
      app_del_rooms = app_n_doubles_del + ' double rooms'
      }
      else if (app_n_doubles_del == 1) {
        app_del_rooms = app_n_doubles_del + ' double room'
      }
    }
    else if (app_n_singles_del != 0){
      if (app_n_singles_del >= 2){
      app_del_rooms = app_n_singles_del + ' single rooms'
      }
      else if (app_n_singles_del == 1) {
        app_del_rooms = app_n_singles_del + ' single room'
      }
    }
    //teacher rooms
    if (app_n_triples_teach != 0 && app_n_doubles_teach == 0 && app_n_singles_teach != 0){
      if (app_n_triples_teach >= 2){
        app_teach_rooms = app_n_triples_teach + ' triple rooms and '
      }
      else if (app_n_triples_teach == 1) {
        app_teach_rooms = app_n_triples_teach + ' triple room and '
      }

      if (app_n_singles_teach >= 2){
        app_teach_rooms = app_teach_rooms + app_n_singles_teach + ' single rooms'
      }
      else if (app_n_singles_teach == 1) {
        app_teach_rooms = app_teach_rooms + app_n_singles_teach + ' single room'
      }
    }
    else if (app_n_triples_teach == 0 && app_n_doubles_teach != 0 && app_n_singles_teach != 0){
      if (app_n_doubles_teach >= 2){
        app_teach_rooms = app_n_doubles_teach + ' double rooms and '
      }
      else if (app_n_doubles_teach == 1) {
        app_teach_rooms = app_n_doubles_teach + ' double room and '
      }

      if (app_n_singles_teach >= 2){
        app_teach_rooms = app_teach_rooms + app_n_singles_teach + ' single rooms'
      }
      else if (app_n_singles_teach == 1) {
        app_teach_rooms = app_teach_rooms + app_n_singles_teach + ' single room'
      }
    }
    else if (app_n_triples_teach != 0){
      if (app_n_triples_teach >= 2){
      app_teach_rooms = app_n_triples_teach + ' triple rooms'
      }
      else if (app_n_triples_teach == 1) {
        app_teach_rooms = app_n_triples_teach + ' triple room'
      }
    }
    else if (app_n_doubles_teach != 0){
      if (app_n_doubles_teach >= 2){
      app_teach_rooms = app_n_doubles_teach + ' double rooms'
      }
      else if (app_n_doubles_teach == 1) {
        app_teach_rooms = app_n_doubles_teach + ' double room'
      }
    }
    else if (app_n_singles_teach != 0){
      if (app_n_singles_teach >= 2){
      app_teach_rooms = app_n_singles_teach + ' single rooms'
      }
      else if (app_n_singles_teach == 1) {
        app_teach_rooms = app_n_singles_teach + ' single room'
      }
    }

  }

  else if (app_n_triples == 0 && app_n_doubles != 0 && app_n_singles != 0){
    if (app_n_doubles >= 2){
    app_total_rooms = app_n_doubles + ' double rooms and '
    }
    else if (app_n_doubles == 1) {
      app_total_rooms = app_n_doubles + ' double room and '
    }

    if (app_n_singles >= 2){
      app_total_rooms = app_total_rooms + app_n_singles + ' single rooms'
    }
    else if (app_n_singles == 1) {
      app_total_rooms = app_total_rooms + app_n_singles + ' single room'
    }

    //delegate rooms
    if (app_n_triples_del == 0 && app_n_doubles_del != 0 && app_n_singles_del != 0){
      if (app_n_doubles_del >= 2){
        app_del_rooms = app_n_doubles_del + ' double rooms and '
      }
      else if (app_n_doubles_del == 1) {
        app_del_rooms = app_n_doubles_del + ' double room and '
      }

      if (app_n_singles_del >= 2){
        app_del_rooms = app_del_rooms + app_n_singles_del + ' single rooms'
      }
      else if (app_n_singles_del == 1) {
        app_del_rooms = app_del_rooms + app_n_singles_del + ' single room'
      }
    }
    else if (app_n_triples_del != 0){
      if (app_n_triples_del >= 2){
      app_del_rooms = app_n_triples_del + ' triple rooms'
      }
      else if (app_n_triples_del == 1) {
        app_del_rooms = app_n_triples_del + ' triple room'
      }
    }
    else if (app_n_doubles_del != 0){
      if (app_n_doubles_del >= 2){
      app_del_rooms = app_n_doubles_del + ' double rooms'
      }
      else if (app_n_doubles_del == 1) {
        app_del_rooms = app_n_doubles_del + ' double room'
      }
    }
    else if (app_n_singles_del != 0){
      if (app_n_singles_del >= 2){
      app_del_rooms = app_n_singles_del + ' single rooms'
      }
      else if (app_n_singles_del == 1) {
        app_del_rooms = app_n_singles_del + ' single room'
      }
    }
    //teacher rooms
    if (app_n_triples_teach == 0 && app_n_doubles_teach != 0 && app_n_singles_teach != 0){
      if (app_n_doubles_teach >= 2){
        app_teach_rooms = app_n_doubles_teach + ' double rooms and '
      }
      else if (app_n_doubles_teach == 1) {
        app_teach_rooms = app_n_doubles_teach + ' double room and '
      }

      if (app_n_singles_teach >= 2){
        app_teach_rooms = app_teach_rooms + app_n_singles_teach + ' single rooms'
      }
      else if (app_n_singles_teach == 1) {
        app_teach_rooms = app_teach_rooms + app_n_singles_teach + ' single room'
      }
    }
    else if (app_n_triples_teach != 0){
      if (app_n_triples_teach >= 2){
      app_teach_rooms = app_n_triples_teach + ' triple rooms'
      }
      else if (app_n_triples_teach == 1) {
        app_teach_rooms = app_n_triples_teach + ' triple room'
      }
    }
    else if (app_n_doubles_teach != 0){
      if (app_n_doubles_teach >= 2){
      app_teach_rooms = app_n_doubles_teach + ' double rooms'
      }
      else if (app_n_doubles_teach == 1) {
        app_teach_rooms = app_n_doubles_teach + ' double room'
      }
    }
    else if (app_n_singles_teach != 0){
      if (app_n_singles_teach >= 2){
      app_teach_rooms = app_n_singles_teach + ' single rooms'
      }
      else if (app_n_singles_teach == 1) {
        app_teach_rooms = app_n_singles_teach + ' single room'
      }
    }
  }

  else if (app_n_triples != 0){
    if (app_n_triples >= 2){
    app_total_rooms = app_n_triples + ' triple rooms'
    }
    else if (app_n_triples == 1) {
      app_total_rooms = app_n_triples + ' triple room'
    }

    //delegate rooms
    if (app_n_triples_del != 0){
      if (app_n_triples_del >= 2){
      app_app_del_rooms = app_n_triples_del + ' triple rooms'
      }
      else if (app_n_triples_del == 1) {
        app_del_rooms = app_n_triples_del + ' triple room'
      }
    }

    //teacher rooms
    if (app_n_triples_teach != 0){
      if (app_n_triples_teach >= 2){
      app_teach_rooms = app_n_triples_teach + ' triple rooms'
      }
      else if (app_n_triples_teach == 1) {
        app_teach_rooms = app_n_triples_teach + ' triple room'
      }
    }
  }

  else if (app_n_doubles != 0){
    if (app_n_doubles >= 2){
    app_total_rooms = app_n_doubles + ' double rooms'
    }
    else if (app_n_doubles == 1) {
      app_total_rooms = app_n_doubles + ' double room'
    }

    //delegate rooms
    if (app_n_doubles_del != 0){
      if (app_n_doubles_del >= 2){
      app_del_rooms = app_n_doubles_del + ' double rooms'
      }
      else if (app_n_doubles_del == 1) {
        app_del_rooms = app_n_doubles_del + ' double room'
      }
    }

     //teacher rooms
    if (app_n_doubles_teach != 0){
      if (app_n_doubles_teach >= 2){
      app_teach_rooms = app_n_doubles_teach + ' double rooms'
      }
      else if (app_n_doubles_teach == 1) {
        app_teach_rooms = app_n_doubles_teach + ' double room'
      }
    }
  }

  else if (app_n_singles != 0){
    if (app_n_singles >= 2){
    app_total_rooms = app_n_singles + ' single rooms'
    }
    else if (app_n_singles == 1) {
      app_total_rooms = app_n_singles + ' single room'
    }
    //delegate rooms
    if (app_n_singles_del != 0){
      if (app_n_singles_del >= 2){
      app_del_rooms = app_n_singles_del + ' single rooms'
      }
      else if (app_n_singles_del == 1) {
        app_del_rooms = app_n_singles_del + ' single room'
      }
    }
    //teacher rooms
    if (app_n_singles_teach != 0){
      if (app_n_singles_teach >= 2){
      app_teach_rooms = app_n_singles_teach + ' single rooms'
      }
      else if (app_n_singles_teach == 1) {
        app_teach_rooms = app_n_singles_teach + ' single room'
      }
    }
  }

  if (app_teach_rooms == ''){
    app_teach_rooms = 'none'
  }

  body.app_total_rooms = app_total_rooms;
  body.app_delegate_rooms = app_del_rooms;
  body.app_teacher_rooms = app_teach_rooms;
    
  body.app_n_participants = app_n_participants;

  var total_rooms = ''
  var del_rooms = ''
  var teach_rooms = ''

  if (n_triples != 0 && n_doubles != 0 && n_singles != 0){
    //total rooms
    if (n_triples >= 2){
    total_rooms = n_triples + ' triple rooms, '
    }
    else if (n_triples == 1) {
      total_rooms = n_triples + ' triple room, '
    }

    if (n_doubles >= 2){
      total_rooms = total_rooms + n_doubles + ' double rooms and '
    }
    else if (n_doubles == 1) {
      total_rooms = total_rooms +  n_doubles + ' double room and '
    }

    if (n_singles >= 2){
      total_rooms = total_rooms + n_singles + ' single rooms'
    }
    else if (n_singles == 1) {
      total_rooms = total_rooms + n_singles + ' single room'
    }

    //delegate rooms
    if (n_triples_del != 0 && n_doubles_del != 0 && n_singles_del != 0){
      if (n_triples_del >= 2){
        del_rooms = n_triples_del + ' triple rooms, '
      }
      else if (n_triples_del == 1) {
        del_rooms = n_triples_del + ' triple room, '
      }

      if (n_doubles_del >= 2){
        del_rooms = del_rooms + n_doubles_del + ' double rooms and '
      }
      else if (n_doubles_del == 1) {
        del_rooms = del_rooms + n_doubles_del + ' double room and '
      }

      if (n_singles_del >= 2){
        del_rooms = del_rooms + n_singles_del + ' single rooms'
      }
      else if (n_singles_del == 1) {
        del_rooms = del_rooms + n_singles_del + ' single room'
      }

    }
    else if (n_triples_del != 0 && n_doubles_del != 0 && n_singles_del == 0){
      if (n_triples_del >= 2){
      del_rooms = n_triples_del + ' triple rooms and '
      }
      else if (n_triples_del == 1) {
        del_rooms = n_triples_del + ' triple room and '
      }

      if (n_doubles_del >= 2){
        del_rooms = del_rooms + n_doubles_del + ' double rooms'
      }
      else if (n_doubles_del == 1) {
        del_rooms = del_rooms + n_doubles_del + ' double room'
      }
    }
    else if (n_triples_del != 0 && n_doubles_del == 0 && n_singles_del != 0){
      if (n_triples_del >= 2){
      del_rooms = n_triples_del + ' triple rooms and '
      }
      else if (n_triples_del == 1) {
        del_rooms = n_triples_del + ' triple room and '
      }

      if (n_singles_del >= 2){
        del_rooms = del_rooms + n_singles_del + ' single rooms'
      }
      else if (n_singles_del == 1) {
        del_rooms = del_rooms + n_singles_del + ' single room'
      }
    }
    else if (n_triples_del == 0 && n_doubles_del != 0 && n_singles_del != 0){
      if (n_doubles_del >= 2){
        del_rooms = n_doubles_del + ' double rooms and '
      }
      else if (n_doubles_del == 1) {
        del_rooms = n_doubles_del + ' double room and '
      }

      if (n_singles_del >= 2){
        del_rooms = del_rooms + n_singles_del + ' single rooms'
      }
      else if (n_singles_del == 1) {
        del_rooms = del_rooms + n_singles_del + ' single room'
      }
    }
    else if (n_triples_del != 0){
      if (n_triples_del >= 2){
      del_rooms = n_triples_del + ' triple rooms'
      }
      else if (n_triples_del == 1) {
        del_rooms = n_triples_del + ' triple room'
      }
    }
    else if (n_doubles_del != 0){
      if (n_doubles_del >= 2){
      del_rooms = n_doubles_del + ' double rooms'
      }
      else if (n_doubles_del == 1) {
        del_rooms = n_doubles_del + ' double room'
      }
    }
    else if (n_singles_del != 0){
      if (n_singles_del >= 2){
      del_rooms = n_singles_del + ' single rooms'
      }
      else if (n_singles_del == 1) {
        del_rooms = n_singles_del + ' single room'
      }
    }
    //teacher rooms
    if (n_triples_teach != 0 && n_doubles_teach != 0 && n_singles_teach != 0){
      if (n_triples_teach >= 2){
        teach_rooms = n_triples_teach + ' triple rooms, '
      }
      else if (n_triples_teach == 1) {
        teach_rooms = n_triples_teach + ' triple room, '
      }

      if (n_doubles_teach >= 2){
        teach_rooms = teach_rooms + n_doubles_teach + ' double rooms and '
      }
      else if (n_doubles_teach == 1) {
        teach_rooms = teach_rooms + n_doubles_teach + ' double room and '
      }

      if (n_singles_teach >= 2){
        teach_rooms = teach_rooms + n_singles_teach + ' single rooms'
      }
      else if (n_singles_teach == 1) {
        teach_rooms = teach_rooms + n_singles_teach + ' single room'
      }

    }
    else if (n_triples_teach != 0 && n_doubles_teach != 0 && n_singles_teach == 0){
      if (n_triples_teach >= 2){
      teach_rooms = n_triples_teach + ' triple rooms and '
      }
      else if (n_triples_teach == 1) {
        teach_rooms = n_triples_teach + ' triple room and '
      }

      if (n_doubles_teach >= 2){
        teach_rooms = teach_rooms + n_doubles_teach + ' double rooms'
      }
      else if (n_doubles_teach == 1) {
        teach_rooms = teach_rooms + n_doubles_teach + ' double room'
      }
    }
    else if (n_triples_teach != 0 && n_doubles_teach == 0 && n_singles_teach != 0){
      if (n_triples_teach >= 2){
      teach_rooms = n_triples_teach + ' triple rooms and '
      }
      else if (n_triples_teach == 1) {
        teach_rooms = n_triples_teach + ' triple room and '
      }

      if (n_singles_teach >= 2){
        teach_rooms = teach_rooms + n_singles_teach + ' single rooms'
      }
      else if (n_singles_teach == 1) {
        teach_rooms = teach_rooms + n_singles_teach + ' single room'
      }
    }
    else if (n_triples_teach == 0 && n_doubles_teach != 0 && n_singles_teach != 0){
      if (n_doubles_teach >= 2){
        teach_rooms = n_doubles_teach + ' double rooms and '
      }
      else if (n_doubles_teach == 1) {
        teach_rooms = n_doubles_teach + ' double room and '
      }

      if (n_singles_teach >= 2){
        teach_rooms = teach_rooms + n_singles_teach + ' single rooms'
      }
      else if (n_singles_teach == 1) {
        teach_rooms = teach_rooms + n_singles_teach + ' single room'
      }
    }
    else if (n_triples_teach != 0){
      if (n_triples_teach >= 2){
      teach_rooms = n_triples_teach + ' triple rooms'
      }
      else if (n_triples_teach == 1) {
        teach_rooms = n_triples_teach + ' triple room'
      }
    }
    else if (n_doubles_teach != 0){
      if (n_doubles_teach >= 2){
      teach_rooms = n_doubles_teach + ' double rooms'
      }
      else if (n_doubles_teach == 1) {
        teach_rooms = n_doubles_teach + ' double room'
      }
    }
    else if (n_singles_teach != 0){
      if (n_singles_teach >= 2){
      teach_rooms = n_singles_teach + ' single rooms'
      }
      else if (n_singles_teach == 1) {
        teach_rooms = n_singles_teach + ' single room'
      }
    }
  }

  else if (n_triples != 0 && n_doubles != 0 && n_singles == 0){
    if (n_triples >= 2){
    total_rooms = n_triples + ' triple rooms and '
    }
    else if (n_triples == 1) {
      total_rooms = n_triples + ' triple room and '
    }

    if (n_doubles >= 2){
      total_rooms = total_rooms + n_doubles + ' double rooms'
    }
    else if (n_doubles == 1) {
      total_rooms = total_rooms + n_doubles + ' double room'
    }

    //delegate rooms
    if (n_triples_del != 0 && n_doubles_del != 0 && n_singles_del == 0){
      if (n_triples_del >= 2){
      del_rooms = n_triples_del + ' triple rooms and '
      }
      else if (n_triples_del == 1) {
        del_rooms = n_triples_del + ' triple room and '
      }

      if (n_doubles_del >= 2){
        del_rooms = del_rooms + n_doubles_del + ' double rooms'
      }
      else if (n_doubles_del == 1) {
        del_rooms = del_rooms + n_doubles_del + ' double room'
      }
    }
    else if (n_triples_del != 0 && n_doubles_del == 0 && n_singles_del != 0){
      if (n_triples_del >= 2){
      del_rooms = n_triples_del + ' triple rooms and '
      }
      else if (n_triples_del == 1) {
        del_rooms = n_triples_del + ' triple room and '
      }

      if (n_singles_del >= 2){
        del_rooms = del_rooms + n_singles_del + ' single rooms'
      }
      else if (n_singles_del == 1) {
        del_rooms = del_rooms + n_singles_del + ' single room'
      }
    }
    else if (n_triples_del == 0 && n_doubles_del != 0 && n_singles_del != 0){
      if (n_doubles_del >= 2){
        del_rooms = n_doubles_del + ' double rooms and '
      }
      else if (n_doubles_del == 1) {
        del_rooms = n_doubles_del + ' double room and '
      }

      if (n_singles_del >= 2){
        del_rooms = del_rooms + n_singles_del + ' single rooms'
      }
      else if (n_singles_del == 1) {
        del_rooms = del_rooms + n_singles_del + ' single room'
      }
    }
    else if (n_triples_del != 0){
      if (n_triples_del >= 2){
      del_rooms = n_triples_del + ' triple rooms'
      }
      else if (n_triples_del == 1) {
        del_rooms = n_triples_del + ' triple room'
      }
    }
    else if (n_doubles_del != 0){
      if (n_doubles_del >= 2){
      del_rooms = n_doubles_del + ' double rooms'
      }
      else if (n_doubles_del == 1) {
        del_rooms = n_doubles_del + ' double room'
      }
    }
    else if (n_singles_del != 0){
      if (n_singles_del >= 2){
      del_rooms = n_singles_del + ' single rooms'
      }
      else if (n_singles_del == 1) {
        del_rooms = n_singles_del + ' single room'
      }
    }
    //teacher rooms
    if (n_triples_teach != 0 && n_doubles_teach != 0 && n_singles_teach == 0){
      if (n_triples_teach >= 2){
        teach_rooms = n_triples_teach + ' triple rooms and '
      }
      else if (n_triples_teach == 1) {
        teach_rooms = n_triples_teach + ' triple room and '
      }

      if (n_doubles_teach >= 2){
        teach_rooms = teach_rooms + n_doubles_teach + ' double rooms'
      }
      else if (n_doubles_teach == 1) {
        teach_rooms = teach_rooms + n_doubles_teach + ' double room'
      }
    }
    else if (n_triples_teach != 0 && n_doubles_teach == 0 && n_singles_teach != 0){
      if (n_triples_teach >= 2){
      teach_rooms = n_triples_teach + ' triple rooms and '
      }
      else if (n_triples_teach == 1) {
        teach_rooms = n_triples_teach + ' triple room and '
      }

      if (n_singles_teach >= 2){
        teach_rooms = teach_rooms + n_singles_teach + ' single rooms'
      }
      else if (n_singles_teach == 1) {
        teach_rooms = teach_rooms + n_singles_teach + ' single room'
      }
    }
    else if (n_triples_teach == 0 && n_doubles_teach != 0 && n_singles_teach != 0){
      if (n_doubles_teach >= 2){
        teach_rooms = n_doubles_teach + ' double rooms and '
      }
      else if (n_doubles_teach == 1) {
        teach_rooms = n_doubles_teach + ' double room and '
      }

      if (n_singles_teach >= 2){
        teach_rooms = teach_rooms + n_singles_teach + ' single rooms'
      }
      else if (n_singles_teach == 1) {
        teach_rooms = teach_rooms + n_singles_teach + ' single room'
      }
    }
    else if (n_triples_teach != 0){
      if (n_triples_teach >= 2){
      teach_rooms = n_triples_teach + ' triple rooms'
      }
      else if (n_triples_teach == 1) {
        teach_rooms = n_triples_teach + ' triple room'
      }
    }
    else if (n_doubles_teach != 0){
      if (n_doubles_teach >= 2){
      teach_rooms = n_doubles_teach + ' double rooms'
      }
      else if (n_doubles_teach == 1) {
        teach_rooms = n_doubles_teach + ' double room'
      }
    }
    else if (n_singles_teach != 0){
      if (n_singles_teach >= 2){
      teach_rooms = n_singles_teach + ' single rooms'
      }
      else if (n_singles_teach == 1) {
        teach_rooms = n_singles_teach + ' single room'
      }
    }
  }

  else if (n_triples != 0 && n_doubles == 0 && n_singles != 0){
    if (n_triples >= 2){
    total_rooms = n_triples + ' triple rooms and '
    }
    else if (n_triples == 1) {
      total_rooms = n_triples + ' triple room and '
    }

    if (n_singles >= 2){
      total_rooms = total_rooms + n_singles + ' single rooms'
    }
    else if (n_singles == 1) {
      total_rooms = total_rooms + n_singles + ' single room'
    }

    //delegate rooms
    if (n_triples_del != 0 && n_doubles_del == 0 && n_singles_del != 0){
      if (n_triples_del >= 2){
      del_rooms = n_triples_del + ' triple rooms and '
      }
      else if (n_triples_del == 1) {
        del_rooms = n_triples_del + ' triple room and '
      }

      if (n_singles_del >= 2){
        del_rooms = del_rooms + n_singles_del + ' single rooms'
      }
      else if (n_singles_del == 1) {
        del_rooms = del_rooms + n_singles_del + ' single room'
      }
    }
    else if (n_triples_del == 0 && n_doubles_del != 0 && n_singles_del != 0){
      if (n_doubles_del >= 2){
        del_rooms = n_doubles_del + ' double rooms and '
      }
      else if (n_doubles_del == 1) {
        del_rooms = n_doubles_del + ' double room and '
      }

      if (n_singles_del >= 2){
        del_rooms = del_rooms + n_singles_del + ' single rooms'
      }
      else if (n_singles_del == 1) {
        del_rooms = del_rooms + n_singles_del + ' single room'
      }
    }
    else if (n_triples_del != 0){
      if (n_triples_del >= 2){
      del_rooms = n_triples_del + ' triple rooms'
      }
      else if (n_triples_del == 1) {
        del_rooms = n_triples_del + ' triple room'
      }
    }
    else if (n_doubles_del != 0){
      if (n_doubles_del >= 2){
      del_rooms = n_doubles_del + ' double rooms'
      }
      else if (n_doubles_del == 1) {
        del_rooms = n_doubles_del + ' double room'
      }
    }
    else if (n_singles_del != 0){
      if (n_singles_del >= 2){
      del_rooms = n_singles_del + ' single rooms'
      }
      else if (n_singles_del == 1) {
        del_rooms = n_singles_del + ' single room'
      }
    }
    //teacher rooms
    if (n_triples_teach != 0 && n_doubles_teach == 0 && n_singles_teach != 0){
      if (n_triples_teach >= 2){
        teach_rooms = n_triples_teach + ' triple rooms and '
      }
      else if (n_triples_teach == 1) {
        teach_rooms = n_triples_teach + ' triple room and '
      }

      if (n_singles_teach >= 2){
        teach_rooms = teach_rooms + n_singles_teach + ' single rooms'
      }
      else if (n_singles_teach == 1) {
        teach_rooms = teach_rooms + n_singles_teach + ' single room'
      }
    }
    else if (n_triples_teach == 0 && n_doubles_teach != 0 && n_singles_teach != 0){
      if (n_doubles_teach >= 2){
        teach_rooms = n_doubles_teach + ' double rooms and '
      }
      else if (n_doubles_teach == 1) {
        teach_rooms = n_doubles_teach + ' double room and '
      }

      if (n_singles_teach >= 2){
        teach_rooms = teach_rooms + n_singles_teach + ' single rooms'
      }
      else if (n_singles_teach == 1) {
        teach_rooms = teach_rooms + n_singles_teach + ' single room'
      }
    }
    else if (n_triples_teach != 0){
      if (n_triples_teach >= 2){
      teach_rooms = n_triples_teach + ' triple rooms'
      }
      else if (n_triples_teach == 1) {
        teach_rooms = n_triples_teach + ' triple room'
      }
    }
    else if (n_doubles_teach != 0){
      if (n_doubles_teach >= 2){
      teach_rooms = n_doubles_teach + ' double rooms'
      }
      else if (n_doubles_teach == 1) {
        teach_rooms = n_doubles_teach + ' double room'
      }
    }
    else if (n_singles_teach != 0){
      if (n_singles_teach >= 2){
      teach_rooms = n_singles_teach + ' single rooms'
      }
      else if (n_singles_teach == 1) {
        teach_rooms = n_singles_teach + ' single room'
      }
    }

  }

  else if (n_triples == 0 && n_doubles != 0 && n_singles != 0){
    if (n_doubles >= 2){
    total_rooms = n_doubles + ' double rooms and '
    }
    else if (n_doubles == 1) {
      total_rooms = n_doubles + ' double room and '
    }

    if (n_singles >= 2){
      total_rooms = total_rooms + n_singles + ' single rooms'
    }
    else if (n_singles == 1) {
      total_rooms = total_rooms + n_singles + ' single room'
    }

    //delegate rooms
    if (n_triples_del == 0 && n_doubles_del != 0 && n_singles_del != 0){
      if (n_doubles_del >= 2){
        del_rooms = n_doubles_del + ' double rooms and '
      }
      else if (n_doubles_del == 1) {
        del_rooms = n_doubles_del + ' double room and '
      }

      if (n_singles_del >= 2){
        del_rooms = del_rooms + n_singles_del + ' single rooms'
      }
      else if (n_singles_del == 1) {
        del_rooms = del_rooms + n_singles_del + ' single room'
      }
    }
    else if (n_triples_del != 0){
      if (n_triples_del >= 2){
      del_rooms = n_triples_del + ' triple rooms'
      }
      else if (n_triples_del == 1) {
        del_rooms = n_triples_del + ' triple room'
      }
    }
    else if (n_doubles_del != 0){
      if (n_doubles_del >= 2){
      del_rooms = n_doubles_del + ' double rooms'
      }
      else if (n_doubles_del == 1) {
        del_rooms = n_doubles_del + ' double room'
      }
    }
    else if (n_singles_del != 0){
      if (n_singles_del >= 2){
      del_rooms = n_singles_del + ' single rooms'
      }
      else if (n_singles_del == 1) {
        del_rooms = n_singles_del + ' single room'
      }
    }
    //teacher rooms
    if (n_triples_teach == 0 && n_doubles_teach != 0 && n_singles_teach != 0){
      if (n_doubles_teach >= 2){
        teach_rooms = n_doubles_teach + ' double rooms and '
      }
      else if (n_doubles_teach == 1) {
        teach_rooms = n_doubles_teach + ' double room and '
      }

      if (n_singles_teach >= 2){
        teach_rooms = teach_rooms + n_singles_teach + ' single rooms'
      }
      else if (n_singles_teach == 1) {
        teach_rooms = teach_rooms + n_singles_teach + ' single room'
      }
    }
    else if (n_triples_teach != 0){
      if (n_triples_teach >= 2){
      teach_rooms = n_triples_teach + ' triple rooms'
      }
      else if (n_triples_teach == 1) {
        teach_rooms = n_triples_teach + ' triple room'
      }
    }
    else if (n_doubles_teach != 0){
      if (n_doubles_teach >= 2){
      teach_rooms = n_doubles_teach + ' double rooms'
      }
      else if (n_doubles_teach == 1) {
        teach_rooms = n_doubles_teach + ' double room'
      }
    }
    else if (n_singles_teach != 0){
      if (n_singles_teach >= 2){
      teach_rooms = n_singles_teach + ' single rooms'
      }
      else if (n_singles_teach == 1) {
        teach_rooms = n_singles_teach + ' single room'
      }
    }
  }

  else if (n_triples != 0){
    if (n_triples >= 2){
    total_rooms = n_triples + ' triple rooms'
    }
    else if (n_triples == 1) {
      total_rooms = n_triples + ' triple room'
    }

    //delegate rooms
    if (n_triples_del != 0){
      if (n_triples_del >= 2){
      del_rooms = n_triples_del + ' triple rooms'
      }
      else if (n_triples_del == 1) {
        del_rooms = n_triples_del + ' triple room'
      }
    }

    //teacher rooms
    if (n_triples_teach != 0){
      if (n_triples_teach >= 2){
      teach_rooms = n_triples_teach + ' triple rooms'
      }
      else if (n_triples_teach == 1) {
        teach_rooms = n_triples_teach + ' triple room'
      }
    }
  }

  else if (n_doubles != 0){
    if (n_doubles >= 2){
    total_rooms = n_doubles + ' double rooms'
    }
    else if (n_doubles == 1) {
      total_rooms = n_doubles + ' double room'
    }

    //delegate rooms
    if (n_doubles_del != 0){
      if (n_doubles_del >= 2){
      del_rooms = n_doubles_del + ' double rooms'
      }
      else if (n_doubles_del == 1) {
        del_rooms = n_doubles_del + ' double room'
      }
    }

     //teacher rooms
    if (n_doubles_teach != 0){
      if (n_doubles_teach >= 2){
      teach_rooms = n_doubles_teach + ' double rooms'
      }
      else if (n_doubles_teach == 1) {
        teach_rooms = n_doubles_teach + ' double room'
      }
    }
  }

  else if (n_singles != 0){
    if (n_singles >= 2){
    total_rooms = n_singles + ' single rooms'
    }
    else if (n_singles == 1) {
      total_rooms = n_singles + ' single room'
    }
    //delegate rooms
    if (n_singles_del != 0){
      if (n_singles_del >= 2){
      del_rooms = n_singles_del + ' single rooms'
      }
      else if (n_singles_del == 1) {
        del_rooms = n_singles_del + ' single room'
      }
    }
    //teacher rooms
    if (n_singles_teach != 0){
      if (n_singles_teach >= 2){
      teach_rooms = n_singles_teach + ' single rooms'
      }
      else if (n_singles_teach == 1) {
        teach_rooms = n_singles_teach + ' single room'
      }
    }
  }

  if (teach_rooms == ''){
    teach_rooms = 'none'
  }

  body.total_rooms = total_rooms
  body.delegate_rooms = del_rooms
  body.teacher_rooms = teach_rooms

  if (extranights == 0){
    body.extranights = 'no';
    body.night = 'nights'
  }
  else if (extranights == 1) {
    body.extranights = extranights;
    body.night = 'night'
  }
  else{
    body.extranights = extranights;
    body.night = 'nights'
  }

  body.dueday = duedate.split("/")[0];
  if (duedate.split("/")[0] == '1' || duedate.split("/")[0] == '11' || duedate.split("/")[0] == '21' || duedate.split("/")[0] == '31'){
    body.ordinal = 'st'}
  else if (duedate.split("/")[0] == '2' || duedate.split("/")[0] == '12' || duedate.split("/")[0] == '22'){
    body.ordinal = 'nd'}
  else if (duedate.split("/")[0] == '3' || duedate.split("/")[0] == '13' || duedate.split("/")[0] == '23'){
    body.ordinal = 'rd'}
  else {body.ordinal = 'th'};
  
  if (duedate.split("/")[1] == '01'){body.duemonth = 'January'}
  else if (duedate.split("/")[1] == '02'){body.duemonth = 'February'}
  else if (duedate.split("/")[1] == '03'){body.duemonth = 'March'}
  else if (duedate.split("/")[1] == '04'){body.duemonth = 'April'}
  else if (duedate.split("/")[1] == '05'){body.duemonth = 'May'}
  else if (duedate.split("/")[1] == '06'){body.duemonth = 'June'}
  else if (duedate.split("/")[1] == '07'){body.duemonth = 'July'}
  else if (duedate.split("/")[1] == '08'){body.duemonth = 'August'}
  else if (duedate.split("/")[1] == '09'){body.duemonth = 'September'}
  else if (duedate.split("/")[1] == '10'){body.duemonth = 'October'}
  else if (duedate.split("/")[1] == '11'){body.duemonth = 'November'}
  else if (duedate.split("/")[1] == '12'){body.duemonth = 'December'};

  body.dueyear = duedate.split("/")[2];

  if (n_teachers!=0){
    if (social == '0'){
      body.social = 'We noticed that you have still not signed up for our social event. We would like to remind you that this is a wonderful opportunity for all participants to network and make long-lasting connections with people from all around the world.'
    }
    else {
      body.social = 'Please note that we included all chaperones and delegates in the social event fees. If any delegates or chaperones are not willing to attend this event, please let us know, and we will resend the payment details.'
    }
  }
  else{body.social=''}

  if ((doc.split('_email')[0]=='AI_PD' && version_n=='01' && allinc_sheet.getRange(school_row,7,1,3).getBackground() != '#f4cccc')||(doc.split('_email')[0]=='STD_PD' && version_n=='01')||(doc.split('_')[1]=='wrongrooms')){
    body.intro = 'Thank you very much for your interest in BIMUN ESADE. My name is Marta Segovia, and I will be your contact person during your application process for BIMUN ESADE 2023.'
    body.intro2 = 'We are glad to confirm your participation in BIMUN ESADE 2023 on April 20th, 21st and 22nd.'

    if (doc.split('_')[1]=='PD'){
      body.end2 = ', as you stated in the form'
      body.intro3 = 'Please find attached the payment details for your school. '

      //Indicar quin fitxer s'ha d'attach
      const file = pd_folder.getFilesByName(pdfName+'.pdf').next();
      var pdID = file.getId();
      
      var pdFile = DriveApp.getFileById(pdID);

      if (doc.split('_email')[0]=='AI_PD'){
        var aiFile = DriveApp.getFileById(aiID);
        var pdf_attachments = [pdFile.getAs(MimeType.PDF), aiFile.getAs(MimeType.PDF)]
      }
      else{
        var pdf_attachments = [pdFile.getAs(MimeType.PDF)]
      }
      

      var threads = GmailApp.search('to:'+emailaddress+' subject:Welcome to BIMUN ESADE 2023')

      if (threads[0] == null){
        GmailApp.createDraft (
          emailaddress,
          'Welcome to BIMUN ESADE 2023',
        '',{
          htmlBody : body.evaluate().getContent() + signature,
          attachments: pdf_attachments
        });
      }

      else {
        if (doc.split('_email')[0]=='AI_PD'){
          body.intro = 'Thank you for your interest for the All Inclusive Pack.'
        }
        threads[0].createDraftReply(
        '',{
          htmlBody : body.evaluate().getContent() + signature,
          attachments: pdf_attachments
          }
        );

        var to = GmailApp.getDrafts()[0].getMessage().getTo()
        var subject = GmailApp.getDrafts()[0].getMessage().getSubject()

        if (to == 'BIMUN ESADE <bimun@barcelonamun.com>'){
          GmailApp.getDrafts()[0].update(
            emailaddress,
            subject,
          '',{
              htmlBody : body.evaluate().getContent() + signature,
              attachments: pdf_attachments
            })
        }
      }

      /** CORRECT VERSION
       * GmailApp.createDraft (
        emailaddress,
        'Welcome to BIMUN ESADE 2023',
      '',{
        htmlBody : body.evaluate().getContent() + signature,
        attachments: [pdFile.getAs(MimeType.PDF), aiFile.getAs(MimeType.PDF)]
      });
      */
    }

    else { //wrongrooms      
      //provisional version
      var threads = GmailApp.search('to:'+emailaddress+' subject:Welcome to BIMUN ESADE 2023')  

      if (threads[0] == null){
        GmailApp.createDraft (
          emailaddress,
          'Welcome to BIMUN ESADE 2023',
          '',{
            htmlBody : body.evaluate().getContent() + signature,
        });
      }
      else{
        threads[0].createDraftReply(
        '',{
          htmlBody : body.evaluate().getContent() + signature,
          }
        );

        var to = GmailApp.getDrafts()[0].getMessage().getTo()
        var subject = GmailApp.getDrafts()[0].getMessage().getSubject()

        if (to == 'BIMUN ESADE <bimun@barcelonamun.com>'){
          GmailApp.getDrafts()[0].update(
            emailaddress,
            subject,
            '',{
              htmlBody : body.evaluate().getContent() + signature,
            })
        }
      }

      /** CORRECT VERSION
       * GmailApp.createDraft (
        emailaddress,
        'Welcome to BIMUN ESADE 2023',
      '',{
        htmlBody : body.evaluate().getContent() + signature,
      });
      */
    }
  }

  else {//aquí també entren els AI_PD que tenien les habitacions malament (cel·la de color vermell)
    if (doc.split('_')[1]=='PD'){
      body.intro = 'Please find attached the new payment document with the specified changes.' 
      body.intro2 = 'We would like to remind you that'
      body.end2 = ''
      body.intro3 = ''

      //Indicar quin fitxer s'ha d'attach
      const file = pd_folder.getFilesByName(pdfName+'.pdf').next();
      var pdID = file.getId();
      
      var pdFile = DriveApp.getFileById(pdID);

      if (doc.split('_email')[0]=='AI_PD'){
        var aiFile = DriveApp.getFileById(aiID);
        var pdf_attachments = [pdFile.getAs(MimeType.PDF), aiFile.getAs(MimeType.PDF)]
      }
      else{
        var pdf_attachments = [pdFile.getAs(MimeType.PDF)]
      }

      var threads = GmailApp.search('to:'+emailaddress+' subject:Welcome to BIMUN ESADE 2023')  

      threads[0].createDraftReply(
      '',{
        htmlBody : body.evaluate().getContent() + signature,
        attachments: pdf_attachments
        }
      );

      var to = GmailApp.getDrafts()[0].getMessage().getTo()
      var subject = GmailApp.getDrafts()[0].getMessage().getSubject()

      if (to == 'BIMUN ESADE <bimun@barcelonamun.com>'){
        GmailApp.getDrafts()[0].update(
          emailaddress,
          subject,
        '',{
            htmlBody : body.evaluate().getContent() + signature,
            attachments: [pdf_attachments]
          })
      }
    }

    else {
      body.list_of_countries = 'https://docs.google.com/document/d/1R7sdiBryzruQrVM2wXgdyzCpsN8ILMOf7vsIaI9rUmk/edit?usp=sharing'
      body.ddf_file = ddf_ss_url
  
      var threads = GmailApp.search('to:'+emailaddress+' subject:Welcome to BIMUN ESADE 2023')  

      threads[0].createDraftReply(
      '',{
        htmlBody : body.evaluate().getContent() + signature,
        }
      );

      var to = GmailApp.getDrafts()[0].getMessage().getTo()
      var subject = GmailApp.getDrafts()[0].getMessage().getSubject()

      if (to == 'BIMUN ESADE <bimun@barcelonamun.com>'){
        GmailApp.getDrafts()[0].update(
          emailaddress,
          subject,
        '',{
            htmlBody : body.evaluate().getContent() + signature,
          })
      }
    }
  }
}

function MakePDF(school_type, pdfName) {
  //Crear un PDF de payment document i guardar-lo a la carpeta assignada (My Drive > DOCUMENTS > Payments > Payment documents)
    
  const fr = 0, fc = 0, lc = school_type.getLastColumn(), lr = school_type.getLastRow();
  const url = "https://docs.google.com/spreadsheets/d/" + schools_ss_id + "/export" +
    "?format=pdf&" +
    "size=7&" +
    "fzr=true&" +
    "portrait=true&" +
    "fitw=true&" +
    "gridlines=false&" +
    "printtitle=false&" +
    "top_margin=0.5&" +
    "bottom_margin=0.25&" +
    "left_margin=0.5&" +
    "right_margin=0.5&" +
    "sheetnames=false&" +
    "pagenum=UNDEFINED&" +
    "attachment=true&" +
    "gid=" + school_type.getSheetId() + '&' +
    "r1=" + fr + "&c1=" + fc + "&r2=" + lr + "&c2=" + lc;

  const params = { method: "GET", headers: { "authorization": "Bearer " + ScriptApp.getOAuthToken() } };
  const blob = UrlFetchApp.fetch(url, params).getBlob().setName(pdfName + '.pdf');

  pd_folder.createFile(blob);


 
}