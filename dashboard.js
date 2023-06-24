var dash_ss_id = '1yFRCS5g5sdw9WKekk_nZEDyi5koUHRJaohBdSXVU2S4';
var dash_ss = SpreadsheetApp.openById(dash_ss_id);
var dash_data = dash_ss.getSheetByName('DATA');
var dash_sheet = dash_ss.getSheetByName('DASHBOARD');

var schools_ss_id = '1b2BWKdnBoTC7DqPE0OjbDH6-2Q3_sF3oM1fRhNm7i8s';
var schools_ss = SpreadsheetApp.openById(schools_ss_id);
var schools_ai_sheet = schools_ss.getSheetByName('all inc');
var schools_std_sheet = schools_ss.getSheetByName('standard');

var ind_ss_id = '1jQehZiyuR0Ht3XK-dmqLIXZcIIyBxYl_Wm8IvtsqC4U';
var ind_ss = SpreadsheetApp.openById(ind_ss_id);
var ind_ai_sheet = ind_ss.getSheetByName('all inc');
var ind_std_sheet = ind_ss.getSheetByName('standard');

var dash_ss_id = '1yFRCS5g5sdw9WKekk_nZEDyi5koUHRJaohBdSXVU2S4';
var dash_ss = SpreadsheetApp.openById(dash_ss_id);
var dash_data = dash_ss.getSheetByName('DATA');
var dash_sheet = dash_ss.getSheetByName('DASHBOARD');
var dash_attendance = dash_ss.getSheetByName('ATTENDANCE');
var dash_folder_id = '1EeUO2THNofjTeho9lAB_3FLNRz_PR3sE'

var schools_ss_id = '1b2BWKdnBoTC7DqPE0OjbDH6-2Q3_sF3oM1fRhNm7i8s';
var schools_ss = SpreadsheetApp.openById(schools_ss_id);
var schools_ai_sheet = schools_ss.getSheetByName('all inc');
var schools_std_sheet = schools_ss.getSheetByName('standard');

var ind_ss_id = '1jQehZiyuR0Ht3XK-dmqLIXZcIIyBxYl_Wm8IvtsqC4U';
var ind_ss = SpreadsheetApp.openById(ind_ss_id);
var ind_ai_sheet = ind_ss.getSheetByName('all inc');
var ind_std_sheet = ind_ss.getSheetByName('standard');

function dash() {

  /** AI schools */
  var AI_s_lastrow = schools_ai_sheet.getLastRow();
  var num_ai_schools = 0;
  
  //number of schools
  var SAIvals = schools_ai_sheet.getRange(3,2,AI_s_lastrow,1).getValues();
  var total_num_ai_schools = SAIvals.findIndex(c=>c[0]=='');
  var SAIvals_colour = schools_ai_sheet.getRange(3,2,total_num_ai_schools,1).getBackgrounds();
  for (var i=0;i<SAIvals_colour.length;i++){
    SAIvals_colour[i] = SAIvals_colour[i][0];
    if (SAIvals_colour[i] == '#6d9eeb'){
      num_ai_schools = num_ai_schools + 1
    }
  }

  //number of delegates
  var colour = 0
  var DAIvals = schools_ai_sheet.getRange(3,5,total_num_ai_schools,1).getValues();
  for (var i=0;i<DAIvals.length;i++){
    DAIvals[i] = DAIvals[i][0];
    colour = i + 3
    if (schools_ai_sheet.getRange(colour,5).getBackground() != '#6d9eeb'){DAIvals[i] = 0}}
  var num_ai_del = Number(DAIvals.reduce((a, b) => a + b, 0));

  //number of teachers
  colour = 0
  var TAIvals = schools_ai_sheet.getRange(3,6,total_num_ai_schools,1).getValues();
  for (var i=0;i<TAIvals.length;i++){
    TAIvals[i] = TAIvals[i][0];
    colour = i + 3
    if (schools_ai_sheet.getRange(colour,6).getBackground() != '#6d9eeb'){TAIvals[i] = 0}}
  var num_ai_teach = Number(TAIvals.reduce((a, b) => a + b, 0));

  //number of rooms
  var TripleAIvals = schools_ai_sheet.getRange(3,7,total_num_ai_schools,1).getValues();
  var DoubleAIvals = schools_ai_sheet.getRange(3,8,total_num_ai_schools,1).getValues();
  var SingleAIvals = schools_ai_sheet.getRange(3,9,total_num_ai_schools,1).getValues();

  colour = 0
  for (var i=0;i<TripleAIvals.length;i++){
    TripleAIvals[i] = TripleAIvals[i][0];
    colour = i + 3
    if (schools_ai_sheet.getRange(colour,7).getBackground() != '#b7e1cd'){TripleAIvals[i] = 0}}
  colour = 0
  for (var i=0;i<DoubleAIvals.length;i++){
    DoubleAIvals[i] = DoubleAIvals[i][0];
    colour = i + 3
    if (schools_ai_sheet.getRange(colour,8).getBackground() != '#b7e1cd'){DoubleAIvals[i] = 0}}
  colour = 0
  for (var i=0;i<SingleAIvals.length;i++){
    SingleAIvals[i] = SingleAIvals[i][0];
    colour = i + 3
    if (schools_ai_sheet.getRange(colour,9).getBackground() != '#b7e1cd'){SingleAIvals[i] = 0}}

  var num_rooms = Number(TripleAIvals.reduce((a, b) => a + b, 0))+Number(DoubleAIvals.reduce((a, b) => a + b, 0))+Number(SingleAIvals.reduce((a, b) => a + b, 0));
  
  //number of socials
  colour = 0
  var SocAIvals = schools_ai_sheet.getRange(3,11,total_num_ai_schools,1).getValues();
  for (var i=0;i<SocAIvals.length;i++){
    SocAIvals[i] = SocAIvals[i][0];
    colour = i + 3
    if (schools_ai_sheet.getRange(colour,2).getBackground() != '#6d9eeb'){SocAIvals[i] = 0}}
  var num_soc = Number(SocAIvals.reduce((a, b) => a + b, 0));

  /** Standard schools */
  var STD_s_lastrow = schools_std_sheet.getLastRow();
  var num_std_schools = 0;
  
  //number of schools
  var SSTDvals = schools_std_sheet.getRange(2,2,STD_s_lastrow,1).getValues();
  var total_num_std_schools = Number(SSTDvals.findIndex(c=>c[0]==''));
  var SSTDvals_colour = schools_std_sheet.getRange(2,2,total_num_std_schools,1).getBackgrounds();
  for (var i=0;i<SSTDvals_colour.length;i++){
    SSTDvals_colour[i] = SSTDvals_colour[i][0];
    if (SSTDvals_colour[i] == '#6d9eeb'){
      num_std_schools = num_std_schools + 1
    }
  }

  //number of delegates
  colour = 0
  var DSTDvals = schools_std_sheet.getRange(2,5,total_num_std_schools,1).getValues();
  for (var i=0;i<DSTDvals.length;i++){
    DSTDvals[i] = DSTDvals[i][0];
    colour = i + 2;
    if (schools_std_sheet.getRange(colour,5).getBackground()!= '#6d9eeb'){DSTDvals[i] = 0}}
  var num_std_del = Number(DSTDvals.reduce((a, b) => a + b, 0));

  //number of teachers
  colour = 0
  var TSTDvals = schools_std_sheet.getRange(2,6,total_num_std_schools,1).getValues();
  for (var i=0;i<TSTDvals.length;i++){
    TSTDvals[i] = TSTDvals[i][0];
    colour = i + 2;
    if (schools_std_sheet.getRange(colour,6).getBackground()!= '#6d9eeb'){TSTDvals[i] = 0}}
  var num_std_teach = Number(TSTDvals.reduce((a, b) => a + b, 0));

  //number of socials
  colour = 0
  var SocSTDvals = schools_std_sheet.getRange(2,7,total_num_std_schools,1).getValues();
  for (var i=0;i<SocSTDvals.length;i++){
    SocSTDvals[i] = SocSTDvals[i][0];
    colour = i + 2
    if (schools_std_sheet.getRange(colour,2).getBackground() != '#6d9eeb'){SocSTDvals[i] = 0;}}
  num_soc = num_soc + Number(SocSTDvals.reduce((a, b) => a + b, 0));

  /** AI individuals */
  var AI_i_lastrow = ind_ai_sheet.getLastRow();
  var num_ai_ind = 0;
  
  //number of individuals
  var IAIvals = ind_ai_sheet.getRange(2,2,AI_i_lastrow,1).getValues();
  var total_num_ai_ind = Number(IAIvals.findIndex(c=>c[0]==''));
  var IAIvals_colour = ind_ai_sheet.getRange(2,2,total_num_ai_ind,1).getBackgrounds();
  for (var i=0; i<IAIvals_colour.length;i++){
    IAIvals_colour[i] = IAIvals_colour[i][0];
    if (IAIvals_colour[i] == '#6d9eeb'){
      num_ai_ind = num_ai_ind + 1
    }
  }

  //number of rooms
  colour = 0
  var Roomsindvals = ind_ai_sheet.getRange(2,4,total_num_ai_ind,1).getValues();
  for (var i=0;i<Roomsindvals.length;i++){
    Roomsindvals[i] = Roomsindvals[i][0];
    colour = i + 2
    if (ind_ai_sheet.getRange(colour,2).getBackground() != '#6d9eeb'){Roomsindvals[i] = 0}
    else{
      if (Roomsindvals[i] =='Single'){
        Roomsindvals[i] = 1;
      }
      else if (Roomsindvals[i] =='Double') {
        Roomsindvals[i] = 0.5;
      }
      else if (Roomsindvals[i] =='Triple') {
        Roomsindvals[i] = 1/3;
      }
    }
  }

  num_rooms = num_rooms + Number(Roomsindvals.reduce((a, b) => a + b, 0));
  
  //number of socials
  colour = 0
  var SocAIindvals = ind_ai_sheet.getRange(2,6,total_num_ai_ind,1).getValues();
  for (var i=0;i<SocAIindvals.length;i++){
    SocAIindvals[i] = SocAIindvals[i][0];
    colour = i + 2
    if (ind_ai_sheet.getRange(colour,2).getBackground() != '#6d9eeb' || SocAIindvals[i] != 'Yes'){SocAIindvals[i] = 0}
    else{SocAIindvals[i] = 1}
  }
  num_soc = num_soc + Number(SocAIindvals.reduce((a, b) => a + b, 0));

  /** Standard individuals */
  var STD_i_lastrow = ind_std_sheet.getLastRow();
  var num_std_ind = 0;

  //number of individuals
  var ISTDvals = ind_std_sheet.getRange(2,2,STD_i_lastrow,1).getValues();
  var total_num_std_ind = Number(ISTDvals.findIndex(c=>c[0]==''));
  var ISTDvals_colour = ind_std_sheet.getRange(2,2,total_num_std_ind,1).getBackgrounds();
  for (var i=0; i<ISTDvals_colour.length;i++){
    ISTDvals_colour[i] = ISTDvals_colour[i][0];
    if (ISTDvals_colour[i] == '#6d9eeb'){
      num_std_ind = num_std_ind + 1
    }
  }

  //number of socials
  colour = 0
  var SocSTDindvals = ind_std_sheet.getRange(2,4,total_num_std_ind,1).getValues();
  for (var i=0;i<SocSTDindvals.length;i++){
    SocSTDindvals[i] = SocSTDindvals[i][0];
    colour = i + 2
    if (ind_std_sheet.getRange(colour,2).getBackground() != '#6d9eeb' || SocSTDindvals[i] !='Yes'){SocSTDindvals[i] = 0}
    else {SocSTDindvals[i] = 1}
  }
  num_soc = num_soc + Number(SocSTDindvals.reduce((a, b) => a + b, 0));

  /**AI PAID schools */
  var num_ai_schools_paid = 0
  var num_ai_del_paid = 0
  var num_ai_teach_paid = 0
  var num_rooms_paid = 0
  var num_soc_paid = 0
  var p = 0

  for (var r = 0; r<total_num_ai_schools; r++){
    p = r + 3;
    if (schools_ai_sheet.getRange(p,14).getValue()=='Yes' && (schools_ai_sheet.getRange(p,14).getBackground()=='#ffffff' || schools_ai_sheet.getRange(p,14).getBackground()=='#e0f7fa')){
      num_ai_schools_paid = num_ai_schools_paid + 1;
      num_ai_del_paid = num_ai_del_paid + schools_ai_sheet.getRange(p,5).getValue();
      num_ai_teach_paid = num_ai_teach_paid + schools_ai_sheet.getRange(p,6).getValue();
      num_rooms_paid = num_rooms_paid + schools_ai_sheet.getRange(p,7).getValue() + schools_ai_sheet.getRange(p,8).getValue() + schools_ai_sheet.getRange(p,9).getValue();
      num_soc_paid = num_soc_paid + schools_ai_sheet.getRange(p,11).getValue();
    }
  };
  
  /**STD PAID schools */
  var num_std_schools_paid = 0
  var num_std_del_paid = 0
  var num_std_teach_paid = 0
  p = 0

  for (r = 0; r<total_num_std_schools; r++){
    p = r + 2
    if (schools_std_sheet.getRange(p,10).getValue()=='Yes' && (schools_std_sheet.getRange(p,10).getBackground()=='#ffffff' || schools_std_sheet.getRange(p,10).getBackground()=='#e0f7fa')){
      num_std_schools_paid = num_std_schools_paid + 1;
      num_std_del_paid = num_std_del_paid + schools_std_sheet.getRange(p,5).getValue();
      num_std_teach_paid = num_std_teach_paid + schools_std_sheet.getRange(p,6).getValue();
      num_soc_paid = num_soc_paid + schools_std_sheet.getRange(p,7).getValue();
    }
  };

  /**AI PAID individuals */
  var num_ai_ind_paid = 0
  p = 0

  for (r = 0; r<total_num_ai_ind; r++){
    p = r + 2
    if (ind_ai_sheet.getRange(p,9).getValue()=='Yes' && (ind_ai_sheet.getRange(p,9).getBackground()=='#ffffff' || ind_ai_sheet.getRange(p,9).getBackground()=='#e0f7fa')){
      num_ai_ind_paid = num_ai_ind_paid + 1;

      if (ind_ai_sheet.getRange(p,4).getValue()=='Single'){
        num_rooms_paid = num_rooms_paid + 1
      }
      else if (ind_ai_sheet.getRange(p,4).getValue()=='Double') {
        num_rooms_paid = num_rooms_paid + 0.5;
      }
      else if (ind_ai_sheet.getRange(p,4).getValue()=='Triple') {
        num_rooms_paid = num_rooms_paid + 1/3;
      };
      
      if (ind_ai_sheet.getRange(p,6).getValue()=='Yes'){
        num_soc_paid = num_soc_paid + 1
      };
    }
  }

  /**STD PAID individuals */
  var num_std_ind_paid = 0
  p = 0

  for (r = 0; r<total_num_std_ind; r++){
    p = r + 2;
    if (ind_std_sheet.getRange(p,8).getValue()=='Yes' && (ind_std_sheet.getRange(p,8).getBackground()=='#ffffff' || ind_std_sheet.getRange(p,8).getBackground()=='#e0f7fa')){
      num_std_ind_paid = num_std_ind_paid + 1;
      if (ind_std_sheet.getRange(p,4).getValue()=='Yes'){
        num_soc_paid = num_soc_paid + 1
      }
    }
  }

  /**OMPLIR DASHBOARD */

  dash_sheet.getRange('B5').setValue(num_ai_del+num_std_del+num_ai_ind+num_std_ind);
  dash_sheet.getRange('B6').setValue(num_ai_del_paid+num_std_del_paid+num_ai_ind_paid+num_std_ind_paid);

  dash_sheet.getRange('A9').setValue(num_ai_teach+num_std_teach);
  dash_sheet.getRange('B9').setValue(num_ai_teach_paid+num_std_teach_paid);

  dash_sheet.getRange('A12').setValue(num_ai_schools_paid);
  dash_sheet.getRange('B12').setValue(num_std_schools_paid);

  dash_sheet.getRange('D5').setValue(num_ai_del);
  dash_sheet.getRange('E5').setValue(num_std_del);
  dash_sheet.getRange('D6').setValue(num_ai_ind);
  dash_sheet.getRange('E6').setValue(num_std_ind);

  dash_sheet.getRange('D8').setValue(num_ai_del_paid);
  dash_sheet.getRange('E8').setValue(num_std_del_paid);
  dash_sheet.getRange('D9').setValue(num_ai_ind_paid);
  dash_sheet.getRange('E9').setValue(num_std_ind_paid);

  if ((num_ai_schools+num_std_schools)==0){
    dash_sheet.getRange('C12').setValue(0)
  }
  else {
    dash_sheet.getRange('C12').setValue((num_ai_schools_paid+num_std_schools_paid)/(num_ai_schools+num_std_schools));  
  }

  dash_sheet.getRange('H6').setValue(Math.ceil(num_rooms)); //round UP to next integer
  dash_sheet.getRange('I6').setValue(Math.ceil(num_rooms_paid)); //round UP to next integer

  dash_sheet.getRange('J11').setValue(num_soc);
  dash_sheet.getRange('J12').setValue(num_soc_paid);

  /** CHART UPDATE */

  var date = dash_sheet.getRange('D2').getValue();
  var data_lastrow = dash_data.getLastRow();

  for (var i = 2; i<=data_lastrow; i++){
    if (date>=dash_data.getRange(i,1).getValue() && date<=dash_data.getRange(i,2).getValue()){
      dash_data.getRange(i,6).setValue(dash_sheet.getRange('B5').getValue());
      dash_data.getRange(i,7).setValue(dash_sheet.getRange('B6').getValue());
      break
    }
  }

  //SOLD-OUT CHART

  if (dash_data.getRange(i,8).getValue()<= 1/3){
    var soldout_bar_colour = '#000B3D'
  }
  else if (dash_data.getRange(i,8).getValue()<= 2/3){
    var soldout_bar_colour = '#F9B937'
  }
  else{
    var soldout_bar_colour = '#F1595C'
  }

  var charts = dash_sheet.getCharts();

  for (var c in charts) {
    var chart = charts[c];
    if (chart.getOptions().get('height') == 50){
      dash_sheet.removeChart(chart)
      break
    }
  }

  var soldout_chart = dash_sheet.newChart()
  .asBarChart()
  .addRange(dash_data.getRange(i,8))
  .setTransposeRowsAndColumns(false)
  .setOption('legend.position', 'none')
  .setOption('theme', 'maximized')
  .setOption('hAxis',{textStyle:{fontSize:0},gridlines:{count:0},minValue:0,maxValue:1})
  .setOption('vAxes.0.textStyle.fontSize', 0)
  .setOption('hAxis.gridlines.count',0)
  .setOption('series', {0: {color: soldout_bar_colour,dataLabel: 'value'}})
  .setOption('backgroundColor', {stroke:'none', strokeWidth:0, fill: 'none'})
  .setOption('height', 50)
  .setOption('width', 797)
  .setPosition(13, 3, 106, 0)
  .build();
  dash_sheet.insertChart(soldout_chart);
}

function attendance () {

  /**AGAFAR DATA */

  var date = dash_sheet.getRange('D2').getValue();
  var attendance_lastrow = dash_attendance.getLastRow();

  for (var i = 2; i<=attendance_lastrow; i++){
    if (date<=dash_attendance.getRange(i,1).getValue()){
      i = i - 1
      break
    }
  }

  for (var p = 2; p<9; p++){
    var Attvals = dash_attendance.getRange(2,p,i,1).getValues();
      for (var a=0;a<Attvals.length;a++){
        Attvals[a] = Attvals[a][0];
        if (Attvals[a] =='Yes'){
          Attvals[a] = 1;
        }
        else if (Attvals[a] =='Late'){
          Attvals[a] = 0.5;
        }
        else {
          Attvals[a] = 0;
        }
      };

    var attendance_percentage = Number(Attvals.reduce((a, b) => a + b, 0))/i;

    dash_attendance.getRange(p,10,1,1).setValue(attendance_percentage);
  }
}

function open(){
  dash()

  var ui = SpreadsheetApp.getUi();
  var htmlOutput = HtmlService
    .createHtmlOutput('<p>Update done!</p>')
    .setWidth(250)
    .setHeight(50);
  ui.showModelessDialog(htmlOutput, 'Dashboard update');
}

function sendEmail () {

  dash()

  attendance()

  var date = dash_sheet.getRange('D2').getDisplayValue();

  /**Generar el Payment document*/
  //Crear un PDF de payment document i guardar-lo a la carpeta assignada (My Drive > DOCUMENTS > Logística i Organización interna >Actas > DASHBOARD)
    
  const fr = 0, fc = 0, lc = 40, lr = 40;
  const url = "https://docs.google.com/spreadsheets/d/" + dash_ss_id + "/export" +
    "?format=pdf&" +
    "size=7&" +
    "fzr=true&" +
    "portrait=false&" +
    "fitw=true&" +
    "gridlines=false&" +
    "printtitle=false&" +
    "top_margin=0.08&" +
    "bottom_margin=0&" +
    "left_margin=0.5&" +
    "right_margin=0.5&" +
    "sheetnames=false&" +
    "pagenum=UNDEFINED&" +
    "attachment=true&" +
    "gid=" + dash_sheet.getSheetId() + '&' +
    "r1=" + fr + "&c1=" + fc + "&r2=" + lr + "&c2=" + lc;

  const params = { method: "GET", headers: { "authorization": "Bearer " + ScriptApp.getOAuthToken() } };
  var file_name = 'BIMUN_ESADE_2023_DASHBOARD_'+ date + '.pdf'
  const blob = UrlFetchApp.fetch(url, params).getBlob().setName(file_name);

  var dash_folder = DriveApp.getFolderById(dash_folder_id);

  dash_folder.createFile(blob);
    
  /**Enviar email amb  PDF attached */

  //Indicar quin fitxer s'ha d'attach
  const file = dash_folder.getFilesByName(file_name).next();
  var dashID = file.getId();
  
  var dashFile = DriveApp.getFileById(dashID);

  GmailApp.sendEmail (
    'torresmasdeu@gmail.com',
    'BIMUN ESADE 2023 DASHBOARD - '+date,
    "Please find attached this week's dashboard, participation rates and tasks.",{
    attachments: [dashFile.getAs(MimeType.PDF)]
  });
}