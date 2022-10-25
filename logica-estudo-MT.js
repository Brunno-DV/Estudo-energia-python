const ss = SpreadsheetApp.getActiveSpreadsheet();
const planLogin = ss.getSheetByName("Login");
const planDados = ss.getSheetByName("Dados");
const planImpressao = ss.getSheetByName("Impressao");
const planLaudo = ss.getSheetByName("Laudo");
const planInformacoes = ss.getSheetByName("Informacoes");
const planSobre = ss.getSheetByName("Sobre");
const planGrafico = ss.getSheetByName("Grafico");
const planBaseDados = ss.getSheetByName("Base_de_dados");
const planEntrada = ss.getSheetByName("Entrada");
const planCalculos = ss.getSheetByName("Calculos");
const planConsulta_ZTC = ss.getSheetByName("Consulta_ZTC");
const planRevisao = ss.getSheetByName("Revisão");
const planCurvaBt = ss.getSheetByName("Curva_BT");
const planDirecional = ss.getSheetByName("Direcional");
const planCurvaFus = ss.getSheetByName("Curva_fusiveis");
const planMelhorias = ss.getSheetByName("Melhorias");


function seguranca(){

    senha = planLogin.getRange("B2").getValue();
    aux = planLogin.getRange("B3").getValue();
    aux1 = planLogin.getRange("B4").getValue();

    if(senha == 0){

      planLogin.showSheet();

      Utilities.sleep(1000);

      if(aux1 == 0){
        SpreadsheetApp.setActiveSheet(planLogin);
        planLogin.getRange("B3").setValue(0);
        planLogin.getRange("B4").setValue(1);
      }

      planDados.hideSheet();
      planImpressao.hideSheet();
      planLaudo.hideSheet();

    }else if(senha == 1){

      planDados.showSheet();
      planImpressao.showSheet();
      planLaudo.showSheet();

      Utilities.sleep(1000);
      
      if(aux == 0){
        SpreadsheetApp.setActiveSheet(planDados);
        planLogin.getRange("B3").setValue(1);
        planLogin.getRange("B4").setValue(0);
      }

      planLogin.hideSheet();

    } 
}

function telaLogin(){

  var template = HtmlService.createTemplateFromFile("useform");

  var html = template.evaluate();
  html.setTitle("SOFTWARE ESTUDO DE PROTEÇÃO DE MT").setHeight(350).setWidth (550);
  SpreadsheetApp.getUi().showModalDialog(html, "SOFTWARE ESTUDO DE PROTEÇÃO DE MT"); // TELA MÓVEL SEM SEGUNDO PLANO
}

function validacaoLogin(data){

  let validacao = [data.validacao];

  if(validacao == 1){

    planLogin.getRange("B2").setValue(1);
    seguranca();

  }
}

function futuro(){

  revisao = planDados.getRange("C25").getValue();
  data = planDados.getRange("C28").getValue();
  autor = planDados.getRange("C34").getValue();
  descricao = planDados.getRange("C31").getValue();

  rev00 = planImpressao.getRange("J43").getValue();
  rev01 = planImpressao.getRange("J44").getValue();
  rev02 = planImpressao.getRange("J45").getValue();
  rev03 = planImpressao.getRange("J46").getValue();

  if(rev00 == ""){
    planImpressao.getRange("D43").setValue(revisao);
    planImpressao.getRange("E43").setValue(data);
    planImpressao.getRange("F43").setValue(autor);
    planImpressao.getRange("J43").setValue(descricao);

  }else if(rev01 == ""){
    planImpressao.getRange("D44").setValue(revisao);
    planImpressao.getRange("E44").setValue(data);
    planImpressao.getRange("F44").setValue(autor);
    planImpressao.getRange("J44").setValue(descricao);

  }else if(rev02 == ""){
    planImpressao.getRange("D45").setValue(revisao);
    planImpressao.getRange("E45").setValue(data);
    planImpressao.getRange("F45").setValue(autor);
    planImpressao.getRange("J45").setValue(descricao);

  }else if(rev03 == ""){
    planImpressao.getRange("D46").setValue(revisao);
    planImpressao.getRange("E46").setValue(data);
    planImpressao.getRange("F46").setValue(autor);
    planImpressao.getRange("J46").setValue(descricao);

  }else if(rev03 != ""){
    planImpressao.getRange("D43").setValue(revisao);
    planImpressao.getRange("E43").setValue(data);
    planImpressao.getRange("F43").setValue(autor);
    planImpressao.getRange("J43").setValue(descricao);

    planImpressao.getRange("D44").setValue("");
    planImpressao.getRange("E44").setValue("");
    planImpressao.getRange("F44").setValue("");
    planImpressao.getRange("J44").setValue("");
    
    planImpressao.getRange("D45").setValue("");
    planImpressao.getRange("E45").setValue("");
    planImpressao.getRange("F45").setValue("");
    planImpressao.getRange("J45").setValue("");

    planImpressao.getRange("D46").setValue("");
    planImpressao.getRange("E46").setValue("");
    planImpressao.getRange("F46").setValue("");
    planImpressao.getRange("J46").setValue("");
  }

  pdf();

  Utilities.sleep(2000);

  planLogin.getRange("B2").setValue(0);

  seguranca();
}

function novo(){

  planDados.getRange("C6").setValue("");
  planDados.getRange("C9").setValue("");
  planDados.getRange("C12").setValue("");
  planDados.getRange("C16").setValue("");
  planDados.getRange("C19").setValue("");
  planDados.getRange("C22").setValue("");
  planDados.getRange("C25").setValue("");
  planDados.getRange("C28").setValue("");
  planDados.getRange("C31").setValue("");
  planDados.getRange("C34").setValue("");
  planDados.getRange("J12").setValue("");
  planDados.getRange("J18").setValue("");
  planDados.getRange("J24").setValue("");
  planDados.getRange("J27").setValue("");
  planDados.getRange("J30").setValue("");
  planDados.getRange("J33").setValue("");
  planDados.getRange("J36").setValue("");
  planDados.getRange("J39").setValue("");
  planDados.getRange("J45").setValue("");
  planDados.getRange("J48").setValue("");
  planDados.getRange("S7").setValue("");
  planDados.getRange("S8").setValue("");
  planDados.getRange("S9").setValue("");
  planDados.getRange("S15").setValue("");
  planDados.getRange("S16").setValue("");
  planDados.getRange("S17").setValue("");
  planDados.getRange("S23").setValue("");
  planDados.getRange("S24").setValue("");
  planDados.getRange("S25").setValue("");
  planDados.getRange("S31").setValue("");
  planDados.getRange("S32").setValue("");
  planDados.getRange("S33").setValue("");
  planDados.getRange("AA11").setValue("");
  planDados.getRange("AA13").setValue("");
  planDados.getRange("AA14").setValue("");
  planDados.getRange("Z16").setValue("");
  planDados.getRange("Z18").setValue("");
  planDados.getRange("AG14").setValue("");
  planDados.getRange("AG15").setValue("");
  planDados.getRange("AG16").setValue("");
  planDados.getRange("AG17").setValue("");
  planDados.getRange("AF20").setValue("");
  planDados.getRange("AG20").setValue("");
  planDados.getRange("AG25").setValue("");
  planDados.getRange("AG26").setValue("");
  planDados.getRange("AG27").setValue("");
  planDados.getRange("AG28").setValue("");
  planDados.getRange("AI26").setValue("");
  planDados.getRange("AI27").setValue("");
  planDados.getRange("AI28").setValue("");
  planDados.getRange("T53").setValue("");
  planDados.getRange("T54").setValue("");
  planDados.getRange("T55").setValue("");
  planDados.getRange("T56").setValue("");
  planDados.getRange("T57").setValue("");
  planDados.getRange("T60").setValue("");
  planDados.getRange("T61").setValue("");
  planDados.getRange("T62").setValue("");
  planDados.getRange("T63").setValue("");
  planDados.getRange("T64").setValue("");
  planDados.getRange("Z54").setValue("");
  planDados.getRange("Z55").setValue("");
  planDados.getRange("Z56").setValue("");
  planDados.getRange("Z61").setValue("");
  planDados.getRange("Z62").setValue("");
  planDados.getRange("Z63").setValue("");
  
  planImpressao.getRange("D41").setValue("");
  planImpressao.getRange("E41").setValue("");
  planImpressao.getRange("F41").setValue("");
  planImpressao.getRange("J41").setValue("");

  planImpressao.getRange("D42").setValue("");
  planImpressao.getRange("E42").setValue("");
  planImpressao.getRange("F42").setValue("");
  planImpressao.getRange("J42").setValue("");
    
  planImpressao.getRange("D43").setValue("");
  planImpressao.getRange("E43").setValue("");
  planImpressao.getRange("F43").setValue("");
  planImpressao.getRange("J43").setValue("");

  planImpressao.getRange("D44").setValue("");
  planImpressao.getRange("E44").setValue("");
  planImpressao.getRange("F44").setValue("");
  planImpressao.getRange("J44").setValue("");
  
}

function ztc(){

  planConsulta_ZTC.showSheet();
  SpreadsheetApp.setActiveSheet(planConsulta_ZTC);
  
}

function volta_ztc(){

  SpreadsheetApp.setActiveSheet(planDados);
  planConsulta_ZTC.hideSheet();

}

function info(){

  planInformacoes.showSheet();
  SpreadsheetApp.setActiveSheet(planInformacoes);
  
}

function volta_info(){

  SpreadsheetApp.setActiveSheet(planDados);
  planInformacoes.hideSheet();

}

function sobre(){

  planSobre.showSheet();
  SpreadsheetApp.setActiveSheet(planSobre);
  
}

function volta_sobre(){

  SpreadsheetApp.setActiveSheet(planDados);
  planSobre.hideSheet();

}

function pdf(){

  var nomePDF = planImpressao.getRange("P22").getValue();

  planLogin.getRange("B2").setValue(2);

  planDados.hideSheet();
  planLaudo.hideSheet();

  Utilities.sleep(2000);

  var folderIter = DriveApp.getFoldersByName("Estudos");
  var pdfFolder = folderIter.next();

  var spredsheet_id = ss.getId();
  var spredsheetFile = DriveApp.getFileById(spredsheet_id);
  var blob = spredsheetFile.getAs (MimeType.PDF);
  pdfFolder.createFile(blob).setName(nomePDF);

  planLaudo.showSheet();

  nomePDF = planLaudo.getRange("R18").getValue();

  planImpressao.hideSheet();

  Utilities.sleep(2000);

  folderIter = DriveApp.getFoldersByName("Estudos");
  pdfFolder = folderIter.next();

  spredsheet_id = ss.getId();
  spredsheetFile = DriveApp.getFileById(spredsheet_id);
  blob = spredsheetFile.getAs (MimeType.PDF);
  pdfFolder.createFile(blob).setName(nomePDF);

  planDados.showSheet();
  planImpressao.showSheet();

  Utilities.sleep(2000);
  
  planLogin.getRange("B2").setValue(1);

}