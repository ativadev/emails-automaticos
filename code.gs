//INICIALIZAÇÃO DAS VARIÁVEIS GLOBAIS
//EMAIL
globalThis.NOME_EMAIL = 
globalThis.EMAIL_OPERADOR =  //emails de destino, separados por vírgula

//BÁSICO DO BÁSICO, NÃO ALTERAR
globalThis.PLANILHA = SpreadsheetApp.getActiveSpreadsheet();
globalThis.UI = SpreadsheetApp.getUi();
//globalThis.FOLDER_ID =

//ABA COM OS DADOS
globalThis.NOME_ABA_COM_OS_DADOS = 
globalThis.ABA_COM_OS_DADOS = PLANILHA.getSheetByName(NOME_ABA_COM_OS_DADOS)


function onOpen(){

  globalThis.UI = SpreadsheetApp.getUi()

  UI.createMenu("ENVIO DE EMAILS")
  .addItem("EMAIL", "EMAIL")
  .addToUi();
}

//EMAILS
function EMAIL(){

  //nome da aba que vai no email
  var NOME_ABA_A_SER_IMPRESSA =

  //nome do modelo de email
  var ARQUIVO_MODELO = 

  Logger.log(Session.getActiveUser().getEmail());

  console.log("Função rodando");

  //destinatários
  var EMAIL_DO_CLIENTE = String(ABA_COM_OS_DADOS.getRange(__LINHA__,__COLUNA__).getValue());
  var EMAIL_DO_REPRESENTANTE1 = String(ABA_COM_OS_DADOS.getRange(__LINHA__,__COLUNA__).getValue());
  var EMAIL_DO_REPRESENTANTE2 = String(ABA_COM_OS_DADOS.getRange(__LINHA__,__COLUNA__).getValue());
  var DESTINATÁRIOS = String(EMAIL_DO_CLIENTE) + ", " + String(EMAIL_DO_REPRESENTANTE1) + ", " + String(EMAIL_DO_REPRESENTANTE2);

  var resposta = UI.alert("Deseja enviar um e-mail para " +DESTINATÁRIOS+"?", UI.ButtonSet.OK_CANCEL);


  //caso o usuário escolha "Cancelar", a execução para
  if(resposta != UI.Button.OK){
    Logger.log("Usuário cancelou");
    return
    }

  console.log("Usuário confirmou")

  //aba que vai no email
  var ABA_A_SER_IMPRESSA = PLANILHA.getSheetByName(NOME_ABA_A_SER_IMPRESSA);

  //número do pedido
  var NUM_PEDIDO = ABA_A_SER_IMPRESSA.getRange(__LINHA__, __COLUNA__).getValue();
  
  Logger.log("PEDIDO "+NUM_PEDIDO);

  //gera o assunto
  var ASSUNTO_EMAIL = 

  //nome do pdf
  var NOME_DO_PDF = 

  //puxa o corpo do email do arquivo .html
  var CORPO = HtmlService.createHtmlOutputFromFile(ARQUIVO_MODELO).getContent();


  //define a mensagem, aba como anexo
  var mensagem = {
    to: EMAIL_OPERADOR,
    cc: DESTINATÁRIOS,
    subject: ASSUNTO_EMAIL,
    htmlBody: CORPO,
    name: NOME_EMAIL,
    attachments: [gerarBLOB(ABA_A_SER_IMPRESSA, NOME_DO_PDF)]
  }

  //envia o email
  enviarEMAIL(mensagem)

}


//FUNÇÕES BÁSICAS DO SCRIPT
//função que envia os emails, pra clareza
function enviarEMAIL(mensagem){
  var restante = MailApp.getRemainingDailyQuota();
  Logger.log("Envios restantes: "+String(restante));
  if (restante>0){
    MailApp.sendEmail(mensagem);
    Logger.log("Email enviado para " + mensagem["to"] + " e " + mensagem["cc"]);
    PLANILHA.toast("Email enviado para " + mensagem["to"] + " e " + mensagem["cc"]);
  } else {
    UI.alert("Ops!..", "Sinto muito, você atingiu seu limite de e-mails de hoje...tente enviar utilizando outro usuário (@gmail.com)!", UI.ButtonSet.OK)
    Logger.log("Limite excedido!")
    return
  }
}

//função que gera o pdf
function gerarBLOB(ABA, NOME) {
  
  var url = 'https://docs.google.com/spreadsheets/d/' + PLANILHA.getId() + '/export?exportFormat=pdf&format=pdf'
    +    '&size=7'
    +    '&portrait=true'
    +    '&scale=4'//PRO CONTEÚDO SE AJUSTAR À PAGINA
    +    '&sheetnames=true&printtitle=false' // hide optional headers and footers
    +    '&pagenum=RIGHT&gridlines=false' // hide page numbers and gridlines
    +    '&fzr=false' // do not repeat row headers (frozen rows) on each page
    +    '&horizontal_alignment=CENTER' //LEFT/CENTER/RIGHT
    +    '&vertical_alignment=TOP' //TOP/MIDDLE/BOTTOM
    +    '&gid=' + ABA.getSheetId(); // the sheet's Id
  
  var token = ScriptApp.getOAuthToken();

  // request export url
  var response = UrlFetchApp.fetch(url, {headers: {'Authorization': 'Bearer ' + token}});

  var theBlob = response.getBlob().setName(NOME);

  console.log("PDF da aba "+ ABA.getName() + " foi gerado")
  
  return theBlob;
};



/*
//FUNÇÃO DE SALVAMENTO EM PDF (OPCIONAL)
function salvarPDF(){

  Logger.log(Session.getActiveUser().getEmail());

  console.log("Função rodando")

  var PASTA_COMPARTILHADA = DriveApp.getFolderById(FOLDER_ID);

  var dir = PASTA_COMPARTILHADA

  var dados = gerarBLOB();

  dir.createFile(dados);

  UI.alert("Arquivo salvo em https://drive.google.com/drive/folders/"+FOLDER_ID, UI.ButtonSet.OK)

}
*/


