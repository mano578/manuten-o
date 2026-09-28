// ================= FUNÇÕES ESSENCIAIS DO SISTEMA =================
function doGet() {
  // Agora sim, rigorosamente igual ao nome do ficheiro (minúsculo)
  return HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .setTitle('Sistema Inovar | Dashboard')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
// ================= UTILITÁRIOS GERAIS =================
function formatarDataInputHtml(dataStr) {
  if (!dataStr) return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");
  try {
    var partes = dataStr.split('T');
    if (partes.length === 2) {
      var dataPartes = partes[0].split('-'); var horaPartes = partes[1].split(':');
      return dataPartes[2] + "/" + dataPartes[1] + "/" + dataPartes[0] + " " + horaPartes[0] + ":" + horaPartes[1] + ":00";
    } else {
      var p = dataStr.split('-'); return p[2] + "/" + p[1] + "/" + p[0];
    }
  } catch (e) {}
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");
}

// ================= VALIDAÇÃO DE LOGIN =================
function validarLogin(usuario, senha) {
  var planilha = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("USUARIOS_SISTEMA");
  if (!planilha) return { sucesso: false, mensagem: "Erro: Aba 'USUARIOS_SISTEMA' não existe." };
  var dados = planilha.getDataRange().getValues();
  for (var i = 1; i < dados.length; i++) {
    var usuarioBanco = String(dados[i][0]).trim();
    var senhaBanco = String(dados[i][1]).trim();
    if (usuarioBanco === usuario && senhaBanco === senha) {
      return {
        sucesso: true, nome: String(dados[i][2]).trim(), usuario: usuarioBanco,
        permissoes: String(dados[i][3]).trim().toLowerCase().split(",").map(function(item) { return item.trim(); }) 
      };
    }
  }
  return { sucesso: false, mensagem: "Usuário ou senha incorretos!" };
}
