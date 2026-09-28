// =========================================================================
// MÓDULO DRIVER DO SISTEMA (UPLOAD, LISTAGEM E EXCLUSÃO)
// =========================================================================

// 🔥 ID DA SUA PASTA APLICADO COM SUCESSO 🔥
var ID_PASTA_DRIVER = "1zCS1ww0ayiLU3b3_CJDQdGCn5Fi232WO"; 

function salvarArquivoDriver(obj) {
  try {
    var pasta = DriveApp.getFolderById(ID_PASTA_DRIVER);
    // Transforma o texto Base64 de volta num arquivo físico
    var blob = Utilities.newBlob(Utilities.base64Decode(obj.dados), obj.mimeType, obj.nome);
    pasta.createFile(blob);
    
    return { sucesso: true, mensagem: "Arquivo salvo na nuvem com sucesso!" };
  } catch (e) {
    return { sucesso: false, mensagem: "Erro no servidor: " + e.message };
  }
}

// NOTA (revisão - BUG REAL CORRIGIDO): este arquivo tinha sua própria
// listarArquivosDriver(), sem parâmetro, sem checagem de admin, lendo da pasta fixa
// ID_PASTA_DRIVER — um ID diferente da pasta "SISTEMA_INOVAR_DRIVER" onde
// salvarArquivoBase64 (em Admin.gs) realmente grava os uploads do painel Admin.
// Como este arquivo é carregado depois de Admin.gs, esta era a definição que
// VENCIA em runtime (a última função com o mesmo nome sobrescreve as anteriores em
// JS/Apps Script) — então a lista de arquivos do Drive exibida no painel Admin nunca
// mostrava os arquivos recém-enviados, pois liam pastas diferentes. Esta definição
// duplicada foi removida; a função ativa agora é a de Admin.gs, que lê a pasta
// correta e confere permissão de administrador. salvarArquivoDriver (acima) e
// apagarArquivoDriverServidor (abaixo) não têm esse conflito de nome e permanecem
// intactas.

function apagarArquivoDriverServidor(id) {
  try {
    // Move o arquivo com o ID correspondente para a lixeira do Google Drive
    DriveApp.getFileById(id).setTrashed(true);
    return true;
  } catch(e) { 
    return false; 
  }
}
