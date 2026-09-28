// ================= CONTROLE DE NOTAS FISCAIS PAGAS =================
function salvarNotaFiscalForm(dados) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var aba = ss.getSheetByName("NOTAS_FISCAIS");
    
    if (!aba) {
      aba = ss.insertSheet("NOTAS_FISCAIS");
      aba.appendRow(["ID", "FORNECEDOR", "DATA", "VALOR", "COMPROVANTE", "NR", "NOTA FISCAL"]);
      aba.getRange("A1:G1").setFontWeight("bold").setBackground("#1e3a8a").setFontColor("#ffffff");
    }
    
    var pastas = DriveApp.getFoldersByName("NOTAS_COMPROVANTES");
    var pastaDestino = pastas.hasNext() ? pastas.next() : DriveApp.createFolder("NOTAS_COMPROVANTES");

    function salvarDoc(base64, mime, nomeSufixo) {
      if (!base64) return "";
      var decoded = Utilities.base64Decode(base64);
      var blob = Utilities.newBlob(decoded, mime, "NF_" + dados.notaFiscal + "_" + nomeSufixo + "_" + new Date().getTime());
      var arquivo = pastaDestino.createFile(blob);
      arquivo.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      return arquivo.getUrl();
    }

    var linkComprovante = salvarDoc(dados.fileBase64, dados.fileMime, "COMPROVANTE");
    var linkNFDigital = salvarDoc(dados.fileBase64_2, dados.fileMime_2, "DIGITAL");

    var dataFormatada = dados.data;
    if(dataFormatada && dataFormatada.includes("-")) {
      var p = dataFormatada.split("-");
      dataFormatada = p[2] + "/" + p[1] + "/" + p[0];
    } else {
      dataFormatada = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy");
    }

    var ultimaLinha = aba.getLastRow();
    var nextId = ultimaLinha > 0 ? parseInt(aba.getRange(ultimaLinha, 1).getValue()) + 1 : 20000;
    if(isNaN(nextId)) nextId = 20000;
    
    aba.appendRow([
      nextId,
      dados.fornecedor.toUpperCase(),
      dataFormatada,
      parseFloat(dados.valor || 0),
      linkComprovante,
      dados.notaFiscal.toUpperCase(), // Fica na Coluna F (NR)
      linkNFDigital // Fica na Coluna G (NOTA FISCAL)
    ]);
    
    return { sucesso: true, mensagem: "Nota e documentos registrados com sucesso!" };
  } catch (erro) { return { sucesso: false, mensagem: "Erro: " + erro.toString() }; }
}

function listarNotasFiscais() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var aba = ss.getSheetByName("NOTAS_FISCAIS");
  if (!aba) return [];
  
  var dados = aba.getDataRange().getDisplayValues();
  var lista = [];
  
  var limiteInicio = Math.max(1, dados.length - 200);
  
  for (var i = limiteInicio; i < dados.length; i++) {
    if (!dados[i][0] && !dados[i][1]) continue; 
    
    lista.push({
      id: dados[i][0], 
      fornecedor: dados[i][1], 
      data: dados[i][2], 
      valor: dados[i][3],
      comprovante: dados[i][4], // Coluna E
      notaFiscal: dados[i][5],  // Coluna F (NR)
      nfDigital: dados[i][6],   // Coluna G (NOTA FISCAL LINK)
      rowIndex: i + 1 // Linha exata do Excel
    });
  }
  return lista.reverse();
}

function excluirArquivoNotaFiscal(rowIndex, colunaNome) {
  try {
    var aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("NOTAS_FISCAIS");
    var col = (colunaNome === 'comprovante') ? 5 : 7; 
    aba.getRange(rowIndex, col).setValue(""); 
    return { sucesso: true, mensagem: "Arquivo removido do registro!" };
  } catch (e) { return { sucesso: false, mensagem: "Erro ao excluir: " + e.toString() }; }
}

function atualizarArquivoNotaFiscal(rowIndex, tipoAnexo, base64Data, mimeType, fileName) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var aba = ss.getSheetByName("NOTAS_FISCAIS");
    if (!aba) return { sucesso: false, message: "Aba NOTAS_FISCAIS não encontrada." };
    
    var col = (tipoAnexo === 'comprovante') ? 5 : 7; 
    
    var pastas = DriveApp.getFoldersByName("NOTAS_COMPROVANTES");
    var pastaDestino = pastas.hasNext() ? pastas.next() : DriveApp.createFolder("NOTAS_COMPROVANTES");
    
    var decoded = Utilities.base64Decode(base64Data);
    var numNF = aba.getRange(rowIndex, 6).getValue(); 
    var nomeSufixo = (tipoAnexo === 'comprovante') ? "COMPROVANTE" : "DIGITAL";
    var nomeArquivo = "NF_" + numNF + "_ATUALIZADO_" + nomeSufixo + "_" + new Date().getTime();
    
    var blob = Utilities.newBlob(decoded, mimeType, nomeArquivo);
    var arquivo = pastaDestino.createFile(blob);
    arquivo.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    var urlNova = arquivo.getUrl();
    aba.getRange(rowIndex, col).setValue(urlNova);
    
    return { sucesso: true, mensagem: "Documento anexado com sucesso!" };
  } catch (e) {
    return { sucesso: false, mensagem: "Erro ao atualizar arquivo: " + e.toString() };
  }
}

function atualizarNotaFiscalBackend(idNota, dados) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var aba = ss.getSheetByName("NOTAS_FISCAIS"); 
    var dadosPlanilha = aba.getDataRange().getValues();
    
    var linhaCorreta = -1;
    for (var i = 1; i < dadosPlanilha.length; i++) {
      if (String(dadosPlanilha[i][0]).trim() === String(idNota).trim()) {
        linhaCorreta = i + 1; 
        break;
      }
    }
    
    if (linhaCorreta === -1) {
      return { sucesso: false, mensagem: "❌ Erro: Nota Fiscal não encontrada na planilha." };
    }
    
    // 1. Atualiza os textos normais
    var dataGravar = dados.data;
    if (dataGravar && dataGravar.includes("-")) {
      var p = dataGravar.split("-");
      dataGravar = p[2] + "/" + p[1] + "/" + p[0];
    }
    
    aba.getRange(linhaCorreta, 2).setValue(String(dados.fornecedor).toUpperCase()); 
    aba.getRange(linhaCorreta, 3).setValue(dataGravar);                            
    aba.getRange(linhaCorreta, 4).setValue(parseFloat(dados.valor) || 0);          
    aba.getRange(linhaCorreta, 6).setValue(dados.numeroNota);                      
    
    // ====================================================================
    // 2. A MÁGICA DOS ANEXOS NA EDIÇÃO (O QUE FALTAVA)
    // ====================================================================
    var pastas = DriveApp.getFoldersByName("NOTAS_COMPROVANTES");
    var pastaDestino = pastas.hasNext() ? pastas.next() : DriveApp.createFolder("NOTAS_COMPROVANTES");

    function salvarDoc(base64, mime, nomeSufixo) {
      if (!base64) return "";
      var decoded = Utilities.base64Decode(base64);
      var blob = Utilities.newBlob(decoded, mime, "NF_" + dados.numeroNota + "_" + nomeSufixo + "_" + new Date().getTime());
      var arquivo = pastaDestino.createFile(blob);
      arquivo.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      return arquivo.getUrl();
    }

    // Se a tela enviou um Comprovante novo na janela de edição, guarda-o na Coluna E (5)
    if (dados.fileBase64 && dados.fileBase64 !== "") {
      var linkComprovante = salvarDoc(dados.fileBase64, dados.fileMime, "COMPROVANTE_ADD");
      aba.getRange(linhaCorreta, 5).setValue(linkComprovante);
    }

    // Se a tela enviou uma NF Digital nova na janela de edição, guarda-a na Coluna G (7)
    if (dados.fileBase64_2 && dados.fileBase64_2 !== "") {
      var linkNFDigital = salvarDoc(dados.fileBase64_2, dados.fileMime_2, "DIGITAL_ADD");
      aba.getRange(linhaCorreta, 7).setValue(linkNFDigital);
    }
    
    return { sucesso: true, mensagem: "✅ Dados e Anexos atualizados com sucesso!" };
  } catch(e) { 
    return { sucesso: false, mensagem: "Erro ao atualizar: " + e.toString() }; 
  }
}
function excluirNotaFiscalBackend(linha) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var aba = ss.getSheetByName("NOTAS_FISCAIS"); 
    aba.deleteRow(linha);
    return { sucesso: true, mensagem: "✅ Nota Fiscal excluída com sucesso!" };
  } catch(e) { 
    return { sucesso: false, mensagem: "Erro ao excluir: " + e.toString() }; 
  }
}
