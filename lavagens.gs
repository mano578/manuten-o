// ==============================================================
// MÓDULO LAVAGEM: ALINHADO COM AS 7 COLUNAS EXATAS DA PLANILHA
// Colunas: A(DATA), B(PREFIXO), C(PLACA), D(SERVIÇO), E(LAVAGEM), F(CUSTO), G(FUNCIONÁRIO)
// ==============================================================

function salvarLavagemForm(dados) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var aba = ss.getSheetByName("LAVAGEM_FROTA") || ss.getSheetByName("LAVAGEM");
    
    if (!aba) {
      return { sucesso: false, mensagem: "Aba LAVAGEM não encontrada. Verifique o nome da aba no Excel." };
    }

    // Formata a data e hora perfeitamente
    var dataFormatada = dados.data;
    if (dataFormatada && dataFormatada.includes("T")) {
      var p = dataFormatada.split("T");
      var d = p[0].split("-");
      dataFormatada = d[2] + "/" + d[1] + "/" + d[0] + " " + p[1];
    }

    var custoFormatado = parseFloat(dados.valorLavagem) || 0;

    // GRAVAÇÃO EXATA NAS 7 COLUNAS (Sem coluna de ID)
    aba.appendRow([
      dataFormatada,                            // Coluna A: DATA
      String(dados.prefixo).toUpperCase(),      // Coluna B: PREFIXO
      String(dados.placa).toUpperCase(),        // Coluna C: PLACA
      String(dados.tipoLavagem).toUpperCase(),  // Coluna D: SERVIÇO
      String(dados.localLavagem).toUpperCase(), // Coluna E: LAVAGEM (LOCAL)
      custoFormatado,                           // Coluna F: CUSTO
      String(dados.responsavel).toUpperCase()   // Coluna G: FUNCIONÁRIO
    ]);

    return { sucesso: true, mensagem: "✅ Lavagem registada com sucesso!" };
  } catch(e) {
    return { sucesso: false, mensagem: "Erro interno no servidor: " + e.toString() };
  }
}

function listarHistoricoLavagensBackend() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var aba = ss.getSheetByName("LAVAGEM_FROTA") || ss.getSheetByName("LAVAGEM");
    if (!aba) return { sucesso: false, mensagem: "Aba LAVAGEM não encontrada." };

    var dados = aba.getDataRange().getValues();
    var lista = [];

    for (var i = 1; i < dados.length; i++) {
      if (dados[i][0] !== "" && dados[i][1] !== "") { 
        
        var dataFormatada = dados[i][0]; // Coluna A (0)
        if (dataFormatada instanceof Date) {
          dataFormatada = Utilities.formatDate(dataFormatada, ss.getSpreadsheetTimeZone(), "dd/MM/yyyy HH:mm");
        }

        var custoBruto = dados[i][5]; // Coluna F (5)
        var custoNum = parseFloat(custoBruto) || 0;

        lista.push({
          rowIndex: i + 1, // Usamos o número da linha do Excel como "ID" invisível para apagar/editar
          data: dataFormatada,                                         // Coluna A
          prefixo: String(dados[i][1] || "").trim().toUpperCase(),     // Coluna B
          placa: String(dados[i][2] || "").trim().toUpperCase(),       // Coluna C
          servico: String(dados[i][3] || "").trim().toUpperCase(),     // Coluna D
          local: String(dados[i][4] || "").trim().toUpperCase(),       // Coluna E
          custo: custoNum,                                             // Coluna F
          funcionario: String(dados[i][6] || "").trim().toUpperCase()  // Coluna G
        });
      }
    }
    return { sucesso: true, dados: lista };
  } catch (e) {
    return { sucesso: false, mensagem: "Erro no servidor: " + e.toString() };
  }
}

function atualizarLavagemBackend(linha, dados) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var aba = ss.getSheetByName("LAVAGEM_FROTA") || ss.getSheetByName("LAVAGEM");
    
    // Atualiza apenas as colunas permitidas no modal de edição, respeitando a sua estrutura:
    aba.getRange(linha, 4).setValue(String(dados.servico).toUpperCase());     // Coluna D: SERVIÇO
    aba.getRange(linha, 5).setValue(String(dados.local).toUpperCase());       // Coluna E: LAVAGEM (LOCAL)
    aba.getRange(linha, 6).setValue(parseFloat(dados.custo) || 0);            // Coluna F: CUSTO
    aba.getRange(linha, 7).setValue(String(dados.funcionario).toUpperCase()); // Coluna G: FUNCIONÁRIO
    
    return { sucesso: true, mensagem: "✅ Serviço de Lavagem atualizado!" };
  } catch(e) {
    return { sucesso: false, mensagem: "Erro ao atualizar: " + e.toString() };
  }
}

function excluirLavagemBackend(linha) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var aba = ss.getSheetByName("LAVAGEM_FROTA") || ss.getSheetByName("LAVAGEM");
    aba.deleteRow(linha);
    return { sucesso: true, mensagem: "🗑️ Registo de lavagem excluído com sucesso!" };
  } catch(e) { 
    return { sucesso: false, mensagem: "Erro ao excluir: " + e.toString() }; 
  }
}
