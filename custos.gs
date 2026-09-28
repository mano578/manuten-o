// ================= MÓDULO BACKEND: CONTROLE DE CUSTOS E RATEIO DE NOTAS =================

function salvarCustoForm(dados) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var aba = ss.getSheetByName("CUSTOS_FROTA") || ss.getSheetByName("CUSTOS_FROTA ");
    
    if (!aba) {
      aba = ss.insertSheet("CUSTOS_FROTA");
      aba.appendRow(["ID", "PREFIXO", "PLACA", "FORNECEDOR", "DESCRIÇÃO", "DATA", "TIPO DE CUSTO", "CUSTO", "NOTA FISCAL"]);
      aba.getRange("A1:I1").setFontWeight("bold").setBackground("#78350f").setFontColor("#ffffff");
    }
    
    var veiculos = dados.veiculosSelecionados;
    if (!veiculos || veiculos.length === 0) {
      return { sucesso: false, mensagem: "Erro: Nenhum veículo selecionado para o rateio." };
    }
    
    var dataFormatada = dados.data;
    if (dataFormatada && dataFormatada.includes("-")) {
      var partes = dataFormatada.split("-");
      dataFormatada = partes[2] + "/" + partes[1] + "/" + partes[0]; 
    } else if (!dataFormatada) {
      dataFormatada = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), "dd/MM/yyyy");
    }
    
    var ultimaLinha = aba.getLastRow();
    var maxId = 0;
    
    // O SEGREDO DO ID ÚNICO: Varre a planilha inteira para achar o maior número absoluto
    if (ultimaLinha > 1) {
      var valoresId = aba.getRange(2, 1, ultimaLinha - 1, 1).getValues();
      for (var k = 0; k < valoresId.length; k++) {
        var idNum = parseInt(valoresId[k][0], 10);
        if (!isNaN(idNum) && idNum > maxId) {
          maxId = idNum;
        }
      }
    }
    
    var nextId = maxId > 0 ? maxId + 1 : 1;
    var linhasParaInserir = [];
    
    for (var i = 0; i < veiculos.length; i++) {
      var v = veiculos[i];
      for (var j = 0; j < v.itens.length; j++) {
        var item = v.itens[j];
        
        linhasParaInserir.push([
          nextId++,                          // Incrementa sempre a partir do maior ID encontrado
          String(v.prefixo).toUpperCase(),
          String(v.placa).toUpperCase(),
          dados.fornecedor.toUpperCase(),
          item.servico.toUpperCase(),
          dataFormatada,
          item.tipo.toUpperCase(),
          parseFloat(item.valor) || 0,
          String(dados.notaFiscal).toUpperCase()
        ]);
      }
    }
    
    if (linhasParaInserir.length > 0) {
      aba.getRange(aba.getLastRow() + 1, 1, linhasParaInserir.length, 9).setValues(linhasParaInserir);
    }
    
    return { sucesso: true, mensagem: "✅ Nota Fiscal registrada com sucesso!" };
  } catch (erro) {
    return { sucesso: false, mensagem: "Erro ao salvar custos da frota: " + erro.toString() };
  }
}

function listarCustosBackend() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var aba = ss.getSheetByName("CUSTOS_FROTA") || ss.getSheetByName("CUSTOS_FROTA ");
    if (!aba) return { sucesso: false, mensagem: "Aba 'CUSTOS_FROTA' não encontrada!" };
    
    var dados = aba.getDataRange().getValues();
    var lista = [];
    
    for (var i = 1; i < dados.length; i++) {
      if (dados[i][0] !== "") { 
        
        var dataFormatada = dados[i][5]; // Coluna F (5)
        if (dataFormatada instanceof Date) {
          dataFormatada = Utilities.formatDate(dataFormatada, ss.getSpreadsheetTimeZone(), "dd/MM/yyyy");
        } else if (!dataFormatada) {
          dataFormatada = "-";
        }

        var custoBruto = dados[i][7]; // Coluna H (7)
        var custoNum = 0;
        if (typeof custoBruto === 'number') { custoNum = custoBruto; } 
        else if (typeof custoBruto === 'string') { custoNum = parseFloat(custoBruto.replace(/\./g, '').replace(',', '.')) || 0; }
        
        lista.push({
          rowIndex: i + 1,
          id: String(dados[i][0] || "").trim().toUpperCase(),       // Coluna A (0)
          prefixo: String(dados[i][1] || "").trim().toUpperCase(),  // Coluna B (1)
          placa: String(dados[i][2] || "").trim().toUpperCase(),    // Coluna C (2)
          fornecedor: String(dados[i][3] || "").trim(),             // Coluna D (3)
          descricao: String(dados[i][4] || "").trim().toUpperCase(),// Coluna E (4) -> DESCRIÇÃO
          data: dataFormatada,                                      // Coluna F (5) -> DATA
          tipo: String(dados[i][6] || "").trim().toUpperCase(),     // Coluna G (6) -> TIPO DE CUSTO
          valor: custoNum,                                          // Coluna H (7) -> CUSTO
          notaFiscal: String(dados[i][8] || "").trim()              // Coluna I (8) -> NOTA FISCAL
        });
      }
    }
    return { sucesso: true, dados: lista.reverse() }; 
  } catch(e) { return { sucesso: false, mensagem: "Erro no servidor: " + e.toString() }; }
}

function atualizarCustoBackend(linha, dados) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var aba = ss.getSheetByName("CUSTOS_FROTA") || ss.getSheetByName("CUSTOS_FROTA ");

    // Conversão da data para gravação organizada no Sheets (dd/MM/yyyy)
    var dataGravar = dados.data;
    if (dataGravar && dataGravar.includes("-")) {
      var p = dataGravar.split("-");
      dataGravar = p[2] + "/" + p[1] + "/" + p[0];
    }

    // 🔒 O ID (Coluna A) é o identificador único do registro e NUNCA é alterado nesta função —
    // permite corrigir todos os demais dados lançados sem risco de duplicar/perder o ID.
    aba.getRange(linha, 2).setValue(String(dados.prefixo || "").toUpperCase());         // Coluna B: PREFIXO
    aba.getRange(linha, 3).setValue(String(dados.placa || "").toUpperCase());           // Coluna C: PLACA
    aba.getRange(linha, 4).setValue(String(dados.fornecedor || "").toUpperCase());      // Coluna D: FORNECEDOR
    aba.getRange(linha, 5).setValue(String(dados.descricao || "").toUpperCase());       // Coluna E: DESCRIÇÃO
    aba.getRange(linha, 6).setValue(dataGravar);                                        // Coluna F: DATA
    aba.getRange(linha, 7).setValue(String(dados.tipo || "").toUpperCase());            // Coluna G: TIPO DE CUSTO
    aba.getRange(linha, 8).setValue(parseFloat(dados.valor) || 0);                      // Coluna H: CUSTO
    aba.getRange(linha, 9).setValue(String(dados.notaFiscal || "").toUpperCase());      // Coluna I: NOTA FISCAL

    return { sucesso: true, mensagem: "✅ Registro de Custo atualizado com sucesso!" };
  } catch(e) { return { sucesso: false, mensagem: "Erro ao atualizar: " + e.toString() }; }
}

function excluirCustoBackend(linha) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var aba = ss.getSheetByName("CUSTOS_FROTA") || ss.getSheetByName("CUSTOS_FROTA ");
    aba.deleteRow(linha);
    return { sucesso: true, mensagem: "🗑️ Registro de custo excluído permanentemente da planilha!" };
  } catch(e) { return { sucesso: false, mensagem: "Erro ao excluir: " + e.toString() }; }
}

// Mantido por compatibilidade de chamadas antigas se houver
function listarCustos() {
  var res = listarCustosBackend();
  return res.sucesso ? res.dados : [];
}
