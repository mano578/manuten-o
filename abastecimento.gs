// ================= CONTROLE DE ABASTECIMENTO =================

function salvarAbastecimentoForm(form) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var aba = ss.getSheetByName("ABASTECIMENTO_FROTA");
    
    // Se a aba não existir, cria exatamente com o padrão da sua base de dados (14 Colunas)
    if (!aba) {
      aba = ss.insertSheet("ABASTECIMENTO_FROTA");
      aba.appendRow(["ID", "VEICULO", "PLACA", "DATA", "COMBUSTÍVEL", "# LITROS", "VALOR LITRO", "VALOR TOTAL", "KM", "# KM TOTAL", "# KM / LITRO", "# LITRO / KM", "R$ / KM", "CENTRO DE CUSTO"]);
      aba.getRange("A1:N1").setFontWeight("bold").setBackground("#38761d").setFontColor("#ffffff");
    }
    
    var dataComp = form.data ? formatarDataInputHtml(form.data).split(" ")[0] : Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy");
    
    var ultimaLinha = aba.getLastRow();
    var nextId = ultimaLinha > 0 ? (ultimaLinha * 15) + 11000 : 11000;
    
    aba.appendRow([
      nextId,
      form.prefixo,
      form.placa,
      dataComp,
      form.combustivel.toUpperCase(),
      parseFloat(form.litros || 0),
      parseFloat(form.valorLitro || 0),
      parseFloat(form.valorTotal || 0),
      parseFloat(form.km || 0),
      parseFloat(form.kmTotal || 0),
      parseFloat(form.kmPorLitro || 0),
      parseFloat(form.litroPorKm || 0),
      parseFloat(form.rsPorKm || 0),
      (form.centroCusto || "A DEFINIR").toUpperCase()
    ]);
    
    return { sucesso: true, mensagem: "Abastecimento manual registrado com sucesso!" };
  } catch (erro) {
    return { sucesso: false, mensagem: "Erro ao registrar abastecimento: " + erro.toString() };
  }
}

function salvarAbastecimentoLote(lote) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var aba = ss.getSheetByName("ABASTECIMENTO_FROTA") || ss.getSheetByName("ABASTECIMENTO");
    
    if (!aba) {
      return { sucesso: false, mensagem: "Aba não encontrada!" };
    }

    var ultimaLinha = aba.getLastRow();
    var maxId = 0;
    if (ultimaLinha > 1) {
      var valoresId = aba.getRange(2, 1, ultimaLinha - 1, 1).getValues();
      for (var k = 0; k < valoresId.length; k++) {
        var idNum = parseInt(valoresId[k][0], 10);
        if (!isNaN(idNum) && idNum > maxId) maxId = idNum;
      }
    }
    var nextId = maxId > 0 ? maxId + 1 : 1;

    var linhasParaInserir = [];

    // O ENCAIXE PERFEITO NAS 14 COLUNAS DA SUA PLANILHA BASE
    for (var i = 0; i < lote.length; i++) {
      var item = lote[i];

      linhasParaInserir.push([
        nextId++,                                      // 1. Coluna A: ID
        String(item.veiculo || "").toUpperCase(),      // 2. Coluna B: VEÍCULO
        String(item.placa || "").toUpperCase(),        // 3. Coluna C: PLACA (Injetada)
        item.data,                                     // 4. Coluna D: DATA
        String(item.combustivel || "").toUpperCase(),  // 5. Coluna E: COMBUSTÍVEL
        parseFloat(item.litros) || 0,                  // 6. Coluna F: # LITROS
        parseFloat(item.valorLitro) || 0,              // 7. Coluna G: VALOR LITRO
        parseFloat(item.valorTotal) || 0,              // 8. Coluna H: VALOR TOTAL
        parseFloat(item.km) || 0,                      // 9. Coluna I: KM
        parseFloat(item.kmTotal) || 0,                 // 10. Coluna J: # KM TOTAL
        parseFloat(item.kmPorLitro) || 0,              // 11. Coluna K: # KM / LITRO
        parseFloat(item.litroPorKm) || 0,              // 12. Coluna L: # LITRO / KM
        parseFloat(item.rsPorKm) || 0,                 // 13. Coluna M: R$ / KM
        String(item.centroCusto || "").toUpperCase()   // 14. Coluna N: CENTRO DE CUSTO
      ]);
    }

    if (linhasParaInserir.length > 0) {
      aba.getRange(aba.getLastRow() + 1, 1, linhasParaInserir.length, 14).setValues(linhasParaInserir);
    }

    return { sucesso: true, mensagem: "✅ " + linhasParaInserir.length + " abastecimentos importados com sucesso!" };
  } catch (e) {
    return { sucesso: false, mensagem: "Erro ao salvar lote: " + e.toString() };
  }
}

// 1. LISTAR (Base original mantida com correção para ignorar linhas vazias no cálculo do limite)
function listarHistoricoAbastecimentosBackend() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var aba = ss.getSheetByName("ABASTECIMENTO_FROTA") || ss.getSheetByName("ABASTECIMENTO");
    if (!aba) return { sucesso: false, mensagem: "Aba de Abastecimento não encontrada." };

    var dados = aba.getDataRange().getDisplayValues(); 
    var lista = [];
    
    // 1. Proteção: Descobre a última linha REAL com dados (ignorando linhas em branco no fim)
    var ultimaLinhaComDados = 0;
    for (var k = dados.length - 1; k >= 0; k--) {
      if (dados[k][0] !== "") {
        ultimaLinhaComDados = k;
        break;
      }
    }
    if (ultimaLinhaComDados === 0) ultimaLinhaComDados = dados.length - 1;

    // 2. Calcula o limite puxando os últimos 150 registros
    var limiteInicio = Math.max(1, ultimaLinhaComDados - 150);

    for (var i = limiteInicio; i <= ultimaLinhaComDados; i++) {
      if (dados[i][0] !== "") {
        lista.push({
          rowIndex: i + 1, 
          id: dados[i][0],
          prefixo: String(dados[i][1]).trim().toUpperCase(),
          placa: String(dados[i][2]).trim().toUpperCase(),
          data: String(dados[i][3]).trim(),
          combustivel: String(dados[i][4]).trim().toUpperCase(),
          litros: String(dados[i][5]).trim(),
          valorLitro: String(dados[i][6]).trim(),
          valorTotal: String(dados[i][7]).trim(),
          centroCusto: String(dados[i][13] || "").trim().toUpperCase()
        });
      }
    }
    
    // 🔥 3. A MÁGICA FINAL: Força a ordenação do Maior ID (Mais Novo) para o Menor ID
    lista.sort(function(a, b) { 
      return parseInt(b.id) - parseInt(a.id); 
    });
    
    // Retorna a lista perfeitamente ordenada para o Frontend
    return { sucesso: true, dados: lista }; 
  } catch(e) { 
    return { sucesso: false, mensagem: "Erro no servidor: " + e.toString() }; 
  }
}

// 2. EXCLUIR
function excluirAbastecimentoBackend(linha) {
  try {
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("ABASTECIMENTO_FROTA").deleteRow(linha);
    return { sucesso: true, mensagem: "🗑️ Abastecimento excluído com sucesso!" };
  } catch(e) { return { sucesso: false, mensagem: "Erro: " + e.toString() }; }
}

// 3. EDITAR (Ajustado para atualizar apenas os dados que o modal de edição envia)
function atualizarAbastecimentoBackend(linha, dados) {
  try {
    var aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("ABASTECIMENTO_FROTA");
    
    // Atualiza APENAS as colunas que o modal envia (Litros, Valor Total e Centro de Custo)
    aba.getRange(linha, 6).setValue(parseFloat(dados.litros) || 0); // Coluna F
    aba.getRange(linha, 8).setValue(parseFloat(dados.valorTotal) || 0); // Coluna H
    aba.getRange(linha, 14).setValue(String(dados.centroCusto).toUpperCase()); // Coluna N

    return { sucesso: true, mensagem: "✅ Abastecimento atualizado!" };
  } catch(e) { return { sucesso: false, mensagem: "Erro: " + e.toString() }; }
}
