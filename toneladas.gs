// ================= CONTROLE DE TONELADAS DA FROTA =================

function salvarToneladasLote(loteDados) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var aba = ss.getSheetByName("TONELADA_FROTA");
    
    // Cria a aba se não existir
    if (!aba) {
      aba = ss.insertSheet("TONELADA_FROTA");
      aba.appendRow(["ID", "DATA", "DIA", "MOTORISTA", "PLACA", "VEÍCULO", "SETOR", "TICKET", "LÍQUIDO", "TURNO", "TIPO DO MATERIAL", "OBSERVAÇÃO"]);
      aba.getRange("A1:L1").setFontWeight("bold").setBackground("#166534").setFontColor("#ffffff");
    }
    
    var linhasParaInserir = [];
    var registrosValidos = 0;
    
    for (var i = 0; i < loteDados.length; i++) {
      var item = loteDados[i];
      
      // Proteção contra duplicação no banco: não insere se já foi marcado como duplicado
      if(item.duplicado) continue; 
      
      linhasParaInserir.push([
        item.id,
        item.data,
        item.dia,
        item.motorista,
        item.placa,
        item.veiculo,
        item.setor,
        item.ticket,
        item.liquido,
        item.turno,
        item.tipoMaterial,
        item.observacao
      ]);
      registrosValidos++;
    }
    
    if (linhasParaInserir.length > 0) {
      aba.getRange(aba.getLastRow() + 1, 1, linhasParaInserir.length, 12).setValues(linhasParaInserir);
    }
    
    return { sucesso: true, mensagem: registrosValidos + " registros de pesagem inseridos com sucesso na base!" };
  } catch (erro) {
    return { sucesso: false, mensagem: "Erro no processamento em lote: " + erro.toString() };
  }
}

function listarToneladas() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var aba = ss.getSheetByName("TONELADA_FROTA");
  if (!aba) return [];
  
  var dados = aba.getDataRange().getDisplayValues();
  var lista = [];
  
  // OTIMIZAÇÃO: Lê apenas as últimas 200 linhas
  var limiteInicio = Math.max(1, dados.length - 200);
  
  for (var i = limiteInicio; i < dados.length; i++) {
    if (!dados[i][0] && !dados[i][1]) continue; // Filtro Anti-Fantasma
    
    lista.push({
      id: dados[i][0],
      data: dados[i][1],
      dia: dados[i][2],
      motorista: dados[i][3],
      placa: dados[i][4],
      veiculo: dados[i][5],
      setor: dados[i][6],
      ticket: dados[i][7],
      liquido: dados[i][8],
      turno: dados[i][9],
      tipoMaterial: dados[i][10],
      observacao: dados[i][11]
    });
  }
  return lista.reverse(); 
}
