// ================= MÓDULO SOS: NOVA ABA SOS_FROTA =================

function salvarNovoSOS(dados) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var abaSos = ss.getSheetByName("SOS_FROTA");
    
    // Se a aba não existir, o sistema cria automaticamente com os cabeçalhos certos!
    if (!abaSos) {
      abaSos = ss.insertSheet("SOS_FROTA");
      abaSos.appendRow(["DATA E HORA DA SOLICITAÇÃO", "PREFIXO", "PLACA", "TIPO DE SOLICITAÇÃO", "DESCRIÇÃO", "STATUS", "SOLICITANTE", "DATA E HORA DA CONCLUSÃO", "SOLICITAÇÃO DE REBOQUE", "TEMPO TOTAL"]);
      abaSos.getRange("A1:J1").setFontWeight("bold").setBackground("#166534").setFontColor("#ffffff");
    }

    // NOVO (revisão 8): usa a data/hora de abertura escolhida pelo usuário no formulário
    // (dados.dataAbertura, no formato "yyyy-MM-ddTHH:mm"), com fallback para o momento atual
    // caso o campo não tenha sido enviado.
    var dataHoraAtual = new Date();
    if (dados.dataAbertura) {
      var dataConvertida = new Date(dados.dataAbertura);
      if (!isNaN(dataConvertida.getTime())) dataHoraAtual = dataConvertida;
    }
    
    // Insere as 10 colunas exatamente na ordem correta
    abaSos.appendRow([
      dataHoraAtual, dados.prefixo, dados.placa, dados.tipo, dados.descricao,
      "EM ABERTO", dados.solicitante, "", dados.reboque, ""
    ]);

    // Atualiza a Manutenção
    var abaManutencao = ss.getSheetByName("MANUTENÇÃO");
    if (abaManutencao) {
      var dadosMan = abaManutencao.getDataRange().getValues();
      for (var i = 1; i < dadosMan.length; i++) {
        if (String(dadosMan[i][1]).trim().toUpperCase() === dados.prefixo.trim().toUpperCase()) {
          abaManutencao.getRange(i + 1, 4).setValue("SOS");
          break;
        }
      }
    }
    return { sucesso: true, mensagem: "🚨 SOS Registrado! Veículo bloqueado na Manutenção." };
  } catch (e) {
    return { sucesso: false, mensagem: "Erro: " + e.toString() };
  }
}

function listarSOS() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var aba = ss.getSheetByName("SOS_FROTA"); 
    if (!aba) return [];
    
    var ultimaLinha = aba.getLastRow();
    if (ultimaLinha <= 1) return [];

    var dados = aba.getRange(2, 1, ultimaLinha - 1, 10).getValues();
    var lista = [];
    var limiteInicio = Math.max(0, dados.length - 150);

    for (var i = limiteInicio; i < dados.length; i++) {
      var row = dados[i];
      if (!row[0] && !row[1]) continue;

      var dataStr = row[0];
      if (dataStr instanceof Date) {
        dataStr = ("0"+dataStr.getDate()).slice(-2)+"/"+("0"+(dataStr.getMonth()+1)).slice(-2)+"/"+dataStr.getFullYear()+" "+("0"+dataStr.getHours()).slice(-2)+":"+("0"+dataStr.getMinutes()).slice(-2);
      } else { dataStr = String(row[0] || ""); }

      var dataFim = row[7];
      if (dataFim instanceof Date) {
        dataFim = ("0"+dataFim.getDate()).slice(-2)+"/"+("0"+(dataFim.getMonth()+1)).slice(-2)+"/"+dataFim.getFullYear()+" "+("0"+dataFim.getHours()).slice(-2)+":"+("0"+dataFim.getMinutes()).slice(-2);
      } else { dataFim = String(row[7] || ""); }

      lista.push({
        linhaOrigem: i + 2,
        dataSolicitacao: dataStr,
        prefixo: String(row[1] || ""),
        placa: String(row[2] || ""),
        tipo: String(row[3] || ""),
        descricao: String(row[4] || ""),
        status: String(row[5] || "EM ABERTO"),
        solicitante: String(row[6] || ""),
        dataConclusao: dataFim,
        reboque: String(row[8] || ""),
        tempoTotal: String(row[9] || "")
      });
    }
    return lista.reverse();
  } catch (e) { return []; }
}

// NOVA FUNÇÃO (revisão 9): corrige dados de um SOS já lançado (prefixo/placa errados, erro de
// digitação em qualquer campo), sem apagar o registo e sem alterar o status do chamado. Só
// sobrescreve os campos que vierem preenchidos em "dados" — o mesmo cuidado já usado nas
// outras rotinas revisadas, para nunca apagar um valor existente com um campo vazio.
function atualizarSOSBackend(linha, dados) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var abaSos = ss.getSheetByName("SOS_FROTA");
    if (!abaSos) return { sucesso: false, mensagem: "Aba SOS_FROTA não encontrada." };

    linha = parseInt(linha, 10);
    if (!linha || linha < 2) return { sucesso: false, mensagem: "Linha do SOS inválida." };

    var prefixoAntigo = String(abaSos.getRange(linha, 2).getValue() || "").trim().toUpperCase();
    var statusAtual = String(abaSos.getRange(linha, 6).getValue() || "");

    if (dados.dataAbertura) {
      var dataConvertida = new Date(dados.dataAbertura);
      if (!isNaN(dataConvertida.getTime())) abaSos.getRange(linha, 1).setValue(dataConvertida);
    }
    if (dados.prefixo) abaSos.getRange(linha, 2).setValue(dados.prefixo);
    if (dados.placa) abaSos.getRange(linha, 3).setValue(dados.placa);
    if (dados.tipo) abaSos.getRange(linha, 4).setValue(dados.tipo);
    if (dados.descricao) abaSos.getRange(linha, 5).setValue(dados.descricao);
    if (dados.solicitante) abaSos.getRange(linha, 7).setValue(dados.solicitante);
    if (dados.reboque) abaSos.getRange(linha, 9).setValue(dados.reboque);

    // Se o prefixo foi corrigido e o chamado ainda está EM ABERTO, ajusta o bloqueio na aba
    // MANUTENÇÃO: libera o veículo errado (que tinha ficado travado por engano) e bloqueia o
    // veículo correto, em vez de deixar o prefixo antigo travado indefinidamente.
    var prefixoNovo = dados.prefixo ? String(dados.prefixo).trim().toUpperCase() : prefixoAntigo;
    if (statusAtual === "EM ABERTO" && dados.prefixo && prefixoNovo !== prefixoAntigo) {
      var abaManutencao = ss.getSheetByName("MANUTENÇÃO");
      if (abaManutencao) {
        var dadosMan = abaManutencao.getDataRange().getValues();
        for (var i = 1; i < dadosMan.length; i++) {
          var pref = String(dadosMan[i][1]).trim().toUpperCase();
          if (pref === prefixoAntigo) abaManutencao.getRange(i + 1, 4).setValue("LIBERADO");
          if (pref === prefixoNovo) abaManutencao.getRange(i + 1, 4).setValue("SOS");
        }
      }
    }

    return { sucesso: true, mensagem: "✅ SOS atualizado com sucesso!" };
  } catch (e) {
    return { sucesso: false, mensagem: "Erro: " + e.toString() };
  }
}

function fecharChamadoSOSBackend(linhaSOS, prefixo, dataFechamentoEscolhida) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var abaSos = ss.getSheetByName("SOS_FROTA"); 
    if (!abaSos) return { sucesso: false, mensagem: "Aba SOS_FROTA não encontrada." };

    // NOVO (revisão 8): aceita a data/hora de conclusão escolhida pelo usuário
    // (dataFechamentoEscolhida, formato "yyyy-MM-ddTHH:mm"), com fallback para o momento
    // atual caso não seja enviada (mantém retrocompatibilidade com chamadas antigas).
    var dataFechamento = new Date();
    if (dataFechamentoEscolhida) {
      var dataFechamentoConvertida = new Date(dataFechamentoEscolhida);
      if (!isNaN(dataFechamentoConvertida.getTime())) dataFechamento = dataFechamentoConvertida;
    }
    var dataFimFormatada = Utilities.formatDate(dataFechamento, ss.getSpreadsheetTimeZone(), "dd/MM/yyyy HH:mm:ss");
    
    var dataAbertura = abaSos.getRange(linhaSOS, 1).getValue();

    abaSos.getRange(linhaSOS, 6).setValue("FECHADO");
    abaSos.getRange(linhaSOS, 8).setValue("'" + dataFimFormatada); 

    if (dataAbertura instanceof Date) {
      var diffMs = dataFechamento.getTime() - dataAbertura.getTime();
      var diffHrs = Math.floor(diffMs / 3600000);
      var diffMins = Math.floor((diffMs % 3600000) / 60000);
      abaSos.getRange(linhaSOS, 10).setValue(diffHrs + "h " + diffMins + "m");
    }

    var abaManutencao = ss.getSheetByName("MANUTENÇÃO");
    if (abaManutencao && prefixo) {
      var dadosMan = abaManutencao.getDataRange().getValues();
      for (var i = 1; i < dadosMan.length; i++) {
        if (String(dadosMan[i][1]).trim().toUpperCase() === String(prefixo).trim().toUpperCase()) {
          abaManutencao.getRange(i + 1, 4).setValue("LIBERADO");
          break;
        }
      }
    }
    return { sucesso: true, message: "✅ SOS Concluído! Data e Hora registradas com sucesso." };
  } catch (e) { return { sucesso: false, mensagem: "Erro: " + e.toString() }; }
}
