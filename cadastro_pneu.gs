// ================= MÓDULO DE GESTÃO E CADASTRO DE PNEUS =================

function salvarNovoPneu(dados) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var abaPneu = ss.getSheetByName("CADASTRO_PNEU_FROTA");
    
    var cabecalhos = [
      "ID PNEU", "DATA/HORA CADASTRO", "NÚMERO DE FOGO", "MARCA/FABRICANTE", "MODELO", 
      "TAMANHO", "VALOR DE COMPRA", "DOT", "SITUAÇÃO INICIAL", "STATUS ATUAL", 
      "VIDA ATUAL", "SULCO ATUAL (mm)", "VEÍCULO INSTALADO", "POSIÇÃO ATUAL", 
      "KM TOTAL RODADO", "CUSTO POR KM (CPK)"
    ];

    if (!abaPneu) {
      abaPneu = ss.insertSheet("CADASTRO_PNEU_FROTA");
      abaPneu.appendRow(cabecalhos);
      abaPneu.getRange(1, 1, 1, cabecalhos.length)
             .setFontWeight("bold")
             .setBackground("#111827")
             .setFontColor("#ffffff")
             .setHorizontalAlignment("center");
      abaPneu.setFrozenRows(1);
    }

    var ultimaLinha = abaPneu.getLastRow();
    
    // 1. Gera ID Único Automático (Ex: PNU-0001)
    var proximoId = 1;
    if (ultimaLinha > 1) {
      var ultimoIdTexto = abaPneu.getRange(ultimaLinha, 1).getValue().toString();
      var num = ultimoIdTexto.replace("PNU-", "");
      if (!isNaN(num) && num !== "") proximoId = parseInt(num) + 1;
    }
    var idFormatado = "PNU-" + ("0000" + proximoId).slice(-4);

    // 2. Validação de Duplicidade do Número de Fogo
    if (ultimaLinha > 1) {
      var fogosExistentes = abaPneu.getRange(2, 3, ultimaLinha - 1, 1).getValues().map(r => String(r[0]).trim().toUpperCase());
      if (fogosExistentes.indexOf(String(dados.numeroFogo).trim().toUpperCase()) > -1) {
        return { sucesso: false, mensagem: "❌ Erro: Já existe um pneu cadastrado com o Número de Fogo " + dados.numeroFogo };
      }
    }

    var dataCadastro = new Date();
    var valorCompra = parseFloat(dados.valor) || 0;

    // 3. Montagem da Linha Patrimonial
    var linhaNova = [
      idFormatado,
      dataCadastro,
      dados.numeroFogo.toUpperCase().trim(),
      dados.marca.toUpperCase().trim(),
      dados.modelo.toUpperCase().trim(),
      dados.tamanho.toUpperCase().trim(),
      valorCompra,
      dados.dot.toUpperCase().trim(),
      dados.situacao.toUpperCase(),       // Novo ou Usado
      "EM ESTOQUE",                       // Status Inicial Padrão
      dados.situacao.toUpperCase() === "NOVO" ? "ORIGINAL" : "1ª RECAPAGEM", 
      parseFloat(dados.sulco) || 0,
      "-",                                // Sem veículo inicialmente
      "-",                                // Sem posição inicialmente
      0,                                  // KM inicial zerado
      0                                   // CPK inicial zerado
    ];

    abaPneu.appendRow(linhaNova);
    return { sucesso: true, id: idFormatado, mensagem: "🛞 Pneu " + idFormatado + " (Fogo: " + dados.numeroFogo + ") cadastrado com sucesso!" };
  } catch (e) {
    return { sucesso: false, mensagem: "Erro no servidor: " + e.toString() };
  }
}

// Retorna a contagem atual de pneus para sugerir o próximo Número de Fogo na tela
function obterProximoSequencialFogo() {
  try {
    var aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("CADASTRO_PNEU_FROTA");
    if (!aba) return { contagem: "01", semanaAno: obterSemanaAnoAtual() };
    
    var totalPneus = aba.getLastRow() - 1;
    var proximoNumero = totalPneus + 1;
    
    return {
      contagem: proximoNumero < 10 ? "0" + proximoNumero : String(proximoNumero),
      semanaAno: obterSemanaAnoAtual()
    };
  } catch(e) { return { contagem: "01", semanaAno: "" }; }
}

function obterSemanaAnoAtual() {
  var d = new Date();
  d.setHours(0,0,0,0);
  d.setDate(d.getDate() + 4 - (d.getDay() || 7));
  var anoInicio = new Date(d.getFullYear(),0,1);
  var semana = Math.ceil((((d - anoInicio) / 86400000) + 1)/7);
  var strSemana = semana < 10 ? "0" + semana : String(semana);
  var strAno = String(d.getFullYear()).slice(-2);
  return strSemana + strAno; // Ex: 2526 (Semana 25 de 2026)
}

function listarPneusBackend() {
  try {
    var aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("CADASTRO_PNEU_FROTA");
    if (!aba) return [];
    var ultimaLinha = aba.getLastRow();
    if (ultimaLinha <= 1) return [];

    var dados = aba.getRange(2, 1, ultimaLinha - 1, 16).getValues();
    var lista = [];

    for (var i = 0; i < dados.length; i++) {
      var row = dados[i];
      if (!row[2]) continue;

      var dataStr = row[1];
      if (dataStr instanceof Date) {
        dataStr = ("0"+dataStr.getDate()).slice(-2)+"/"+("0"+(dataStr.getMonth()+1)).slice(-2)+" "+("0"+dataStr.getHours()).slice(-2)+":"+("0"+dataStr.getMinutes()).slice(-2);
      }

      lista.push({
        linhaOrigem: i + 2, id: row[0], dataCadastro: dataStr, numeroFogo: row[2],
        marca: row[3], modelo: row[4], tamanho: row[5], valor: row[6] || 0,
        dot: row[7], situacao: row[8], status: row[9], vida: row[10],
        sulco: row[11], veiculo: row[12], posicao: row[13], kmRodado: row[14] || 0, cpk: row[15] || 0
      });
    }
    return lista.reverse();
  } catch (e) { return []; }
}

// NOTA (revisão): atualizarDadosCadastroPneu estava duplicada aqui e em
// Movimentacao_Pneus.gs. Mantida apenas a versão de Movimentacao_Pneus.gs, que trata
// números com separador decimal em vírgula (padrão BR) e permite forçar o KM Rodado
// manualmente na tela — a versão daqui era mais antiga e menos completa.
