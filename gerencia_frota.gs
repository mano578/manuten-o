// =========================================================================
// 1. ROBÔ AUTÔNOMO DE REGISTRO DE DISPONIBILIDADE (MATEMÁTICA CORRIGIDA)
// =========================================================================
function registrarCheckDiarioFrotaAutonomo() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    // 1. LER ABA MANUTENÇÃO PRIMEIRO (Criar o Dicionário de Status)
    // O sistema "decora" como está o status (Coluna D) de cada Prefixo (Coluna B)
    var abaManu = ss.getSheetByName("MANUTENÇÃO");
    if (!abaManu) return;
    
    var dadosManu = abaManu.getDataRange().getValues();
    var statusDosVeiculos = {};

    for (var m = 1; m < dadosManu.length; m++) {
      var prefixoManu = String(dadosManu[m][1]).trim().toUpperCase(); // Coluna B (Prefixo)
      var statusManu = String(dadosManu[m][3]).trim().toUpperCase();  // Coluna D (Status)
      if (prefixoManu !== "") {
        statusDosVeiculos[prefixoManu] = statusManu;
      }
    }

    // 2. LER CADASTRO E FAZER A CONTAGEM EXATA (Coluna C e Coluna Y)
    var abaCadastro = ss.getSheetByName("CADASTRO_FROTA") || ss.getSheetByName("CADASTRO_VEICULOS");
    if (!abaCadastro) return;
    
    var dadosCad = abaCadastro.getDataRange().getValues();
    var controleTipos = {};

    for (var i = 1; i < dadosCad.length; i++) {
      var prefixoCad = String(dadosCad[i][2]).trim().toUpperCase(); // Coluna C (Prefixo)
      var tipoEquipamento = String(dadosCad[i][24]).trim().toUpperCase(); // Coluna Y (Informação/Tipo)

      // Se a linha tem um prefixo cadastrado e um tipo definido na Coluna Y
      if (prefixoCad !== "" && tipoEquipamento !== "") {
        
        // Se este tipo (ex: CAÇAMBA 6M) ainda não estiver na lista, cria o espaço para ele
        if (!controleTipos[tipoEquipamento]) {
          controleTipos[tipoEquipamento] = { totalReal: 0, liberadosAgora: 0 };
        }

        // Soma +1 no Total Real deste tipo de equipamento
        controleTipos[tipoEquipamento].totalReal++;

        // CRUZA COM A MANUTENÇÃO: O veículo exato desta linha está liberado lá?
        var statusAtual = statusDosVeiculos[prefixoCad] || "";
        if (statusAtual === "LIBERADO" || statusAtual === "DEVOLVIDO" || statusAtual === "OK") {
          // Se sim, soma +1 nos Liberados Agora
          controleTipos[tipoEquipamento].liberadosAgora++;
        }
      }
    }

    // 3. PREPARAR A ABA GERENCIA_FROTA PARA GRAVAR
    var abaGerencia = ss.getSheetByName("GERENCIA_FROTA");
    if (!abaGerencia) {
      abaGerencia = ss.insertSheet("GERENCIA_FROTA");
      abaGerencia.appendRow(["DATA", "EQUIPAMENTO", "TOTAL REAL", "CHECK 8:00", "CHECK 11:00", "CHECK 15:00", "CHECK 19:00", "MÉDIA DO DIA"]);
      abaGerencia.getRange("A1:H1").setFontWeight("bold").setBackground("#4ade80");
    }

    var dataDeHoje = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy");
    var horaAtual = new Date().getHours();

    // Descobrir qual coluna deve ser preenchida baseado na hora que o Google rodou
    var colunaAlvo = 0;
    if (horaAtual >= 7 && horaAtual <= 9) colunaAlvo = 4;       // 8:00 -> Coluna D
    else if (horaAtual >= 10 && horaAtual <= 13) colunaAlvo = 5;// 11:00 -> Coluna E
    else if (horaAtual >= 14 && horaAtual <= 17) colunaAlvo = 6;// 15:00 -> Coluna F
    else if (horaAtual >= 18 && horaAtual <= 21) colunaAlvo = 7;// 19:00 -> Coluna G
    else return; // Fora do horário comercial estipulado, ignora a gravação.

    var dadosGerencia = abaGerencia.getDataRange().getValues();
    var linhasDeHoje = {}; 

    // Achar as linhas que correspondem ao dia de hoje na planilha
    for (var g = 1; g < dadosGerencia.length; g++) {
      var dataG = dadosGerencia[g][0];
      if (dataG instanceof Date) {
        dataG = Utilities.formatDate(dataG, Session.getScriptTimeZone(), "dd/MM/yyyy");
      } else {
        dataG = String(dataG).trim();
      }
      var tipoG = String(dadosGerencia[g][1]).trim().toUpperCase();
      
      if (dataG === dataDeHoje) {
        linhasDeHoje[tipoG] = g + 1; 
      }
    }

    // 4. GRAVAR NA PLANILHA E CALCULAR A MÉDIA
    var tiposArray = Object.keys(controleTipos).sort();

    for (var t = 0; t < tiposArray.length; t++) {
      var tipo = tiposArray[t];
      var info = controleTipos[tipo];
      var linha = linhasDeHoje[tipo];

      if (linha) {
         // Atualiza a medição no horário correspondente
         abaGerencia.getRange(linha, colunaAlvo).setValue(info.liberadosAgora);
         abaGerencia.getRange(linha, 3).setValue(info.totalReal); 
         
         // Cálculo matemático da média (soma só os horários que já passaram)
         var checksHoje = abaGerencia.getRange(linha, 4, 1, 4).getValues()[0];
         var soma = 0, divisor = 0;
         for (var c = 0; c < 4; c++) { 
           if (checksHoje[c] !== "" && !isNaN(checksHoje[c])) { 
             soma += parseInt(checksHoje[c]); 
             divisor++; 
           } 
         }
         var media = divisor > 0 ? Math.round(soma / divisor) : info.liberadosAgora;
         abaGerencia.getRange(linha, 8).setValue(media); 

      } else {
         // Cria a linha na planilha caso seja o primeiro check do dia para este tipo
         var novaLinha = [dataDeHoje, tipo, info.totalReal, "", "", "", "", ""];
         novaLinha[colunaAlvo - 1] = info.liberadosAgora;
         novaLinha[7] = info.liberadosAgora; 
         abaGerencia.appendRow(novaLinha);
      }
    }
    
    // Força o Google a gravar na hora
    SpreadsheetApp.flush();
  } catch (e) {
    Logger.log("Erro: " + e.message);
  }
}
function iniciarRotinaInvisivel() {
  // Limpa programações antigas para não sobrecarregar
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'registrarCheckDiarioFrotaAutonomo' || 
        triggers[i].getHandlerFunction() === 'registrarCheckDiarioFrota') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  
  // Instrui o servidor do Google a rodar a função, independente do Web App estar aberto
  ScriptApp.newTrigger('registrarCheckDiarioFrotaAutonomo').timeBased().atHour(8).everyDays(1).create();
  ScriptApp.newTrigger('registrarCheckDiarioFrotaAutonomo').timeBased().atHour(11).everyDays(1).create();
  ScriptApp.newTrigger('registrarCheckDiarioFrotaAutonomo').timeBased().atHour(15).everyDays(1).create();
  ScriptApp.newTrigger('registrarCheckDiarioFrotaAutonomo').timeBased().atHour(19).everyDays(1).create();
}

// =========================================================================
// 2. FUNÇÃO QUE ENTREGA OS DADOS PARA O WEB APP DESENHAR A TELA
// =========================================================================
function buscarRelatorioGerencia(dataBuscada) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var aba = ss.getSheetByName("GERENCIA_FROTA");
  if (!aba) return [];
  
  if (!dataBuscada) {
    dataBuscada = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy");
  }

  var dados = aba.getDataRange().getValues();
  var relatorio = [];

  for (var i = 1; i < dados.length; i++) {
    var dataPlanilha = dados[i][0];
    if (dataPlanilha instanceof Date) dataPlanilha = Utilities.formatDate(dataPlanilha, Session.getScriptTimeZone(), "dd/MM/yyyy");
    
    if (String(dataPlanilha).trim() === dataBuscada) {
      relatorio.push({
        equipamento: dados[i][1],
        totalReal: dados[i][2],
        check8: dados[i][3] === "" ? "-" : dados[i][3],
        check11: dados[i][4] === "" ? "-" : dados[i][4],
        check15: dados[i][5] === "" ? "-" : dados[i][5],
        check19: dados[i][6] === "" ? "-" : dados[i][6],
        media: dados[i][7] === "" ? "-" : dados[i][7]
      });
    }
  }
  return relatorio;
}
function criarDespertadoresDeChecagem() {
  // Apaga gatilhos antigos para não duplicar
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'registrarCheckDiarioFrota') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  
  // Cria os 4 relógios (8h, 11h, 15h e 19h)
  ScriptApp.newTrigger('registrarCheckDiarioFrota').timeBased().atHour(8).everyDays(1).create();
  ScriptApp.newTrigger('registrarCheckDiarioFrota').timeBased().atHour(11).everyDays(1).create();
  ScriptApp.newTrigger('registrarCheckDiarioFrota').timeBased().atHour(15).everyDays(1).create();
  ScriptApp.newTrigger('registrarCheckDiarioFrota').timeBased().atHour(19).everyDays(1).create();
}
