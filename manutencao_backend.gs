// ================= MOTOR DO CENTRO DE CONTROLE DE MANUTENÇÃO =================

// 1. INICIALIZAR A ABA DE TRABALHO
function iniciarAbaPlanoManutencao() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var aba = ss.getSheetByName("PLANO_MANUTENCAO");
  
  if (!aba) {
    aba = ss.insertSheet("PLANO_MANUTENCAO");
    var cabecalhos = [
      "ID", "CATEGORIA", "DATA REALIZAÇÃO", "PREFIXO", "PLACA", "TIPO EQUIPAMENTO", 
      "TIPO MEDIÇÃO", "KM/HORA REALIZADO", "INTERVALO", "ALVO DA PRÓXIMA (KM/DATA)", 
      "TIPO MANUTENÇÃO", "DESCRIÇÃO / COMPONENTE", "EXECUTOR", "LINK ARQUIVO", "CONCLUÍDO (SIM/NAO)", "DATA DA BAIXA"
    ];
    aba.appendRow(cabecalhos);
    aba.getRange("A1:P1").setFontWeight("bold").setBackground("#0f172a").setFontColor("#ffffff");
    aba.setFrozenRows(1);
  } else {
    // Garante que a coluna DATA DA BAIXA existe em planilhas antigas
    if (aba.getRange("P1").getValue() === "") {
        aba.getRange("P1").setValue("DATA DA BAIXA").setFontWeight("bold").setBackground("#0f172a").setFontColor("#ffffff");
    }
  }
  return aba;
}

// 2. SALVAR NOVO REGISTRO E ARMAR O RADAR (COM SUPORTE A AGENDAMENTOS E HISTÓRICO)
function salvarManutencaoUnificadaBackend(dados) {
  try {
    var aba = iniciarAbaPlanoManutencao();
    var idFmt = "MNT-" + new Date().getTime().toString().slice(-6);
    
    var urlArquivo = "";
    if (dados.arquivoBase64) {
       var blobArquivo = Utilities.newBlob(Utilities.base64Decode(dados.arquivoBase64), dados.arquivoMimeType, dados.arquivoNome);
       var pastas = DriveApp.getFoldersByName("FROTA_MANUTENCAO");
       var pasta = pastas.hasNext() ? pastas.next() : DriveApp.createFolder("FROTA_MANUTENCAO");
       var file = pasta.createFile(blobArquivo);
       file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
       urlArquivo = file.getUrl();
    }

    var catSegura = String(dados.categoria || "").toUpperCase();
    var prefSeguro = String(dados.prefixo || "").toUpperCase();
    var placaSegura = String(dados.placa || "").toUpperCase();
    var equipSeguro = String(dados.tipoEquipamento || "").toUpperCase(); 
    var tipoMedSeguro = String(dados.tipoMedicao || "").toUpperCase();
    var tipoMntSeguro = String(dados.tipoManutencao || "").toUpperCase();
    var execSeguro = String(dados.executor || "").toUpperCase();
    
    var statusReg = String(dados.statusRegistro || "REALIZADO").toUpperCase();
    var concluidoVal = (statusReg === "HISTORICO") ? "SIM" : "NAO"; 
    
    // Se for lançado diretamente como histórico, já salva a data de hoje como baixa automática
    var dataBaixaVal = (statusReg === "HISTORICO") ? Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy") : "";
    var intervaloNum = parseFloat(dados.intervalo) || 0;

    if (statusReg === "AGENDADO") {
       intervaloNum = 0;
    } else {
       if (catSegura.includes("TACÓGRAFO") || catSegura.includes("TACOGRAFO") || tipoMntSeguro.includes("TACÓGRAFO") || tipoMntSeguro.includes("TACOGRAFO")) {
          if (tipoMedSeguro !== "KM" && tipoMedSeguro !== "HORAS") {
            tipoMedSeguro = "ANOS";
            if (intervaloNum === 0) intervaloNum = 2; 
          }
       }
    }

    var alvoProxima = "";
    if (tipoMedSeguro === "KM") {
       alvoProxima = (parseFloat(dados.kmHoraAtual) || 0) + intervaloNum;
    } 
    else if (tipoMedSeguro === "HORAS") {
       var diasParaFrente = Math.ceil((intervaloNum / 108) * 7);
       var dataBase = new Date(dados.dataRealizacao + "T12:00:00");
       dataBase.setDate(dataBase.getDate() + diasParaFrente);
       alvoProxima = Utilities.formatDate(dataBase, Session.getScriptTimeZone(), "yyyy-MM-dd");
    } 
    else if (tipoMedSeguro === "DIAS") {
       var dataBaseDias = new Date(dados.dataRealizacao + "T12:00:00");
       dataBaseDias.setDate(dataBaseDias.getDate() + intervaloNum);
       alvoProxima = Utilities.formatDate(dataBaseDias, Session.getScriptTimeZone(), "yyyy-MM-dd");
    }
    else if (tipoMedSeguro === "ANOS") {
       var dataBaseAnos = new Date(dados.dataRealizacao + "T12:00:00");
       dataBaseAnos.setFullYear(dataBaseAnos.getFullYear() + intervaloNum);
       alvoProxima = Utilities.formatDate(dataBaseAnos, Session.getScriptTimeZone(), "yyyy-MM-dd");
    }

    var descFinal = dados.descricao;
    if (statusReg === "AGENDADO") {
       descFinal = "[AGENDADO] " + descFinal;
    }

    aba.appendRow([
      idFmt, catSegura, dados.dataRealizacao, prefSeguro, 
      placaSegura, equipSeguro, tipoMedSeguro, 
      dados.kmHoraAtual, dados.intervalo, alvoProxima, tipoMntSeguro, 
      descFinal, execSeguro, urlArquivo, concluidoVal, dataBaixaVal // 🔥 ADICIONADO DATA DA BAIXA
    ]);

    return { sucesso: true, mensagem: "✅ Registro processado e salvo com sucesso!" };
  } catch (e) {
    return { sucesso: false, mensagem: "Erro ao salvar: " + e.toString() };
  }
}

// 3. LER OS ALERTAS E HISTÓRICO 
function listarRadarManutencaoBackend() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var abaPlano = iniciarAbaPlanoManutencao();
    
    var dadosExibicao = abaPlano.getDataRange().getDisplayValues();
    var dadosReais = abaPlano.getDataRange().getValues();
    var alertas = [];

    var kmAtualFrota = {};
    var abaAbs = ss.getSheetByName("ABASTECIMENTO_FROTA");
    
    if (abaAbs) {
      var dadosAbs = abaAbs.getDataRange().getValues();
      for (var a = 1; a < dadosAbs.length; a++) {
         var prefAbs = String(dadosAbs[a][1]).trim().toUpperCase(); 
         var kmLido = parseFloat(dadosAbs[a][8]); 
         
         if (!isNaN(kmLido) && prefAbs !== "") {
            if (!kmAtualFrota[prefAbs] || kmLido > kmAtualFrota[prefAbs]) {
                kmAtualFrota[prefAbs] = kmLido;
            }
         }
      }
    }

    for (var i = 1; i < dadosReais.length; i++) {
      if (dadosReais[i][0] === "") continue; 

      var isConcluido = (dadosReais[i][14] === "SIM");
      var dataBaixaExibicao = dadosExibicao[i][15] || ""; // Lê a nova coluna

      var alerta = {
        linha: i + 1, 
        id: dadosExibicao[i][0], 
        categoria: dadosExibicao[i][1], 
        dataRealizada: dadosExibicao[i][2], 
        prefixo: dadosExibicao[i][3], 
        placa: dadosExibicao[i][4], 
        tipoEquipamento: dadosExibicao[i][5], 
        medicao: dadosExibicao[i][6],
        kmRealizado: parseFloat(dadosReais[i][7]) || 0,
        descricao: dadosExibicao[i][11], 
        arquivo: dadosExibicao[i][13],
        isConcluido: isConcluido,
        dataBaixa: dataBaixaExibicao, // Guarda a data da baixa para uso visual se necessário
        status: "NORMAL", 
        cor: "#10b981", 
        prioridade: 0
      };

      var medicao = String(dadosExibicao[i][6]).trim().toUpperCase();

      if(isConcluido) {
          alerta.status = "CONCLUÍDO";
          alerta.cor = "#94a3b8";
          alerta.alertaTxt = `✅ ARQUIVADO`;
          alerta.infoDetalhe = "Manutenção Finalizada" + (dataBaixaExibicao ? ` em ${dataBaixaExibicao}` : "");
          alerta.prioridade = 999999;
      } 
      else {
          if (medicao === "HORAS" || medicao === "DIAS" || medicao === "ANOS") {
            var alvoData = dadosReais[i][9]; 
            var dataAlvo = new Date(); 
            
            if (alvoData instanceof Date) {
              dataAlvo = alvoData;
            } else if (alvoData && String(alvoData).trim() !== "") {
              dataAlvo = new Date(String(alvoData).substring(0, 10) + "T12:00:00");
            }
            
            if (isNaN(dataAlvo.getTime())) dataAlvo = new Date(); 
            
            var hoje = new Date(); hoje.setHours(12,0,0,0);
            var diasFaltantes = Math.ceil((dataAlvo - hoje) / 86400000) || 0;

            if (diasFaltantes < 0) {
              alerta.status = "ATRASADO"; alerta.cor = "#ef4444"; alerta.alertaTxt = `⚠️ ATRASADO ${Math.abs(diasFaltantes)} DIAS`; alerta.prioridade = -9999 + diasFaltantes;
            } else if (diasFaltantes <= 30) { 
              alerta.status = "AGENDAR"; alerta.cor = "#f59e0b"; alerta.alertaTxt = `🔔 VENCE EM ${diasFaltantes} DIAS`; alerta.prioridade = diasFaltantes;
            } else {
              alerta.alertaTxt = `✅ Faltam ${diasFaltantes} Dias`; alerta.prioridade = diasFaltantes + 1000;
            }
            alerta.infoDetalhe = "Previsão: " + Utilities.formatDate(dataAlvo, Session.getScriptTimeZone(), "dd/MM/yy");
          } 
          else if (medicao === "KM") {
            var kmAlvo = parseFloat(dadosReais[i][9]) || 0; 
            var kmAtual = kmAtualFrota[alerta.prefixo] || alerta.kmRealizado || 0;
            
            var kmFaltante = kmAlvo - kmAtual;
            if (isNaN(kmFaltante)) kmFaltante = 0; 

            if (kmFaltante < 0) {
              alerta.status = "ATRASADO"; alerta.cor = "#ef4444"; alerta.alertaTxt = `⚠️ ATRASOU ${Math.abs(kmFaltante)} KM`; alerta.prioridade = -9999 + kmFaltante;
            } else if (kmFaltante <= 1500) {
              alerta.status = "AGENDAR"; alerta.cor = "#f59e0b"; alerta.alertaTxt = `🔔 FALTAM ${kmFaltante} KM`; alerta.prioridade = kmFaltante;
            } else {
              alerta.alertaTxt = `✅ Faltam ${kmFaltante} KM`; alerta.prioridade = kmFaltante + 100000;
            }
            alerta.infoDetalhe = `Atual: ${kmAtual} KM | Troca: ${kmAlvo} KM`;
          }
      }

      alertas.push(alerta);
    }

    alertas.sort(function(a, b) { return a.prioridade - b.prioridade; }); 
    return { sucesso: true, dados: alertas };
    
  } catch(e) { 
    return { sucesso: false, mensagem: "Erro no servidor (Linha " + e.lineNumber + "): " + e.message }; 
  }
}

// 4. MARCAR COMO CONCLUÍDO COM DATA DA BAIXA 🔥
function concluirManutencaoRadarBackend(linha, idRegistro, dataBaixa) {
  var aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("PLANO_MANUTENCAO");
  if(aba && aba.getRange(linha, 1).getValue() === idRegistro) {
     aba.getRange(linha, 15).setValue("SIM"); // Coluna O é a CONCLUÍDO
     
     // Formata a data de baixa para gravar bonito no Excel se foi enviada (DD/MM/YYYY)
     var dataFormatada = dataBaixa;
     if (dataBaixa && dataBaixa.includes("-")) {
         var partes = dataBaixa.split("-");
         dataFormatada = partes[2] + "/" + partes[1] + "/" + partes[0];
     }
     
     aba.getRange(linha, 16).setValue(dataFormatada || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy")); // Coluna P é a DATA DA BAIXA
     return { sucesso: true, mensagem: "Manutenção arquivada e concluída!" };
  }
  return { sucesso: false, mensagem: "Erro ao localizar registro." };
}
