// 1. ANALISA SE O CAMINHÃO TEM OS OU SOS ABERTA ANTES DE LIBERAR
function analisarPendenciasVeiculoBackend(prefixo) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var pendencias = [];
    
    // Procura na Aba SOS
    // NOTA (revisão): a aba real onde os chamados de S.O.S são criados/gravados é "SOS_FROTA"
    // (ver Operacao.gs, que faz insertSheet("SOS_FROTA")). O nome "CENTRAL_SOS" usado aqui
    // originalmente não existe na planilha, então esta checagem de pendência sempre falhava
    // silenciosamente (getSheetByName retornava null e o bloco era ignorado). Adicionados os
    // nomes reais como fallback para a checagem funcionar de fato.
    var abaSOS = ss.getSheetByName("SOS_FROTA") || ss.getSheetByName("CENTRAL_SOS") || ss.getSheetByName("SOS");
    if(abaSOS) {
       var dadosSOS = abaSOS.getDataRange().getValues();
       var countSOS = 0;
       for(var i=1; i<dadosSOS.length; i++) {
          // Assumindo que Prefixo é coluna B (1) e Status é coluna J (9)
          if(String(dadosSOS[i][1]).trim().toUpperCase() === prefixo.toUpperCase() && String(dadosSOS[i][9]).toUpperCase() !== "CONCLUÍDO") {
             countSOS++;
          }
       }
       if(countSOS > 0) pendencias.push("- " + countSOS + " chamado(s) de S.O.S em aberto.");
    }

    // Procura na Aba Ordem de Serviço
    // NOTA (revisão): a aba real de Ordens de Serviço é "ORDEM_SERVICO_FROTA" (ver Toneladas.gs,
    // que faz insertSheet("ORDEM_SERVICO_FROTA")). O nome "ORDEM_SERVICO" usado aqui originalmente
    // não existe na planilha, então esta checagem também falhava silenciosamente. Adicionados os
    // nomes reais como fallback.
    var abaOS = ss.getSheetByName("ORDEM_SERVICO_FROTA") || ss.getSheetByName("ORDEM_SERVICO") || ss.getSheetByName("ORDEM DE SERVIÇO");
    if(abaOS) {
       var dadosOS = abaOS.getDataRange().getValues();
       var countOS = 0;
       for(var j=1; j<dadosOS.length; j++) {
          // Assumindo que Prefixo é coluna B (1) e Status é coluna K (10)
          if(String(dadosOS[j][1]).trim().toUpperCase() === prefixo.toUpperCase() && String(dadosOS[j][10]).toUpperCase() !== "FECHADA") {
             countOS++;
          }
       }
       if(countOS > 0) pendencias.push("- " + countOS + " Ordem(ns) de Serviço em aberto.");
    }

    if(pendencias.length > 0) {
       return { temPendencia: true, msgPendencias: pendencias.join("\n") };
    } else {
       return { temPendencia: false };
    }
  } catch(e) {
    return { temPendencia: false };
  }
}

// 2. EFETIVA A LIBERAÇÃO (FECHA TUDO E ATUALIZA CALENDÁRIO)
// ================= 4. EFETIVAR LIBERAÇÃO E GRAVAR (FORÇANDO ATÉ À ÚLTIMA LINHA FÍSICA) =================
function efetivarLiberacaoVeiculoBackend(prefixoAlvo) {
  try {
    var prefixo = String(prefixoAlvo).trim().toUpperCase(); 
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    var abaManu = ss.getSheetByName("MANUTENÇÃO");
    if (!abaManu) {
       return { sucesso: false, mensagem: "❌ Erro: Aba 'MANUTENÇÃO' não encontrada no banco de dados." };
    }

    var ultimaLinhaFisica = abaManu.getLastRow();
    // Puxa as informações lendo exatamente o formato de texto da tela (Display)
    var dadosM = abaManu.getRange(1, 1, ultimaLinhaFisica, 4).getDisplayValues(); 
    var veiculoEncontrado = false;
    var linhaOndeGravou = 0;

    // Varre a coluna B (Prefixo) procurando o caminhão selecionado
    for (var m = 1; m < dadosM.length; m++) {
       var prefixoPlanilha = String(dadosM[m][1]).trim().toUpperCase(); 
       
       if (prefixoPlanilha === prefixo) {
          // Escreve fisicamente na Coluna 4 (D) daquela linha exata
          abaManu.getRange(m + 1, 4).setValue("DEVOLVIDO"); 
          veiculoEncontrado = true;
          linhaOndeGravou = m + 1;
          break; 
       }
    }

    var msgRelatorio = "";
    if (!veiculoEncontrado) {
       msgRelatorio = "⚠️ O sistema não encontrou o prefixo '" + prefixo + "' listado na Coluna B da aba MANUTENÇÃO. Verifique se o nome está idêntico.";
    } else {
       msgRelatorio = "✅ Sucesso! O veículo " + prefixo + " foi atualizado para DEVOLVIDO na linha " + linhaOndeGravou + " da planilha.\n\n";
    }

    // =========================================================================
    // ENCERRAMENTO DE ORDENS E SOS (Mantido e protegido)
    try {
      // NOTA (revisão): mesma correção acima — aba real é "SOS_FROTA".
      var abaSOS = ss.getSheetByName("SOS_FROTA") || ss.getSheetByName("CENTRAL_SOS") || ss.getSheetByName("SOS");
      if (abaSOS) {
         var dadosS = abaSOS.getDataRange().getValues();
         var colP = encontrarColuna(dadosS, "PREFIXO", "VEICULO");
         var colS = encontrarColuna(dadosS, "STATUS", "SITUAÇÃO");
         if (colP > -1 && colS > -1) {
            for (var i = 1; i < dadosS.length; i++) {
               if (String(dadosS[i][colP]).trim().toUpperCase() === prefixo) {
                  var stat = String(dadosS[i][colS]).toUpperCase();
                  if (stat !== "CONCLUÍDO" && stat !== "FINALIZADO") {
                     abaSOS.getRange(i + 1, colS + 1).setValue("CONCLUÍDO");
                  }
               }
            }
         }
      }

      // NOTA (revisão): mesma correção acima — aba real é "ORDEM_SERVICO_FROTA".
      var abaOS = ss.getSheetByName("ORDEM_SERVICO_FROTA") || ss.getSheetByName("ORDEM_SERVICO") || ss.getSheetByName("ORDEM DE SERVIÇO");
      if (abaOS) {
         var dadosO = abaOS.getDataRange().getValues();
         var colP = encontrarColuna(dadosO, "PREFIXO", "VEICULO");
         var colS = encontrarColuna(dadosO, "STATUS", "SITUAÇÃO");
         if (colP > -1 && colS > -1) {
            for (var j = 1; j < dadosO.length; j++) {
               if (String(dadosO[j][colP]).trim().toUpperCase() === prefixo) {
                  var stat = String(dadosO[j][colS]).toUpperCase();
                  if (stat !== "FECHADA" && stat !== "CONCLUÍDA") {
                     abaOS.getRange(j + 1, colS + 1).setValue("FECHADA");
                  }
               }
            }
         }
      }
    } catch(errSub) {
       msgRelatorio += "(Nota: Houve um detalhe ao fechar os chamados das outras abas, mas o status principal foi processado).\n";
    }
    // =========================================================================

    // Força o banco de dados do Google a salvar as alterações imediatamente
    SpreadsheetApp.flush();

    return { sucesso: true, mensagem: msgRelatorio + "Painel atualizado." };

  } catch(e) {
    return { sucesso: false, mensagem: "❌ Erro interno no servidor: " + e.toString() };
  }
}
