// ================= MÓDULO CALENDÁRIO DE MANUTENÇÃO =================

function configurarAbaCalendario() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var aba = ss.getSheetByName("CALENDARIO_FROTA");
  if (!aba) {
    aba = ss.insertSheet("CALENDARIO_FROTA");
    aba.appendRow(["ID", "PREFIXO", "DATA_PREVISTA", "SERVIÇO", "CRITICIDADE", "TIPO", "STATUS", "CUSTO_EST", "RESPONSÁVEL", "OBS"]);
    aba.getRange("A1:J1").setFontWeight("bold").setBackground("#1e3a8a").setFontColor("#ffffff");
  }
}

function listarEventosCalendario(mes, ano) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var aba = ss.getSheetByName("CALENDARIO_FROTA");
  if (!aba) return [];
  
  var dados = aba.getDataRange().getValues();
  var eventos = [];
  
  for (var i = 1; i < dados.length; i++) {
    var dataOriginal = new Date(dados[i][2]);
    // Filtra pelo mês e ano selecionados na tela
    if (dataOriginal.getMonth() + 1 == mes && dataOriginal.getFullYear() == ano) {
      eventos.push({
        id: dados[i][0],
        prefixo: dados[i][1],
        data: Utilities.formatDate(dataOriginal, "GMT", "yyyy-MM-dd"),
        servico: dados[i][3],
        criticidade: dados[i][4],
        tipo: dados[i][5],
        status: dados[i][6],
        custo: dados[i][7],       // 🔥 AGORA PUXA O CUSTO (Coluna 8)
        responsavel: dados[i][8], // 🔥 AGORA PUXA O RESPONSÁVEL (Coluna 9)
        obs: dados[i][9],         // 🔥 AGORA PUXA A OBSERVAÇÃO (Coluna 10)
        rowIndex: i + 1
      });
    }
  }
  return eventos;
}
// =====================================================================
// FUNÇÕES DE GESTÃO DA AGENDA (SALVAR, EDITAR, STATUS, EXCLUIR)
// =====================================================================
function salvarEventoCalendario(dados) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var aba = ss.getSheetByName("CALENDARIO_FROTA");
    
    if (!aba) {
      aba = ss.insertSheet("CALENDARIO_FROTA");
      aba.appendRow(["ID", "PREFIXO", "DATA_PREVISTA", "SERVIÇO", "CRITICIDADE", "TIPO", "STATUS", "CUSTO_EST", "RESPONSÁVEL", "OBS"]);
      aba.getRange("A1:J1").setFontWeight("bold").setBackground("#1e3a8a").setFontColor("#ffffff");
    }
    
    // =========================================================================
    // INTELIGÊNCIA: Se o rowIndex estiver preenchido, ele ATUALIZA a linha
    // =========================================================================
    if (dados.rowIndex && String(dados.rowIndex).trim() !== "") {
        var linhaCorreta = parseInt(dados.rowIndex); // Garante que é um número
        
        aba.getRange(linhaCorreta, 2).setValue(dados.prefixo);
        aba.getRange(linhaCorreta, 3).setValue(dados.data);
        aba.getRange(linhaCorreta, 4).setValue(dados.servico);
        aba.getRange(linhaCorreta, 5).setValue(dados.criticidade);
        aba.getRange(linhaCorreta, 6).setValue(dados.tipo);
        // A coluna 7 é o "Status", por isso não lhe tocamos (fica o que já lá estava)
        aba.getRange(linhaCorreta, 8).setValue(dados.custo || 0);
        aba.getRange(linhaCorreta, 9).setValue(dados.responsavel);
        aba.getRange(linhaCorreta, 10).setValue(dados.obs);
        
        return { sucesso: true, mensagem: "✅ Manutenção ATUALIZADA com sucesso!" };
        
    } else {
        // =========================================================================
        // Se NÃO tem rowIndex, então é um NOVO agendamento
        // =========================================================================
        var id = "EV" + new Date().getTime();
        aba.appendRow([
          id, 
          dados.prefixo, 
          dados.data, 
          dados.servico, 
          dados.criticidade,
          dados.tipo, 
          "Agendado", 
          dados.custo || 0, 
          dados.responsavel, 
          dados.obs
        ]);
        
        return { sucesso: true, mensagem: "✅ Manutenção AGENDADA com sucesso!" };
    }
  } catch (e) { 
    return { sucesso: false, mensagem: "Erro no servidor: " + e.toString() }; 
  }
}

function alterarStatusCalendario(rowIndex, novoStatus) {
  try {
    var aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("CALENDARIO_FROTA");
    aba.getRange(rowIndex, 7).setValue(novoStatus); // Coluna G é o Status
    return { sucesso: true, mensagem: "✅ Status alterado para: " + novoStatus };
  } catch (e) { return { sucesso: false, mensagem: e.toString() }; }
}

function excluirEventoCalendario(rowIndex) {
  try {
    var aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("CALENDARIO_FROTA");
    aba.deleteRow(rowIndex);
    return { sucesso: true, mensagem: "🗑️ Agendamento excluído da agenda!" };
  } catch (e) { return { sucesso: false, mensagem: e.toString() }; }
}

// =====================================================================
// FUNÇÃO PARA GERAR ALERTAS NO PAINEL GERAL (PRÓXIMOS 7 DIAS)
// =====================================================================
function listarAlertasPainelGeral() {
  var aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("CALENDARIO_FROTA");
  if (!aba) return [];
  
  var dados = aba.getDataRange().getValues();
  var listaAlertas = [];
  var hoje = new Date(); hoje.setHours(0,0,0,0);
  var limite = new Date(); limite.setDate(hoje.getDate() + 7); limite.setHours(23,59,59,999);
  
  for(var i=1; i<dados.length; i++) {
     if(dados[i][6] === "Realizado") continue; // Ignora os já concluídos
     
     var dataEvento = new Date(dados[i][2]);
     // Puxa tudo o que está atrasado (menor que hoje) ou nos próximos 7 dias
     if (dataEvento <= limite) {
        listaAlertas.push({
           prefixo: dados[i][1],
           data: Utilities.formatDate(dataEvento, "GMT", "dd/MM/yyyy"),
           servico: dados[i][3],
           criticidade: dados[i][4],
           status: dados[i][6],
           atrasado: dataEvento < hoje
        });
     }
  }
  
  // Ordena para que os mais urgentes (Atrasados e Alta criticidade) fiquem no topo
  listaAlertas.sort(function(a,b) {
     if (a.atrasado && !b.atrasado) return -1;
     if (!a.atrasado && b.atrasado) return 1;
     if (a.criticidade === "Alta" && b.criticidade !== "Alta") return -1;
     return 0;
  });
  return listaAlertas;
}
// =========================================================================
// MOTOR DE INTELIGÊNCIA ARTIFICIAL: ANÁLISE PREDITIVA DE MANUTENÇÃO
// =========================================================================

function processarManutencaoPreditivaIA() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var abaAbs = ss.getSheetByName("ABASTECIMENTO_FROTA");
    var abaCal = ss.getSheetByName("CALENDARIO_FROTA");
    
    if (!abaAbs || !abaCal) {
      return { sucesso: false, mensagem: "⚠️ Erro: Certifique-se de que as abas ABASTECIMENTO_FROTA e CALENDARIO_FROTA existem." };
    }
    
    var dadosAbs = abaAbs.getDataRange().getValues();
    if (dadosAbs.length < 2) return { sucesso: true, mensagem: "Sem dados de abastecimento suficientes para calcular." };
    
    // 1. Agrupar dados por veículo para calcular a média de rodagem e o KM atual
    var frotaDados = {};
    
    for (var i = 1; i < dadosAbs.length; i++) {
      var prefixo = String(dadosAbs[i][1]).trim().toUpperCase();
      if (!prefixo || prefixo === "-") continue;
      
      // Coluna D (3) = Data | Coluna J (9) = KM TOTAL (Odomômetro)
      var dataPartes = String(dadosAbs[i][3]).split("/");
      var dataObj = dataPartes.length === 3 ? new Date(dataPartes[2], dataPartes[1] - 1, dataPartes[0]) : new Date(dadosAbs[i][3]);
      var kmTotal = parseFloat(dadosAbs[i][9]) || 0;
      
      if (isNaN(dataObj.getTime()) || kmTotal === 0) continue;
      
      if (!frotaDados[prefixo]) {
        frotaDados[prefixo] = { abastecimentos: [] };
      }
      frotaDados[prefixo].abastecimentos.push({ data: dataObj, km: kmTotal });
    }
    
    var agendadosContador = 0;
    var hoje = new Date();
    
    // Puxar agendamentos preditivos existentes para não duplicar
    var eventosExistentes = abaCal.getDataRange().getValues();
    var chavesExistentes = {};
    for (var c = 1; c < eventosExistentes.length; c++) {
      // Se for preditivo, guarda a combinação Prefixo + Serviço para evitar duplicados
      if (eventosExistentes[c][5] === "Preditiva") {
        chavesExistentes[eventosExistentes[c][1] + "_" + eventosExistentes[c][3]] = true;
      }
    }

    // 2. Executar o cálculo preditivo para cada camião
    for (var camiao in frotaDados) {
      var registros = frotaDados[camiao].abastecimentos;
      if (registros.length < 2) continue; // Precisa de pelo menos 2 registros para calcular a média
      
      // Ordena por data (da mais antiga para a mais recente)
      registros.sort((a, b) => a.data - b.data);
      
      var primeiro = registros[0];
      var ultimo = registros[registros.length - 1];
      
      var diferencaTempo = ultimo.data - primeiro.data;
      var diasDecorridos = diferencaTempo / (1000 * 60 * 60 * 24);
      var kmRodados = ultimo.km - primeiro.km;
      
      // Se rodou km e passou tempo, calcula a média diária real de rodagem
      var kmPorDia = (diasDecorridos > 0 && kmRodados > 0) ? (kmRodados / diasDecorridos) : 0;
      if (kmPorDia <= 0) kmPorDia = 50; // Média padrão de segurança caso os dados estejam inconsistentes
      
      var kmAtual = ultimo.km;
      
      // Regra Preditiva: Próxima revisão ideal nos múltiplos de 10.000 KM
      var proximaMetaKm = Math.ceil((kmAtual + 1) / 10000) * 10000;
      var kmQueFaltam = proximaMetaKm - kmAtual;
      
      // Calcula em quantos dias ele vai atingir a meta
      var diasParaMeta = Math.ceil(kmQueFaltam / kmPorDia);
      
      // Projeta a data exata do futuro
      var dataProjetada = new Date();
      dataProjetada.setDate(hoje.getDate() + diasParaMeta);
      
      var dataFmt = Utilities.formatDate(dataProjetada, Session.getScriptTimeZone(), "yyyy-MM-dd");
      var nomeServicoIA = "REVISÃO PREDITIVA (" + (proximaMetaKm/1000) + "K KM)";
      
      // Se a IA já agendou esta revisão específica para este camião, pula para não duplicar!
      if (chavesExistentes[camiao + "_" + nomeServicoIA]) continue;
      
      // 3. Inserir o agendamento gerado pela IA no Calendário
      var idEvento = "IA_" + new Date().getTime() + "_" + camiao;
      abaCal.appendRow([
        idEvento,
        camiao,
        dataFmt,
        nomeServicoIA,
        "Média",
        "Preditiva",
        "Agendado",
        0, // Custo estimado inicial
        "SISTEMA IA INOVAR",
        "Gerado automaticamente com base na média de rodagem real de " + kmPorDia.toFixed(1) + " KM/dia calculada a partir dos abastecimentos."
      ]);
      
      agendadosContador++;
    }
    
    return { sucesso: true, mensagem: "🤖 IA Inovar: Análise concluída! " + agendadosContador + " novas manutenções preditivas foram projetadas e agendadas no calendário!" };
    
  } catch(e) {
    return { sucesso: false, mensagem: "Erro no processamento da IA: " + e.toString() };
  }
}
// =========================================================================
// SISTEMA AUTOMÁTICO DE RELATÓRIOS E ALERTAS POR E-MAIL
// =========================================================================

function enviarRelatorioManutencaoSemanal() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var abaCal = ss.getSheetByName("CALENDARIO_FROTA");
    
    if (!abaCal) return;
    
    var dados = abaCal.getDataRange().getValues();
    if (dados.length < 2) return; // Se só tiver o cabeçalho, não faz nada
    
    var hoje = new Date();
    hoje.setHours(0, 0, 0, 0);
    
    var daquiAUmaSemana = new Date();
    daquiAUmaSemana.setDate(hoje.getDate() + 7);
    daquiAUmaSemana.setHours(23, 59, 59, 999);
    
    var manutencoesDaSemana = [];
    
    // 1. Filtrar os eventos que estão agendados para os próximos 7 dias
    for (var i = 1; i < dados.length; i++) {
      var dataEvento = new Date(dados[i][2]);
      var status = String(dados[i][6]).trim();
      
      if (dataEvento >= hoje && dataEvento <= daquiAUmaSemana && status === "Agendado") {
        manutencoesDaSemana.push({
          id: dados[i][0],
          prefixo: dados[i][1],
          data: Utilities.formatDate(dataEvento, "GMT", "dd/MM/yyyy"),
          servico: dados[i][3],
          criticidade: dados[i][4],
          tipo: dados[i][5],
          custo: parseFloat(dados[i][7]) || 0,
          responsavel: dados[i][8],
          obs: dados[i][9]
        });
      }
    }
    
    // Se não houver nenhuma manutenção agendada para a semana, não envia e-mail para não encher a caixa
    if (manutencoesDaSemana.length === 0) return;
    
    // 2. Ordenar por criticidade (Alta primeiro) e depois por data
    manutencoesDaSemana.sort(function(a, b) {
      if (a.criticidade === "Alta" && b.criticidade !== "Alta") return -1;
      if (a.criticidade !== "Alta" && b.criticidade === "Alta") return 1;
      return new Date(a.data) - new Date(b.data);
    });
    
    // 3. Montar o Corpo do E-mail em HTML profissional
    var emailDestino = Session.getActiveUser().getEmail(); // <-- Envia para o seu e-mail atual. Se quiser outro, mude para: "seu_email@inovar.com"
    var assunto = "📋 [INOVAR FROTA] Agenda Semanal de Manutenções";
    
    var htmlCorpo = "<div style='font-family: Arial, sans-serif; color: #1e293b; max-width: 700px; margin: 0 auto; padding: 20px; border: 1px solid #e2e8f0; border-radius: 8px;'>";
    htmlCorpo += "<h2 style='color: #1e3a8a; border-bottom: 2px solid #3b82f6; padding-bottom: 10px; margin-top: 0;'>🤖 Relatório Semanal de Manutenções da Frota</h2>";
    htmlCorpo += "<p>Olá, Ingrid. O robô da <strong>Inovar Soluções Ambientais</strong> analisou a agenda e o status da frota. Abaixo estão listadas as manutenções programadas para os próximos 7 dias:</p>";
    
    htmlCorpo += "<table style='width: 100%; border-collapse: collapse; margin-top: 20px; font-size: 13px;'>";
    htmlCorpo += "<thead><tr style='background-color: #1e3a8a; color: white;'>";
    htmlCorpo += "<th style='padding: 10px; border: 1px solid #cbd5e1; text-align: left;'>Camião</th>";
    htmlCorpo += "<th style='padding: 10px; border: 1px solid #cbd5e1; text-align: left;'>Data</th>";
    htmlCorpo += "<th style='padding: 10px; border: 1px solid #cbd5e1; text-align: left;'>Serviço</th>";
    htmlCorpo += "<th style='padding: 10px; border: 1px solid #cbd5e1; text-align: center;'>Criticidade</th>";
    htmlCorpo += "<th style='padding: 10px; border: 1px solid #cbd5e1; text-align: left;'>Tipo</th>";
    htmlCorpo += "<th style='padding: 10px; border: 1px solid #cbd5e1; text-align: left;'>Responsável</th>";
    htmlCorpo += "</tr></thead><tbody>";
    
    manutencoesDaSemana.forEach(function(m) {
      var corTexto = "#475569";
      var corFundo = "#f1f5f9";
      
      if (m.criticidade === "Alta") { corTexto = "#b91c1c"; corFundo = "#fee2e2"; }
      else if (m.criticidade === "Média") { corTexto = "#b45309"; corFundo = "#fef3c7"; }
      else if (m.criticidade === "Baixa") { corTexto = "#065f46"; corFundo = "#d1fae5"; }
      
      htmlCorpo += "<tr style='border-bottom: 1px solid #cbd5e1;'>";
      htmlCorpo += "<td style='padding: 10px; font-weight: bold; color: #1e3a8a;'>" + m.prefixo + "</td>";
      htmlCorpo += "<td style='padding: 10px; white-space: nowrap;'>" + m.data + "</td>";
      htmlCorpo += "<td style='padding: 10px; text-transform: uppercase;'><strong>" + m.servico + "</strong>" + (m.obs ? "<br><span style='font-size:11px; color:#64748b;'>Obs: " + m.obs + "</span>" : "") + "</td>";
      htmlCorpo += "<td style='padding: 10px; text-align: center;'><span style='background-color:" + corFundo + "; color:" + corTexto + "; padding: 3px 8px; border-radius: 4px; font-weight: bold; font-size: 11px;'>" + m.criticidade.toUpperCase() + "</span></td>";
      htmlCorpo += "<td style='padding: 10px;'>" + m.tipo + "</td>";
      htmlCorpo += "<td style='padding: 10px; text-transform: uppercase; font-size:11px; color:#64748b;'>" + (m.responsavel || "-") + "</td>";
      htmlCorpo += "</tr>";
    });
    
    htmlCorpo += "</tbody></table>";
    
    htmlCorpo += "<p style='margin-top: 30px; font-size: 11px; color: #94a3b8; border-top: 1px solid #e2e8f0; padding-top: 10px; text-align: center;'>";
    htmlCorpo += "Este é um e-mail automático gerado pelo módulo preditivo do Sistema Inovar Frota. Não responda a esta mensagem.";
    htmlCorpo += "</p></div>";
    
    // 4. Efetuar o envio do e-mail corporativo
    MailApp.sendEmail({
      to: emailDestino,
      subject: assunto,
      htmlBody: htmlCorpo
    });
    
  } catch(e) {
    Logger.log("Erro ao enviar relatório por e-mail: " + e.toString());
  }
}
// =========================================================================
// NOVO MOTOR DE GRAVAÇÃO (CRIAR E EDITAR) - À PROVA DE FALHAS
// =========================================================================
function processarGravacaoCalendario(dados) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var aba = ss.getSheetByName("CALENDARIO_FROTA");
    
    if (!aba) {
      aba = ss.insertSheet("CALENDARIO_FROTA");
      aba.appendRow(["ID", "PREFIXO", "DATA_PREVISTA", "SERVIÇO", "CRITICIDADE", "TIPO", "STATUS", "CUSTO_EST", "RESPONSÁVEL", "OBS"]);
      aba.getRange("A1:J1").setFontWeight("bold").setBackground("#1e3a8a").setFontColor("#ffffff");
    }
    
    // SE VIER COM O NÚMERO DA LINHA (rowIndex), ELE ATUALIZA A LINHA EXATA!
    if (dados.rowIndex && String(dados.rowIndex).trim() !== "") {
        var linhaCorreta = parseInt(dados.rowIndex);
        
        aba.getRange(linhaCorreta, 2).setValue(String(dados.prefixo).toUpperCase());
        aba.getRange(linhaCorreta, 3).setValue(dados.data);
        aba.getRange(linhaCorreta, 4).setValue(String(dados.servico).toUpperCase());
        aba.getRange(linhaCorreta, 5).setValue(dados.criticidade);
        aba.getRange(linhaCorreta, 6).setValue(dados.tipo);
        // Não mexe no status (coluna 7)
        aba.getRange(linhaCorreta, 8).setValue(parseFloat(dados.custo) || 0);
        aba.getRange(linhaCorreta, 9).setValue(String(dados.responsavel).toUpperCase());
        aba.getRange(linhaCorreta, 10).setValue(dados.obs);
        
        return { sucesso: true, mensagem: "✅ Manutenção ATUALIZADA com sucesso na linha " + linhaCorreta + "!" };
        
    } else {
        // SE NÃO TIVER LINHA, ELE CRIA UMA NOVA
        var id = "EV" + new Date().getTime();
        aba.appendRow([
          id, 
          String(dados.prefixo).toUpperCase(), 
          dados.data, 
          String(dados.servico).toUpperCase(), 
          dados.criticidade,
          dados.tipo, 
          "Agendado", 
          parseFloat(dados.custo) || 0, 
          String(dados.responsavel).toUpperCase(), 
          dados.obs
        ]);
        
        return { sucesso: true, mensagem: "✅ Manutenção AGENDADA com sucesso!" };
    }
  } catch (e) { 
    return { sucesso: false, mensagem: "Erro no servidor: " + e.toString() }; 
  }
}
