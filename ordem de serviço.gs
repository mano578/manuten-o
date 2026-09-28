// ================= MOTOR DE ORDENS DE SERVIÇO (CORRIGIDO E MULTI-ITENS) =================
 
function salvarNovaOrdemServico(dados) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var abaOS = ss.getSheetByName("ORDEM_SERVICO_FROTA");
    
    var cabecalhos = [
      "ID O.S.", "DATA/HORA INICIAL", "STATUS", "NÍVEL PRIORIDADE", "TIPO MANUTENÇÃO", 
      "SITUAÇÃO DA PARADA", "SOLICITANTE", "PREFIXO", "PLACA", "CENTRO DE CUSTO", 
      "KM ATUAL", "TIPO DE O.S.", "SETOR RESPONSÁVEL", "FORNECEDOR", "DEFEITO/RELATO", 
      "PRODUTO/PEÇA", "QUANTIDADE PROD", "VALOR UNIT PROD", "VALOR TOTAL PROD", 
      "SERVIÇO PRESTADO", "QUANTIDADE SERV", "VALOR UNIT SERV", "VALOR TOTAL SERV", 
      "CUSTO TOTAL O.S.", "DIAGNÓSTICO TÉCNICO", "MECÂNICO/RESPONSÁVEL", "LINK ASSINATURA", 
      "LINKS ANEXOS/PDF", "DATA/HORA FECHAMENTO", "TEMPO TOTAL REPARO"
    ];

    if (!abaOS) {
      abaOS = ss.insertSheet("ORDEM_SERVICO_FROTA");
      abaOS.appendRow(cabecalhos);
      abaOS.getRange(1, 1, 1, cabecalhos.length).setFontWeight("bold").setBackground("#1e3a8a").setFontColor("#ffffff").setHorizontalAlignment("center");
      abaOS.setFrozenRows(1);
    }

    var ultimaLinha = abaOS.getLastRow();
    var proximoIdNumerico = 1;
    if (ultimaLinha > 1) {
      var ultimoIdTexto = abaOS.getRange(ultimaLinha, 1).getValue().toString();
      var numeroExtraido = ultimoIdTexto.replace("OS-", "");
      if (!isNaN(numeroExtraido) && numeroExtraido !== "") {
        proximoIdNumerico = parseInt(numeroExtraido) + 1;
      }
    }
    var idFormatado = "OS-" + ("0000" + proximoIdNumerico).slice(-4);

    var placaVeiculo = dados.placa;
    var centroCustoVeiculo = dados.centroCusto;

    var strProdutos = dados.produtos.map(p => p.nome).join('\n');
    var strQtdProd = dados.produtos.map(p => p.qtd).join('\n');
    var strValProd = dados.produtos.map(p => "R$ " + p.valor.toFixed(2)).join('\n');
    var totalProd = dados.produtos.reduce((acc, p) => acc + (p.qtd * p.valor), 0);

    var strServicos = dados.servicos.map(s => s.nome).join('\n');
    var strQtdServ = dados.servicos.map(s => s.qtd).join('\n');
    var strValServ = dados.servicos.map(s => "R$ " + s.valor.toFixed(2)).join('\n');
    var totalServ = dados.servicos.reduce((acc, s) => acc + (s.qtd * s.valor), 0);

    var custoTotalOS = totalProd + totalServ;

    var linhaNova = [
      idFormatado,                     
      dados.dataInicial, 
      "ABERTO",                        
      dados.prioridade.toUpperCase(),  
      dados.tipoManutencao.toUpperCase(), 
      dados.situacaoParada.toUpperCase(), 
      dados.solicitante.toUpperCase(), 
      dados.prefixo.toUpperCase(),     
      placaVeiculo,                    
      centroCustoVeiculo,              
      dados.km,           
      dados.tipoOS.toUpperCase(),      
      dados.setor.toUpperCase(),       
      dados.tipoOS.toUpperCase() === "EXTERNO" ? dados.fornecedor.toUpperCase() : "-", 
      dados.defeito,  // <--- CORREÇÃO AQUI (Estava a puxar "dados.descricao")               
      strProdutos,                         
      strQtdProd,                         
      strValProd,                         
      totalProd,                       
      strServicos,                   
      strQtdServ,                         
      strValServ,                         
      totalServ,                       
      custoTotalOS,                    
      dados.diagnostico,               
      dados.mecanico.toUpperCase(),    
      dados.linkAssinatura || "",      
      dados.linksAnexos || "",         
      "",                              
      ""                               
    ];

    abaOS.appendRow(linhaNova);

    var abaManutencao = ss.getSheetByName("MANUTENÇÃO");
    if (abaManutencao && dados.situacaoParada.toUpperCase() === "NÃO PROGRAMADA") {
      var dadosMan = abaManutencao.getDataRange().getValues();
      for (var j = 1; j < dadosMan.length; j++) {
        if (String(dadosMan[j][1]).trim().toUpperCase() === dados.prefixo.trim().toUpperCase()) {
          abaManutencao.getRange(j + 1, 4).setValue("RETIDO - O.S.");
          break;
        }
      }
    }

    return { sucesso: true, id: idFormatado, mensagem: "✅ Ordem de Serviço " + idFormatado + " aberta com sucesso no sistema!" };
  } catch (e) {
    return { sucesso: false, mensagem: "Erro ao registrar O.S.: " + e.toString() };
  }
}

// Inteligência para buscar o KM na aba ABASTECIMENTO_FROTA
// Inteligência para buscar o KM na aba ABASTECIMENTO_FROTA (Pela data mais recente)
function buscarKmAbastecimento(prefixo) {
  try {
    var aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("ABASTECIMENTO_FROTA");
    if (!aba) return "";
    var dados = aba.getDataRange().getValues();

    var ultimaDataMili = 0;
    var kmEncontrado = "";

    // A=0(ID), B=1(VEÍCULO), C=2(PLACA), D=3(DATA), E=4, F=5, G=6, H=7, I=8(KM)
    for (var i = 1; i < dados.length; i++) {
      if (String(dados[i][1]).trim().toUpperCase() === prefixo.trim().toUpperCase()) {
        var dataAbast = dados[i][3]; // Lê a Coluna D (Data)
        var kmLido = parseFloat(dados[i][8]); // Lê a Coluna I (KM)

        if (dataAbast instanceof Date && !isNaN(kmLido)) {
          // Se a data desta linha for maior (mais recente) que a que temos na memória
          if (dataAbast.getTime() >= ultimaDataMili) {
            ultimaDataMili = dataAbast.getTime();
            kmEncontrado = kmLido;
          }
        } else if (!isNaN(kmLido)) {
          // Fallback: se a data não estiver bem formatada, pega o KM que estiver mais abaixo na tabela
          kmEncontrado = kmLido;
        }
      }
    }
    return kmEncontrado !== "" ? kmEncontrado : "";
  } catch (e) { return ""; }
}

// Função para Excluir permanentemente uma O.S.
function excluirOrdemServicoBackend(linha) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var abaOS = ss.getSheetByName("ORDEM_SERVICO_FROTA");
    if (!abaOS) return { sucesso: false, mensagem: "Aba não encontrada." };
    
    abaOS.deleteRow(linha); // Exclui a linha inteira da planilha
    return { sucesso: true, mensagem: "🗑️ Ordem de Serviço excluída com sucesso!" };
  } catch(e) { return { sucesso: false, mensagem: "Erro ao excluir: " + e.toString() }; }
}

function fecharOrdemServicoBackend(linhaOS, prefixo) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var abaOS = ss.getSheetByName("ORDEM_SERVICO_FROTA");
    if (!abaOS) return { sucesso: false, mensagem: "Aba de Ordens de Serviço não localizada." };

    var dataFechamento = new Date();
    var dataFimFormatada = Utilities.formatDate(dataFechamento, ss.getSpreadsheetTimeZone(), "dd/MM/yyyy HH:mm:ss");
    var dataAbertura = abaOS.getRange(linhaOS, 2).getValue(); 

    abaOS.getRange(linhaOS, 3).setValue("FECHADO"); 
    abaOS.getRange(linhaOS, 29).setValue("'" + dataFimFormatada); 

    if (dataAbertura instanceof Date) {
      var diffMs = dataFechamento.getTime() - dataAbertura.getTime();
      var diffHrs = Math.floor(diffMs / 3600000);
      var diffMins = Math.floor((diffMs % 3600000) / 60000);
      abaOS.getRange(linhaOS, 30).setValue(diffHrs + "h " + diffMins + "m"); 
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
    return { sucesso: true, mensagem: "🔒 O.S. encerrada com sucesso! Tempo de reparo consolidado." };
  } catch (e) { return { sucesso: false, mensagem: "Erro ao fechar O.S.: " + e.toString() }; }
}

function listarOrdensServicoFrontend() {
  try {
    var abaOS = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("ORDEM_SERVICO_FROTA");
    if (!abaOS) return [];
    var ultimaLinha = abaOS.getLastRow();
    if (ultimaLinha <= 1) return [];

    var dados = abaOS.getRange(2, 1, ultimaLinha - 1, 30).getValues();
    var lista = [];
    var limiteInicio = Math.max(0, dados.length - 1000);

    for (var i = limiteInicio; i < dados.length; i++) {
      var row = dados[i];
      if (!row[0]) continue;
      var dataStr = row[1];
      if (dataStr instanceof Date) dataStr = ("0"+dataStr.getDate()).slice(-2)+"/"+("0"+(dataStr.getMonth()+1)).slice(-2)+" "+("0"+dataStr.getHours()).slice(-2)+":"+("0"+dataStr.getMinutes()).slice(-2);

      lista.push({ linhaOrigem: i + 2, id: row[0], dataInicial: dataStr, status: row[2], prioridade: row[3], prefixo: row[7], placa: row[8], tipoOS: row[11], setor: row[12], custoTotal: row[23] || 0, tempoTotal: row[29] || "" });
    }
    return lista.reverse();
  } catch (e) { return []; }
}
// ================= MÓDULO DE EDIÇÃO DE O.S. =================

function obterDetalhesOSBackend(linha) {
  try {
    var abaOS = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("ORDEM_SERVICO_FROTA");
    var row = abaOS.getRange(linha, 1, 1, 30).getValues()[0];

    // Reconstrói as listas de produtos e serviços
    var prodsNome = String(row[15]||"").split('\n');
    var prodsQtd = String(row[16]||"").split('\n');
    var prodsVal = String(row[17]||"").split('\n');
    var listaProdutos = [];
    for(var i=0; i<prodsNome.length; i++){
       if(prodsNome[i].trim() !== "") {
          var v = prodsVal[i] ? parseFloat(prodsVal[i].replace("R$", "").trim()) : 0;
          listaProdutos.push({nome: prodsNome[i], qtd: parseFloat(prodsQtd[i])||0, valor: v});
       }
    }

    var servsNome = String(row[19]||"").split('\n');
    var servsQtd = String(row[20]||"").split('\n');
    var servsVal = String(row[21]||"").split('\n');
    var listaServicos = [];
    for(var i=0; i<servsNome.length; i++){
       if(servsNome[i].trim() !== "") {
          var v = servsVal[i] ? parseFloat(servsVal[i].replace("R$", "").trim()) : 0;
          listaServicos.push({nome: servsNome[i], qtd: parseFloat(servsQtd[i])||0, valor: v});
       }
    }

    var dataIni = row[1];
    var dataFormatada = "";
    if (dataIni instanceof Date) {
       var tzOffset = dataIni.getTimezoneOffset() * 60000;
       dataFormatada = (new Date(dataIni - tzOffset)).toISOString().slice(0,16);
    } else { dataFormatada = dataIni; }

    return {
      sucesso: true,
      dados: {
        id: row[0], dataInicial: dataFormatada, 
        status: row[2], // <--- CORREÇÃO: AGORA ELE LÊ E ENVIA O STATUS!
        prioridade: row[3], tipoManutencao: row[4],
        situacaoParada: row[5], solicitante: row[6], prefixo: row[7], placa: row[8],
        km: row[10], tipoOS: row[11], setor: row[12], fornecedor: row[13] === "-" ? "" : row[13],
        defeito: row[14], produtos: listaProdutos, servicos: listaServicos, diagnostico: row[24], mecanico: row[25],assinaturaMecanico: row[26] || "", dataFechamento: row[28] ? row[28].toString().replace("'", "") : "", // Envia a data de fecho se existir
        tempoTotal: row[29] || "" // Envia o Downtime se existir
      }
    };
  } catch(e) { return { sucesso: false, mensagem: e.toString() }; }
}

function atualizarOrdemServicoBackend(linha, dados) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var abaOS = ss.getSheetByName("ORDEM_SERVICO_FROTA");

    var rowData = abaOS.getRange(linha, 1, 1, 30).getValues()[0];

    rowData[1] = dados.dataInicial; 
    rowData[3] = dados.prioridade.toUpperCase();
    rowData[4] = dados.tipoManutencao.toUpperCase();
    rowData[5] = dados.situacaoParada.toUpperCase();
    rowData[6] = dados.solicitante.toUpperCase();
    rowData[7] = dados.prefixo.toUpperCase();
    rowData[8] = dados.placa;
    rowData[10] = dados.km;
    rowData[11] = dados.tipoOS.toUpperCase();
    rowData[12] = dados.setor.toUpperCase();
    rowData[13] = dados.tipoOS.toUpperCase() === "EXTERNO" ? dados.fornecedor.toUpperCase() : "-";
    rowData[14] = dados.defeito; // <--- CORREÇÃO AQUI (Estava a puxar "dados.descricao")

    var strProdutos = dados.produtos.map(p => p.nome).join('\n');
    var strQtdProd = dados.produtos.map(p => p.qtd).join('\n');
    var strValProd = dados.produtos.map(p => "R$ " + p.valor.toFixed(2)).join('\n');
    var totalProd = dados.produtos.reduce((acc, p) => acc + (p.qtd * p.valor), 0);

    var strServicos = dados.servicos.map(s => s.nome).join('\n');
    var strQtdServ = dados.servicos.map(s => s.qtd).join('\n');
    var strValServ = dados.servicos.map(s => "R$ " + s.valor.toFixed(2)).join('\n');
    var totalServ = dados.servicos.reduce((acc, s) => acc + (s.qtd * s.valor), 0);

    rowData[15] = strProdutos; rowData[16] = strQtdProd; rowData[17] = strValProd; rowData[18] = totalProd;
    rowData[19] = strServicos; rowData[20] = strQtdServ; rowData[21] = strValServ; rowData[22] = totalServ;
    rowData[23] = totalProd + totalServ;
    rowData[24] = dados.diagnostico;
    rowData[25] = dados.mecanico.toUpperCase();
    if(dados.linkAssinatura) rowData[26] = dados.linkAssinatura;

    abaOS.getRange(linha, 1, 1, 30).setValues([rowData]);

    return { sucesso: true, mensagem: "✏️ O.S. atualizada e recalculada com sucesso!" };
  } catch(e) { return { sucesso: false, mensagem: e.toString() }; }
}

// ================= MÓDULO DE EMISSÃO DE DOCUMENTO PDF =================

// ================= MÓDULO DE EMISSÃO DE DOCUMENTO PDF =================
function gerarPdfOrdemServicoBackend(linha) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var abaOS = ss.getSheetByName("ORDEM_SERVICO_FROTA");
    if (!abaOS) return { sucesso: false, mensagem: "Aba de Ordens de Serviço não localizada." };

    var row = abaOS.getRange(linha, 1, 1, 30).getValues()[0];
    
    var idOs = row[0];
    var status = row[2];
    var prioridade = row[3];
    var tipoManutencao = row[4];
    var situacaoParada = row[5];
    var solicitante = row[6];
    var prefixo = row[7];
    var placa = row[8];
    var centroCusto = row[9];
    var km = parseFloat(row[10] || 0).toLocaleString('pt-BR');
    var tipoOS = row[11];
    var setor = row[12];
    var fornecedor = row[11] === "EXTERNO" ? row[13] : "OFICINA PRÓPRIA INOVAR";
    var defeito = row[14] || "Nenhum relato informado.";
    
    var diagnostico = row[24] || "Aguardando parecer técnico final.";
    var mecanico = row[25] || "-";
    
    // 🔥 VARIÁVEL DECLARADA AQUI (Evita o erro "not defined") 🔥
    var urlAssinatura = row[26] || ""; 
    var assinaturaBase64 = converterImagemParaBase64(urlAssinatura);
    
    var dataFechamento = row[28] ? row[28].toString().replace("'", "") : "-";
    var downtime = row[29] || "-";
    var custoTotal = parseFloat(row[23] || 0).toFixed(2);

    // Formata a data inicial de abertura
    var dataIniStr = row[1];
    if (dataIniStr instanceof Date) {
      dataIniStr = ("0"+dataIniStr.getDate()).slice(-2)+"/"+("0"+(dataIniStr.getMonth()+1)).slice(-2)+"/"+dataIniStr.getFullYear()+" "+("0"+dataIniStr.getHours()).slice(-2)+":"+("0"+dataIniStr.getMinutes()).slice(-2);
    }

    // Processamento da tabela de Peças/Produtos do PDF
    var prodsNome = String(row[15]||"").split('\n');
    var prodsQtd = String(row[16]||"").split('\n');
    var prodsVal = String(row[17]||"").split('\n');
    var tabelaProdutosHtml = "";
    
    for(var i=0; i<prodsNome.length; i++) {
      if(prodsNome[i].trim() !== "") {
        tabelaProdutosHtml += `<tr>
          <td style="padding:6px; border:1px solid #cbd5e1;">${prodsNome[i]}</td>
          <td style="padding:6px; border:1px solid #cbd5e1; text-align:center;">${prodsQtd[i]||1}</td>
          <td style="padding:6px; border:1px solid #cbd5e1; text-align:right;">${prodsVal[i]||"R$ 0,00"}</td>
        </tr>`;
      }
    }
    if(tabelaProdutosHtml === "") tabelaProdutosHtml = '<tr><td colspan="3" style="padding:8px; text-align:center; color:#64748b; font-style:italic; border:1px solid #e2e8f0;">Nenhum produto cadastrado nesta manutenção.</td></tr>';

    // Processamento da tabela de Serviços do PDF
    var servsNome = String(row[19]||"").split('\n');
    var servsQtd = String(row[20]||"").split('\n');
    var servsVal = String(row[21]||"").split('\n');
    var tabelaServicosHtml = "";
    
    for(var i=0; i<servsNome.length; i++) {
      if(servsNome[i].trim() !== "") {
        tabelaServicosHtml += `<tr>
          <td style="padding:6px; border:1px solid #cbd5e1;">${servsNome[i]}</td>
          <td style="padding:6px; border:1px solid #cbd5e1; text-align:center;">${servsQtd[i]||1}</td>
          <td style="padding:6px; border:1px solid #cbd5e1; text-align:right;">${servsVal[i]||"R$ 0,00"}</td>
        </tr>`;
      }
    }
    if(tabelaServicosHtml === "") tabelaServicosHtml = '<tr><td colspan="3" style="padding:8px; text-align:center; color:#64748b; font-style:italic; border:1px solid #e2e8f0;">Nenhum serviço registrado para esta manutenção.</td></tr>';

    // CONSTRUÇÃO E ESTILIZAÇÃO DO DOCUMENTO GRÁFICO EM HTML
    var htmlTemplate = `
      <html>
        <head>
          <style>
            body { font-family: 'Arial', sans-serif; color: #1e293b; padding: 10px; margin: 0; line-height: 1.3; }
            .header-table { width: 100%; border-collapse: collapse; margin-bottom: 20px; }
            .titulo-doc { font-size: 20px; font-weight: bold; color: #1e3a8a; margin: 0; }
            .secao-titulo { background: #1e3a8a; color: white; padding: 6px 10px; font-size: 12px; font-weight: bold; text-transform: uppercase; margin-top: 15px; margin-bottom: 8px; border-radius: 3px; }
            .grid-dados { width: 100%; border-collapse: collapse; font-size: 11px; margin-bottom: 5px; }
            .grid-dados td { padding: 5px 8px; border: 1px solid #e2e8f0; vertical-align: top; }
            .label-ficha { font-weight: bold; color: #475569; }
            .caixa-texto { font-size: 11px; background: #f8fafc; border: 1px solid #cbd5e1; padding: 8px; border-radius: 4px; min-height: 35px; margin-bottom: 5px; white-space: pre-wrap; }
            .tabela-itens { width: 100%; border-collapse: collapse; font-size: 10px; margin-top: 5px; }
            .tabela-itens th { background: #f1f5f9; padding: 6px; border: 1px solid #cbd5e1; font-weight: bold; text-align: left; color: #334155; }
            .total-box { text-align: right; font-size: 14px; font-weight: bold; color: #b91c1c; padding: 10px; background: #fef2f2; border: 1px solid #fee2e2; border-radius: 4px; margin-top: 15px; }
          </style>
        </head>
        <body>
          
          <table class="header-table">
            <tr>
              <td style="width: 70%;">
                <span style="font-size: 11px; font-weight: bold; color: #64748b; letter-spacing: 1px;">INOVAR SOLUÇÕES AMBIENTAIS</span>
                <p class="titulo-doc">Ficha de Ordem de Serviço</p>
                <span style="font-size: 11px; color:#475569;">Sistema Integrado de Gestão e Monitoramento de Manutenção da Frota</span>
              </td>
              <td style="width: 30%; text-align: right; vertical-align: middle;">
                <div style="background: #f1f5f9; border: 2px solid #cbd5e1; padding: 8px 12px; border-radius: 6px; display: inline-block; text-align: center;">
                  <span style="font-size: 10px; font-weight: bold; color: #475569; display: block;">CÓDIGO DA O.S.</span>
                  <span style="font-size: 18px; font-weight: bold; color: #1e3a8a;">${idOs}</span>
                </div>
              </td>
            </tr>
          </table>

          <div class="secao-titulo">1. Identificação do Veículo e Solicitante</div>
          <table class="grid-dados">
            <tr>
              <td style="width: 25%;"><span class="label-ficha">Prefixo Frota:</span><br><span style="font-size:12px; font-weight:bold; color:#1e3a8a;">${prefixo}</span></td>
              <td style="width: 25%;"><span class="label-ficha">Placa Veículo:</span><br><strong>${placa}</strong></td>
              <td style="width: 25%;"><span class="label-ficha">KM / Horímetro:</span><br>${km} KM</td>
              <td style="width: 25%;"><span class="label-ficha">Centro de Custo:</span><br>${centroCusto}</td>
            </tr>
            <tr>
              <td><span class="label-ficha">Solicitante:</span><br>${solicitante}</td>
              <td><span class="label-ficha">Nível Prioridade:</span><br><strong>${prioridade}</strong></td>
              <td><span class="label-ficha">Tipo Manutenção:</span><br>${tipoManutencao}</td>
              <td><span class="label-ficha">Situação da Parada:</span><br>${situacaoParada}</td>
            </tr>
          </table>

          <div class="secao-titulo">2. Informações de Execução Oficina</div>
          <table class="grid-dados">
            <tr>
              <td style="width: 33%;"><span class="label-ficha">Tipo de O.S.:</span><br>${tipoOS}</td>
              <td style="width: 33%;"><span class="label-ficha">Setor de Reparo:</span><br>${setor}</td>
              <td style="width: 34%;"><span class="label-ficha">Fornecedor Responsável:</span><br><strong>${fornecedor}</strong></td>
            </tr>
            <tr>
              <td><span class="label-ficha">Data/Hora Abertura:</span><br>${dataIniStr}</td>
              <td><span class="label-ficha">Data/Hora Fechamento:</span><br>${dataFechamento}</td>
              <td><span class="label-ficha">Tempo de Máquina Parada:</span><br><strong style="color:#b91c1c;">${downtime}</strong></td>
            </tr>
          </table>

          <div class="secao-titulo">3. Defeito / Relato do Problema Constatado</div>
          <div class="caixa-texto">${defeito}</div>

          <div class="secao-titulo">4. Produtos, Insumos e Peças Aplicadas</div>
          <table class="tabela-itens">
            <thead>
              <tr>
                <th style="width: 60%;">Descrição do Item / Peça Substituída</th>
                <th style="width: 15%; text-align: center;">Quantidade</th>
                <th style="width: 25%; text-align: right;">Valor Unitário</th>
              </tr>
            </thead>
            <tbody>
              ${tabelaProdutosHtml}
            </tbody>
          </table>

          <div class="secao-titulo">5. Serviços de Mão de Obra Prestados</div>
          <table class="tabela-itens">
            <thead>
              <tr>
                <th style="width: 60%;">Descrição da Atividade Mecânica / Elétrica</th>
                <th style="width: 15%; text-align: center;">Qtd Horas</th>
                <th style="width: 25%; text-align: right;">Valor Unitário</th>
              </tr>
            </thead>
            <tbody>
              ${tabelaServicosHtml}
            </tbody>
          </table>

          <div class="secao-titulo">6. Diagnóstico e Parecer Técnico Final</div>
          <div style="border:1px solid #cbd5e1; padding:12px; font-size:11px; margin-bottom:30px; border-radius:4px; min-height: 30px;">
            ${diagnostico ? diagnostico.toUpperCase() : 'AGUARDANDO PARECER TÉCNICO FINAL.'}
          </div>

          <table style="width:100%; border:none; margin-top:10px;">
            <tr>
              <td style="font-size:11px; vertical-align: bottom;">
                <strong>Mecânico/Responsável:</strong> ${mecanico}
                
                <div style="height: 65px; display: flex; align-items: flex-end; margin-top: 5px; margin-bottom: 5px;">
                   ${assinaturaBase64 ? `<img src="${assinaturaBase64}" style="max-height: 60px; max-width: 340px; object-fit: contain;">` : '<div style="height: 40px;"></div>'}
                </div>
                
                <div style="border-top:1px solid #000; width:350px; padding-top:5px; font-size:9px;">Assinatura Digital / Carimbo Eletrônico Vincular Técnico</div>
              </td>
              <td style="text-align:right; vertical-align: bottom;">
                 <div style="border:1px solid #fecaca; background-color:#fef2f2; padding:10px 20px; display:inline-block; text-align:center; min-width: 150px;">
                    <div style="font-size:9px; color:#b91c1c; font-weight:bold; margin-bottom:3px;">VALOR CONSOLIDADO DA O.S.</div>
                    <div style="font-size:18px; font-weight:bold; color:#b91c1c;">R$ ${custoTotal}</div>
                 </div>
              </td>
            </tr>
          </table>

        </body>
      </html>
    `;

    // Conversão do HTML bruto para o motor PDF do Google Apps Script
    var blobPdf = HtmlService.createHtmlOutput(htmlTemplate).getAs('application/pdf');
    var bytes = blobPdf.getBytes();
    var base64Convertido = Utilities.base64Encode(bytes);

    return { sucesso: true, base64: base64Convertido, filename: "ORDEM_SERVICO_" + idOs + ".pdf" };
  } catch (e) {
    return { sucesso: false, mensagem: "Falha interna no motor de PDF: " + e.toString() };
  }
}
// ================= GERADOR DE LOTE DE PDF MULTI-PÁGINAS (ORDEM DE SERVIÇO) =================
// ================= GERADOR DE LOTE DE PDF MULTI-PÁGINAS (ORDEM DE SERVIÇO) =================
function gerarLotePdfOSBackend(linhas) {
  try {
    var htmlLote = "<html><head><style>body { font-family: Arial, sans-serif; color: #1e293b; }</style></head><body>";
    
    for (var i = 0; i < linhas.length; i++) {
      var linha = linhas[i];
      var resDetalhes = obterDetalhesOSBackend(linha); 
      
      if (!resDetalhes || !resDetalhes.sucesso) continue;
      
      var os = resDetalhes.dados;
      var linhasProdutos = "";
      var linhasServicos = "";
      var custoTotalCalc = 0;
      
      // 4. PRODUTOS E PEÇAS
      if (os.produtos && os.produtos.length > 0) {
         os.produtos.forEach(function(p) {
           var tot = p.qtd * p.valor;
           custoTotalCalc += tot;
           linhasProdutos += `<tr>
             <td style="border-bottom:1px solid #e2e8f0; padding:8px; font-size:10px;">${p.nome.toUpperCase()}</td>
             <td style="border-bottom:1px solid #e2e8f0; padding:8px; text-align:center; font-size:10px;">${p.qtd}</td>
             <td style="border-bottom:1px solid #e2e8f0; padding:8px; text-align:right; font-size:10px;">R$ ${parseFloat(p.valor).toFixed(2)}</td>
             <td style="border-bottom:1px solid #e2e8f0; padding:8px; text-align:right; font-size:10px;">R$ ${tot.toFixed(2)}</td>
           </tr>`;
         });
      } else {
         linhasProdutos = '<tr><td colspan="4" style="border-bottom:1px solid #e2e8f0; padding:12px; text-align:center; color:#64748b; font-style:italic; font-size:10px;">Nenhum produto cadastrado nesta manutenção.</td></tr>';
      }
      
      // 5. SERVIÇOS E MÃO DE OBRA
      if (os.servicos && os.servicos.length > 0) {
         os.servicos.forEach(function(s) {
           var tot = s.qtd * s.valor;
           custoTotalCalc += tot;
           linhasServicos += `<tr>
             <td style="border-bottom:1px solid #e2e8f0; padding:8px; font-size:10px;">${s.nome.toUpperCase()}</td>
             <td style="border-bottom:1px solid #e2e8f0; padding:8px; text-align:center; font-size:10px;">${s.qtd}</td>
             <td style="border-bottom:1px solid #e2e8f0; padding:8px; text-align:right; font-size:10px;">R$ ${parseFloat(s.valor).toFixed(2)}</td>
             <td style="border-bottom:1px solid #e2e8f0; padding:8px; text-align:right; font-size:10px;">R$ ${tot.toFixed(2)}</td>
           </tr>`;
         });
      } else {
         linhasServicos = '<tr><td colspan="4" style="border-bottom:1px solid #e2e8f0; padding:12px; text-align:center; color:#64748b; font-style:italic; font-size:10px;">Nenhum serviço registrado para esta manutenção.</td></tr>';
      }
      
      var idOSFormatado = String(os.id).toUpperCase().replace("OS-", "");
      var fornecedorFinal = os.fornecedor ? os.fornecedor.toUpperCase() : "OFICINA PRÓPRIA INOVAR";
      
      // 🔥 CONVERSÃO DA ASSINATURA EXCLUSIVA AQUI 🔥
      var assinaturaBase64 = converterImagemParaBase64(os.assinaturaMecanico);

      var htmlPagina = `
        <div style="padding: 10px; font-family: Arial, sans-serif; color: #1e293b;">
          
          <table style="width:100%; border:none; margin-bottom:20px;">
            <tr>
              <td>
                <div style="font-size:12px; font-weight:bold; color:#1e3a8a; margin-bottom:2px;">INOVAR SOLUÇÕES AMBIENTAIS</div>
                <div style="font-size:22px; font-weight:bold; color:#1e3a8a; background-color:#dbeafe; display:inline-block; padding:2px 6px; margin-bottom:4px;">Ficha de Ordem de Serviço</div>
                <div style="font-size:11px; color:#64748b;">Sistema Integrado de Gestão e Monitoramento de Manutenção da Frota</div>
              </td>
              <td style="text-align:right; vertical-align:middle;">
                <div style="border:2px solid #cbd5e1; border-radius:6px; padding:8px 15px; display:inline-block; text-align:center; min-width: 100px;">
                  <div style="font-size:10px; color:#1e3a8a; margin-bottom:2px; font-weight:bold;">CÓDIGO DA O.S.</div>
                  <div style="font-size:22px; font-weight:bold; color:#1e3a8a;">OS-${idOSFormatado}</div>
                </div>
              </td>
            </tr>
          </table>

          <h3 style="font-size:12px; color:#94a3b8; border-bottom:none; margin-bottom:5px; text-transform:uppercase;">1. IDENTIFICAÇÃO DO VEÍCULO E SOLICITANTE</h3>
          <table style="width:100%; border-collapse:collapse; font-size:11px; margin-bottom:20px;">
            <tr>
              <td style="border:1px solid #cbd5e1; padding:8px; width:25%;"><strong>Prefixo Frota:</strong><br><span style="font-size:14px; color:#1e3a8a; font-weight:bold;">${os.prefixo}</span></td>
              <td style="border:1px solid #cbd5e1; padding:8px; width:25%;"><strong>Placa Veículo:</strong><br>${os.placa}</td>
              <td style="border:1px solid #cbd5e1; padding:8px; width:25%;"><strong>KM / Horímetro:</strong><br>${os.km} KM</td>
              <td style="border:1px solid #cbd5e1; padding:8px; width:25%;"><strong>Centro de Custo:</strong><br>${os.centroCusto || 'A DEFINIR'}</td>
            </tr>
            <tr>
              <td style="border:1px solid #cbd5e1; padding:8px;"><strong>Solicitante:</strong><br>${os.solicitante}</td>
              <td style="border:1px solid #cbd5e1; padding:8px;"><strong>Nível Prioridade:</strong><br><strong>${os.prioridade}</strong></td>
              <td style="border:1px solid #cbd5e1; padding:8px;"><strong>Tipo Manutenção:</strong><br>${os.tipoManutencao}</td>
              <td style="border:1px solid #cbd5e1; padding:8px;"><strong>Situação da Parada:</strong><br>${os.situacaoParada}</td>
            </tr>
          </table>

          <h3 style="font-size:12px; color:#94a3b8; margin-bottom:5px; text-transform:uppercase;">2. INFORMAÇÕES DE EXECUÇÃO OFICINA</h3>
          <table style="width:100%; border-collapse:collapse; font-size:11px; margin-bottom:20px;">
            <tr>
              <td style="border:1px solid #cbd5e1; padding:8px; width:33%;"><strong>Tipo de O.S.:</strong><br>${os.tipoOS}</td>
              <td style="border:1px solid #cbd5e1; padding:8px; width:33%;"><strong>Setor de Reparo:</strong><br>${os.setor}</td>
              <td style="border:1px solid #cbd5e1; padding:8px; width:34%;"><strong>Fornecedor Responsável:</strong><br>${fornecedorFinal}</td>
            </tr>
            <tr>
              <td style="border:1px solid #cbd5e1; padding:8px;"><strong>Data/Hora Abertura:</strong><br>${os.dataInicial || '-'}</td>
              <td style="border:1px solid #cbd5e1; padding:8px;"><strong>Data/Hora Fechamento:</strong><br>${os.dataFechamento || '-'}</td>
              <td style="border:1px solid #cbd5e1; padding:8px;"><strong>Tempo de Parada:</strong><br><span style="color:#b91c1c; font-weight:bold;">${os.tempoTotal || '-'}</span></td>
            </tr>
          </table>

          <h3 style="font-size:12px; color:#94a3b8; margin-bottom:5px; text-transform:uppercase;">3. DEFEITO / RELATO DO PROBLEMA CONSTATADO</h3>
          <div style="border:1px solid #cbd5e1; padding:12px; font-size:11px; margin-bottom:20px; border-radius:4px; min-height: 30px;">
            ${os.defeito ? os.defeito.toUpperCase() : '-'}
          </div>

          <h3 style="font-size:12px; color:#94a3b8; margin-bottom:5px; text-transform:uppercase;">4. PRODUTOS, INSUMOS E PEÇAS APLICADAS</h3>
          <table style="width:100%; border-collapse:collapse; font-size:11px; margin-bottom:20px;">
            <thead>
              <tr>
                <th style="border-top:1px solid #cbd5e1; border-bottom:1px solid #cbd5e1; padding:8px; text-align:left; color:#1e3a8a;">Descrição do Item / Peça Substituída</th>
                <th style="border-top:1px solid #cbd5e1; border-bottom:1px solid #cbd5e1; padding:8px; text-align:center; color:#1e3a8a; width:15%;">Quantidade</th>
                <th style="border-top:1px solid #cbd5e1; border-bottom:1px solid #cbd5e1; padding:8px; text-align:right; color:#1e3a8a; width:20%;">Valor Unitário</th>
                <th style="border-top:1px solid #cbd5e1; border-bottom:1px solid #cbd5e1; padding:8px; text-align:right; color:#1e3a8a; width:20%;">Valor Total</th>
              </tr>
            </thead>
            <tbody>
              ${linhasProdutos}
            </tbody>
          </table>

          <h3 style="font-size:12px; color:#94a3b8; margin-bottom:5px; text-transform:uppercase;">5. SERVIÇOS DE MÃO DE OBRA PRESTADOS</h3>
          <table style="width:100%; border-collapse:collapse; font-size:11px; margin-bottom:20px;">
            <thead>
              <tr>
                <th style="border-top:1px solid #cbd5e1; border-bottom:1px solid #cbd5e1; padding:8px; text-align:left; color:#1e3a8a;">Descrição da Atividade Mecânica / Elétrica</th>
                <th style="border-top:1px solid #cbd5e1; border-bottom:1px solid #cbd5e1; padding:8px; text-align:center; color:#1e3a8a; width:15%;">Qtd Horas</th>
                <th style="border-top:1px solid #cbd5e1; border-bottom:1px solid #cbd5e1; padding:8px; text-align:right; color:#1e3a8a; width:20%;">Valor Unitário</th>
                <th style="border-top:1px solid #cbd5e1; border-bottom:1px solid #cbd5e1; padding:8px; text-align:right; color:#1e3a8a; width:20%;">Valor Total</th>
              </tr>
            </thead>
            <tbody>
              ${linhasServicos}
            </tbody>
          </table>

          <h3 style="font-size:12px; color:#94a3b8; margin-bottom:5px; text-transform:uppercase;">6. DIAGNÓSTICO E PARECER TÉCNICO FINAL</h3>
          <div style="border:1px solid #cbd5e1; padding:12px; font-size:11px; margin-bottom:30px; border-radius:4px; min-height: 30px;">
            ${os.diagnostico ? os.diagnostico.toUpperCase() : 'AGUARDANDO PARECER TÉCNICO FINAL.'}
          </div>

          <table style="width:100%; border:none; margin-top:10px;">
            <tr>
              <td style="font-size:11px; vertical-align: bottom;">
                <strong>Mecânico/Responsável:</strong> ${os.mecanico ? os.mecanico.toUpperCase() : 'A DEFINIR'}
                
                <div style="height: 65px; display: flex; align-items: flex-end; margin-top: 5px; margin-bottom: 5px;">
                   ${assinaturaBase64 ? `<img src="${assinaturaBase64}" style="max-height: 60px; max-width: 340px; object-fit: contain;">` : '<div style="height: 40px;"></div>'}
                </div>
                
                <div style="border-top:1px solid #000; width:350px; padding-top:5px; font-size:9px;">Assinatura Digital / Carimbo Eletrônico Vinculado Técnico</div>
              </td>
              <td style="text-align:right; vertical-align: bottom;">
                 <div style="border:1px solid #fecaca; background-color:#fef2f2; padding:10px 20px; display:inline-block; text-align:center; min-width: 150px;">
                    <div style="font-size:9px; color:#b91c1c; font-weight:bold; margin-bottom:3px;">VALOR CONSOLIDADO DA O.S.</div>
                    <div style="font-size:18px; font-weight:bold; color:#b91c1c;">R$ ${custoTotalCalc.toFixed(2)}</div>
                 </div>
              </td>
            </tr>
          </table>
        </div>
      `;
      
      htmlLote += htmlPagina;
      
      if (i < linhas.length - 1) {
        htmlLote += '<div style="page-break-after: always;"></div>';
      }
    }
    
    htmlLote += "</body></html>";
    
    var blobPdf = Utilities.newBlob(htmlLote, MimeType.HTML).getAs(MimeType.PDF);
    blobPdf.setName("Relatorio_OS_Lote.pdf");
    
    var base64 = Utilities.base64Encode(blobPdf.getBytes());
    return { sucesso: true, base64: base64 };

  } catch (e) {
    return { sucesso: false, mensagem: e.toString() };
  }
}
// ================= CONVERSOR DE IMAGENS PARA O PDF =================
function converterImagemParaBase64(url) {
  if (!url || url.trim() === "") return "";
  try {
    // Verifica se é um link do Google Drive e extrai o ID
    var match = url.match(/id=([a-zA-Z0-9_-]+)/) || url.match(/\/d\/([a-zA-Z0-9_-]+)/);
    if (match && match[1]) {
      var id = match[1];
      var blob = DriveApp.getFileById(id).getBlob();
      return "data:" + blob.getContentType() + ";base64," + Utilities.base64Encode(blob.getBytes());
    } else {
      // Se for imagem externa comum
      var res = UrlFetchApp.fetch(url);
      return "data:" + res.getBlob().getContentType() + ";base64," + Utilities.base64Encode(res.getBlob().getBytes());
    }
  } catch(e) {
    return ""; // Se der erro na imagem, retorna vazio para não travar a emissão do PDF
  }
}
