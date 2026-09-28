// =========================================================================
// MOTOR CENTRAL DE RELATÓRIOS DA FROTA (VERSÃO COM BLINDAGEM DE DATAS)
// =========================================================================

function gerarRelatoriosFrotaBackend(filtros) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var modulos = filtros.modulos || [];
    var dataIniStr = filtros.inicio;
    var dataFimStr = filtros.fim;
    
    // 1. BLINDAGEM DE FUSO HORÁRIO (Converte o texto do input HTML para Data Real)
    var dtInicio = new Date(2000, 0, 1);
    if (dataIniStr) {
       var pIni = dataIniStr.split('-');
       dtInicio = new Date(pIni[0], parseInt(pIni[1])-1, pIni[2], 0, 0, 0);
    }
    
    var dtFim = new Date(2099, 11, 31);
    if (dataFimStr) {
       var pFim = dataFimStr.split('-');
       dtFim = new Date(pFim[0], parseInt(pFim[1])-1, pFim[2], 23, 59, 59);
    }

    // 2. TRADUTOR DE DATAS DO GOOGLE SHEETS (Resolve o problema do DD/MM vs MM/DD)
    function lerDataBrasileira(valorPlanilha) {
       if (!valorPlanilha) return null;
       if (valorPlanilha instanceof Date) return valorPlanilha; // Já é data
       
       var texto = String(valorPlanilha).trim().split(" ")[0]; // Pega só a data
       if (texto.includes("/")) {
          var partes = texto.split("/");
          if(partes.length === 3) {
             // Força o formato Ano, Mês, Dia
             return new Date(partes[2], parseInt(partes[1])-1, partes[0], 12, 0, 0);
          }
       }
       var dFallback = new Date(valorPlanilha);
       return isNaN(dFallback.getTime()) ? null : dFallback;
    }

    // 3. BUSCADOR DE VALOR (Garante que acha o valor na linha, ignorando vazios no fim)
    function lerUltimoValorMonetario(linha) {
       for (var c = linha.length - 1; c >= 0; c--) {
          var v = parseFloat(linha[c]);
          if (!isNaN(v) && v > 0) return v;
       }
       return 0;
    }

    var resposta = {};

    // ----------------------------------------------------
    // RELATÓRIO 1: CUSTOS / RATEIO (MAPEAMENTO CORRIGIDO)
    // ----------------------------------------------------
    if (modulos.includes("CUSTOS")) {
      // Puxa exatamente o nome da aba que aparece na sua imagem
      var abaCustos = ss.getSheetByName("CUSTOS_FROTA");
      var listaCustos = [];
      
      if (abaCustos) {
        var dados = abaCustos.getDataRange().getValues();
        
        for (var i = 1; i < dados.length; i++) {
           // A data está na Coluna F (Índice 5 no JavaScript)
           var dataLinha = lerDataBrasileira(dados[i][5]); 
           
           if (dataLinha && dataLinha >= dtInicio && dataLinha <= dtFim) {
              
              // Garante que o valor financeiro (Coluna H / Índice 7) é lido corretamente, mesmo que tenha vírgula
              var valorBruto = dados[i][7];
              var valorLimpo = typeof valorBruto === 'string' ? parseFloat(valorBruto.replace(',', '.')) : parseFloat(valorBruto);
              if (isNaN(valorLimpo)) valorLimpo = 0;

              listaCustos.push({
                data: Utilities.formatDate(dataLinha, Session.getScriptTimeZone(), "dd/MM/yyyy"),
                nf: dados[i][0] || "-", // Usamos o ID da Coluna A como código da NF
                prefixo: dados[i][1] || "S/N", // Coluna B
                fornecedor: dados[i][3] || "-", // Coluna D
                peca: dados[i][4] || "-", // Coluna E
                valorTotal: valorLimpo // Coluna H
              });
           }
        }
      }
      resposta.CUSTOS = listaCustos;
    }

    // ----------------------------------------------------
    // RELATÓRIO 2: ABASTECIMENTO (COM PLACA INCLUÍDA)
    // ----------------------------------------------------
    if (modulos.includes("ABASTECIMENTOS")) {
      var abaAbs = ss.getSheetByName("ABASTECIMENTO_FROTA");
      var listaAbs = [];
      
      if (abaAbs) {
        var dados = abaAbs.getDataRange().getValues();
        for (var i = 1; i < dados.length; i++) {
           var dataLinha = lerDataBrasileira(dados[i][3]); 
           
           if (dataLinha && dataLinha >= dtInicio && dataLinha <= dtFim) {
              
              var valorBruto = dados[i][7];
              var valorLimpo = typeof valorBruto === 'string' ? parseFloat(valorBruto.replace(',', '.')) : parseFloat(valorBruto);
              if (isNaN(valorLimpo)) valorLimpo = 0;

              var litrosBruto = dados[i][5];
              var litrosLimpo = typeof litrosBruto === 'string' ? parseFloat(litrosBruto.replace(',', '.')) : parseFloat(litrosBruto);
              if (isNaN(litrosLimpo)) litrosLimpo = 0;

              var kmBruto = dados[i][9]; 
              var kmLimpo = typeof kmBruto === 'string' ? parseFloat(kmBruto.replace(',', '.')) : parseFloat(kmBruto);
              if (isNaN(kmLimpo)) kmLimpo = 0;

              listaAbs.push({
                data: Utilities.formatDate(dataLinha, Session.getScriptTimeZone(), "dd/MM/yyyy"),
                prefixo: dados[i][1] || "S/N", 
                placa: dados[i][2] || "", // <-- ADICIONADO A PLACA (COLUNA C)
                combustivel: dados[i][4] || "N/A", 
                litros: litrosLimpo,
                kmPercorrido: kmLimpo,
                valorTotal: valorLimpo
              });
           }
        }
      }
      resposta.ABASTECIMENTOS = listaAbs;
    }
   // ----------------------------------------------------
    // RELATÓRIO 3: RADAR DE MANUTENÇÃO (HÍBRIDO: ALERTAS + HISTÓRICO POR DATA DA BAIXA)
    // ----------------------------------------------------
    if (modulos.includes("RADAR")) {
       var listaRadar = [];
       var statusFiltro = filtros.statusRadar || "ATIVOS"; // ATIVOS, HISTORICO ou TODOS

       // 1. Puxa os alertas pendentes/atrasados do motor principal (Sempre visíveis no painel ativo)
       if (statusFiltro === "ATIVOS" || statusFiltro === "TODOS") {
          var dadosRadarGeral = listarRadarManutencaoBackend();
          if (dadosRadarGeral && dadosRadarGeral.sucesso) {
             dadosRadarGeral.dados.forEach(function(alerta) {
                if (!alerta.isConcluido) {
                   listaRadar.push(alerta);
                }
             });
          }
       }

       // 2. Puxa as manutenções CONCLUÍDAS filtrando estritamente pela DATA DA BAIXA (Coluna P)
       if (statusFiltro === "HISTORICO" || statusFiltro === "TODOS") {
          var abaPlano = ss.getSheetByName("PLANO_MANUTENCAO");
          if (abaPlano) {
             var dadosPlano = abaPlano.getDataRange().getValues();
             var dadosExibicao = abaPlano.getDataRange().getDisplayValues();

             for (var i = 1; i < dadosPlano.length; i++) {
                // Lê a Coluna O (Índice 14) para checar a conclusão
                var concluidoVal = String(dadosPlano[i][14] || "").trim().toUpperCase();
                
                if (concluidoVal === "SIM" || concluidoVal === "S" || concluidoVal === "CONCLUIDO") {
                   
                   // 🔥 ALTERAÇÃO: Puxa a data da Coluna P (Índice 15) em vez da Coluna C
                   var dataBaixa = lerDataBrasileira(dadosPlano[i][15]); 
                   
                   // Se o filtro de período estiver ativo, valida contra a DATA DA BAIXA
                   if (dataBaixa && (dataBaixa < dtInicio || dataBaixa > dtFim)) {
                      continue;
                   }
                   
                   // Se a célula da baixa estiver vazia mas houver um filtro de data ativo na tela, ignora o registro
                   if (!dataBaixa && filtros.inicio && filtros.fim) {
                      continue;
                   }
                   
                   var dataBaixaStr = dadosExibicao[i][15] || "-";
                   var dataRealizStr = dadosExibicao[i][2] || "-";
                   var medicao = String(dadosPlano[i][6] || "").toUpperCase();
                   
                   listaRadar.push({
                      prefixo: dadosPlano[i][3] || "S/N",
                      placa: dadosPlano[i][4] || "",
                      categoria: dadosPlano[i][1] || "GERAL",
                      descricao: dadosPlano[i][11] || "-",
                      dataRealizada: dataRealizStr, // Mantém a visualização da data original do serviço
                      kmRealizado: dadosPlano[i][7] || 0,
                      medicao: medicao,
                      infoDetalhe: "Concluído em: " + dataBaixaStr, // 🔥 Exibe o dia exato em que foi dada a baixa
                      isConcluido: true,
                      status: "CONCLUIDO",
                      alertaTxt: "CONCLUÍDO"
                   });
                }
             }
          }
       }

       resposta.RADAR = listaRadar;
    }

    // ----------------------------------------------------
    // RELATÓRIO 4: FROTA CADASTRADA (COM FILTRO DE CENTRO DE CUSTO)
    // ----------------------------------------------------
    if (modulos.includes("FROTA")) {
      var abaFrota = ss.getSheetByName("CADASTRO_VEICULOS");
      var listaFrota = [];
      
      if (abaFrota) {
        var dados = abaFrota.getDataRange().getValues();
        var cabecalhos = dados[0].map(function(c) { return String(c).trim().toUpperCase(); });
        
        var idxPrefixo = cabecalhos.indexOf("PREFIXO") !== -1 ? cabecalhos.indexOf("PREFIXO") : 2; 
        var idxPlaca = cabecalhos.indexOf("PLACA") !== -1 ? cabecalhos.indexOf("PLACA") : 3; 
        var idxMarca = cabecalhos.indexOf("MARCA/MODELO") !== -1 ? cabecalhos.indexOf("MARCA/MODELO") : 6; 
        var idxTipo = cabecalhos.indexOf("TIPO EQUIP.") !== -1 ? cabecalhos.indexOf("TIPO EQUIP.") : 12; 
        var idxFornecedor = cabecalhos.indexOf("FORNECEDOR") !== -1 ? cabecalhos.indexOf("FORNECEDOR") : 15; 
        var idxCentroCusto = cabecalhos.indexOf("CENTRO DE CUSTO") !== -1 ? cabecalhos.indexOf("CENTRO DE CUSTO") : 23; 

        var ccFiltro = filtros.centroCusto || "TODOS";

        for (var i = 1; i < dados.length; i++) {
           if (String(dados[i][idxPrefixo] || "").trim() === "") continue; 
           
           var ccLido = String(dados[i][idxCentroCusto] || "").trim().toUpperCase();
           var ccTratado = ccLido !== "" ? ccLido : "NÃO DEFINIDO";

           // 🔥 FILTRAGEM CIRÚRGICA: Se não for "TODOS", ignora os outros Centros de Custo
           if (ccFiltro !== "TODOS" && ccTratado !== ccFiltro) {
              continue;
           }
           
           listaFrota.push({
              prefixo: String(dados[i][idxPrefixo] || "").toUpperCase(),
              placa: String(dados[i][idxPlaca] || "").toUpperCase(),
              marcaModelo: String(dados[i][idxMarca] || "-"),
              tipoEquipamento: String(dados[i][idxTipo] || "-"),
              fornecedor: String(dados[i][idxFornecedor] || "-"),
              centroCusto: ccTratado
           });
        }
      }
      resposta.FROTA = listaFrota;
    }

    // Devolve o pacote pronto para o Front-End
    return { sucesso: true, dados: resposta };

  } catch(e) {
    return { sucesso: false, mensagem: "Falha ao gerar o motor de dados: " + e.message + " (Linha " + e.lineNumber + ")" };
  }
}
