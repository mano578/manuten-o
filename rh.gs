function listarFuncionarios() {
  var aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("REF FUNCIONÁRIO"); if (!aba) return [];
  var dados = aba.getDataRange().getValues(); var lista = [];
  for (var i = 1; i < dados.length; i++) {
    var nome = String(dados[i][1]).trim(); if (!nome) continue;
    lista.push({
      originalIndex: i, nome: nome, cpf: String(dados[i][2]),
      admissao: dados[i][3] ? Utilities.formatDate(new Date(dados[i][3]), Session.getScriptTimeZone(), "dd/MM/yyyy") : "-",
      funcao: String(dados[i][4]), matricula: String(dados[i][5]), equipe: String(dados[i][7]),
      status: String(dados[i][10]) || "ATIVO", assinatura: String(dados[i][11])
    });
  }
  return lista;
}

function salvarFuncionarioForm(formulario) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet(); var aba = ss.getSheetByName("REF FUNCIONÁRIO");
    if (!aba) { aba = ss.insertSheet("REF FUNCIONÁRIO"); aba.appendRow(["ID", "NOME DO FUNCIONÁRIO", "CPF", "ADMISSÃO", "FUNÇÃO", "MATRICULA", "DATA DA FOLGA", "EQUIPE", "LICENÇA / ATESTADO / FERIAS", "NOME", "STATUS", "ASSINATURA"]); aba.getRange("A1:L1").setFontWeight("bold").setBackground("#e2e8f0"); }
    var pastas = DriveApp.getFoldersByName("SISTEMA_ASSINATURAS"); var pastaDestino = pastas.hasNext() ? pastas.next() : DriveApp.createFolder("SISTEMA_ASSINATURAS");
    var linkAssinatura = "";
    if (formulario.fileAssinatura && formulario.fileAssinatura.getName() !== "") {
      var arquivo = pastaDestino.createFile(formulario.fileAssinatura); arquivo.setName("ASSINATURA_" + formulario.nome.replace(/\s+/g, '_').toUpperCase());
      arquivo.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); linkAssinatura = arquivo.getUrl();
    }
    aba.appendRow([ aba.getLastRow(), formulario.nome.toUpperCase(), formulario.cpf, formulario.admissao ? formatarDataInputHtml(formulario.admissao).split(" ")[0] : "", formulario.funcao.toUpperCase(), formulario.matricula, "", formulario.equipe.toUpperCase(), "", formulario.nome.split(" ")[0].toUpperCase(), formulario.status.toUpperCase(), linkAssinatura ]);
    return { sucesso: true, mensagem: "Funcionário cadastrado!" };
  } catch (erro) { return { sucesso: false, mensagem: "Erro: " + erro.toString() }; }
}

function obterHistoricoFuncionario(nomeFuncionario) {
  var ss = SpreadsheetApp.getActiveSpreadsheet(); var historico = { escalas: [], atestados: [], advertencias: [], suspensoes: [] };
  var nomeBuscado = String(nomeFuncionario).trim().toUpperCase();
  function buscarNaAba(nomeAba, indexNomeColuna, colunasDesejadas) {
    var aba = ss.getSheetByName(nomeAba); if (!aba) return [];
    var dados = aba.getDataRange().getDisplayValues(); var resultados = [];
    for (var i = 1; i < dados.length; i++) {
      if (String(dados[i][indexNomeColuna]).trim().toUpperCase() === nomeBuscado) {
        var item = {}; colunasDesejadas.forEach(function(col) { item[col.chave] = dados[i][col.idx]; }); resultados.push(item);
      }
    }
    return resultados.reverse();
  }
  historico.escalas = buscarNaAba("ESCALA TRABALHADO", 2, [ {chave: "data", idx: 1}, {chave: "tipo", idx: 4}, {chave: "dataFolga", idx: 5} ]);
  historico.atestados = buscarNaAba("ATESTADOS", 1, [ {chave: "data", idx: 3}, {chave: "motivo", idx: 2}, {chave: "dias", idx: 4}, {chave: "retorno", idx: 5}, {chave: "arquivo", idx: 6} ]);
  historico.advertencias = buscarNaAba("ADVERTÊNCIA", 2, [ {chave: "data", idx: 1}, {chave: "motivo", idx: 4}, {chave: "tipo", idx: 5}, {chave: "arquivo", idx: 7} ]);
  historico.suspensoes = buscarNaAba("SUSPENSÃO", 2, [ {chave: "data", idx: 1}, {chave: "motivo", idx: 4}, {chave: "tipo", idx: 5}, {chave: "arquivo", idx: 7} ]);
  return historico;
}

function atualizarFuncionario(indexOriginal, novaEquipe, novoStatus) {
  try {
    var aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("REF FUNCIONÁRIO");
    if (!aba) return { sucesso: false, mensagem: "Aba não encontrada." };
    aba.getRange(indexOriginal + 1, 8).setValue(String(novaEquipe).trim().toUpperCase());
    aba.getRange(indexOriginal + 1, 11).setValue(String(novoStatus).trim().toUpperCase());
    return { sucesso: true, mensagem: "Dados do funcionário atualizados na planilha!" };
  } catch (e) { 
    return { sucesso: false, mensagem: "Erro crítico: " + e.toString() }; 
  }
}

// NOVA FUNÇÃO: ATUALIZA APENAS A ASSINATURA DE UM FUNCIONÁRIO EXISTENTE
function atualizarAssinaturaNoBanco(indexOriginal, nomeFunc, base64Data, mimeType, fileName) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var aba = ss.getSheetByName("REF FUNCIONÁRIO");
    if (!aba) return { sucesso: false, mensagem: "Aba de Funcionários não encontrada." };
    
    var pastas = DriveApp.getFoldersByName("SISTEMA_ASSINATURAS"); 
    var pastaDestino = pastas.hasNext() ? pastas.next() : DriveApp.createFolder("SISTEMA_ASSINATURAS");
    var decoded = Utilities.base64Decode(base64Data);
    var blob = Utilities.newBlob(decoded, mimeType, "ASSINATURA_" + nomeFunc.replace(/\s+/g, '_').toUpperCase() + "_" + new Date().getTime());
    var arquivo = pastaDestino.createFile(blob);
    arquivo.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    var linkAssinatura = arquivo.getUrl();
    // Coluna 12 (L) guarda a assinatura
    aba.getRange(indexOriginal + 1, 12).setValue(linkAssinatura);
    
    return { sucesso: true, mensagem: "Assinatura do funcionário vinculada com sucesso!", url: linkAssinatura };
  } catch (e) { 
    return { sucesso: false, mensagem: "Erro ao gravar assinatura: " + e.toString() }; 
  }
}

function salvarOcorrenciaRH(dados) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var linkArquivo = "";
    
    if (dados.fileBase64 && dados.fileName) {
      var pastas = DriveApp.getFoldersByName("RH_ANEXOS_COMPROVANTES"); 
      var pastaDestino = pastas.hasNext() ? pastas.next() : DriveApp.createFolder("RH_ANEXOS_COMPROVANTES");
      var decoded = Utilities.base64Decode(dados.fileBase64);
      var blob = Utilities.newBlob(decoded, dados.fileMime, dados.nomeFunc.split(" ")[0] + "_" + dados.tipo + "_" + dados.fileName);
      var arquivo = pastaDestino.createFile(blob);
      arquivo.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      linkArquivo = arquivo.getUrl();
    }

    var dataStr = dados.data ? formatarDataInputHtml(dados.data).split(" ")[0] : Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy");

    if (dados.tipo === 'ATESTADO') {
      var abaAt = ss.getSheetByName("ATESTADOS");
      if(!abaAt) { abaAt = ss.insertSheet("ATESTADOS"); abaAt.appendRow(["ID", "FUNCIONÁRIO", "CID / MOTIVO", "DATA DO ATESTADO", "DIAS ATESTADO", "DATA DE RETORNO ATESTADO", "IMAGEM"]); }
      
      var dataRetornoStr = "";
      if(dados.data && dados.dias) {
        var dr = new Date(dados.data + 'T12:00:00');
        dr.setDate(dr.getDate() + parseInt(dados.dias));
        dataRetornoStr = Utilities.formatDate(dr, Session.getScriptTimeZone(), "dd/MM/yyyy");
      }
      abaAt.appendRow([abaAt.getLastRow(), dados.nomeFunc.toUpperCase(), dados.motivo.toUpperCase(), dataStr, dados.dias, dataRetornoStr, linkArquivo]);

    } else if (dados.tipo === 'ADVERTENCIA') {
      var abaAdv = ss.getSheetByName("ADVERTÊNCIA");
      if(!abaAdv) { abaAdv = ss.insertSheet("ADVERTÊNCIA"); abaAdv.appendRow(["ID", "DATA", "FUNCIONÁRIO", "CPF", "ATO", "TIPO DE ADVERTÊNCIA", "ADVERTÊNCIA APLICADA POR", "PDF"]); }
      abaAdv.appendRow([abaAdv.getLastRow(), dataStr, dados.nomeFunc.toUpperCase(), dados.cpfFunc, dados.motivo, dados.artigo, "SISTEMA INOVAR", linkArquivo]);

    } else if (dados.tipo === 'SUSPENSAO') {
      var abaSusp = ss.getSheetByName("SUSPENSÃO");
      if(!abaSusp) { abaSusp = ss.insertSheet("SUSPENSÃO"); abaSusp.appendRow(["ID", "DATA", "FUNCIONÁRIO", "CPF", "ATO", "TIPO DE ADVERTÊNCIA", "ADVERTÊNCIA APLICADA POR", "PDF"]); }
      abaSusp.appendRow([abaSusp.getLastRow(), dataStr, dados.nomeFunc.toUpperCase(), dados.cpfFunc, dados.motivo, dados.artigo, "SISTEMA INOVAR", linkArquivo]);
    }

    return { sucesso: true, mensagem: dados.tipo + " registado com sucesso!" };
  } catch (erro) {
    return { sucesso: false, mensagem: "Erro ao registar: " + erro.toString() };
  }
}

function salvarEscalaTrabalho(dadosEscala) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet(); var aba = ss.getSheetByName("ESCALA TRABALHADO");
    if (!aba) { aba = ss.insertSheet("ESCALA TRABALHADO"); aba.appendRow(["ID", "DATA", "FUNCIONÁRIO", "FUNÇÃO", "TIPO DE ESCALA", "DATA DA FOLGA", "MATRICULA", "PDF", "EQUIPE"]); aba.getRange("A1:I1").setFontWeight("bold").setBackground("#d9ead3"); }
    
    var dataEscala = formatarDataInputHtml(dadosEscala.data).split(" ")[0];
    var funcionarios = dadosEscala.funcionariosSelecionados;
    
    for (var i = 0; i < funcionarios.length; i++) {
      var f = funcionarios[i];
      var dataFolgaIndividual = f.folga ? formatarDataInputHtml(f.folga).split(" ")[0] : "";
      aba.appendRow([ aba.getLastRow(), dataEscala, f.nome, f.funcao, dadosEscala.tipoEscala.toUpperCase(), dataFolgaIndividual, f.matricula, "", f.equipe ]);
    }
    return { sucesso: true, mensagem: "Escala gerada! O PDF está pronto para impressão.", dataEscala: dataEscala, tipoEscala: dadosEscala.tipoEscala.toUpperCase() };
  } catch (erro) { return { sucesso: false, mensagem: "Erro ao gerar escala: " + erro.toString() }; }
}
// ================= MÓDULO: EXCLUIR ESCALA COMPLETA =================

function excluirEscalaBackend(dataAlvo, tipoAlvo) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var aba = ss.getSheetByName("ESCALA TRABALHADO");
    if (!aba) return { sucesso: false, mensagem: "Aba ESCALA TRABALHADO não encontrada." };

    var dados = aba.getDataRange().getValues();
    var excluiuAlgo = false;

    // Limpeza dos parâmetros de busca
    var dataBusca = String(dataAlvo).trim();
    var tipoBusca = String(tipoAlvo).trim().toUpperCase();

    // ⚠️ IMPORTANTE: A varredura TEM de ser de baixo para cima!
    // Se apagarmos de cima para baixo, os números das linhas mudam e o robô perde-se.
    for (var i = dados.length - 1; i > 0; i--) {
       var dataLinha = String(dados[i][1]).trim();
       var tipoLinha = String(dados[i][4]).trim().toUpperCase();

       // Caso o Excel tenha formatado a data como objeto de Calendário, convertemos para texto
       if (dados[i][1] instanceof Date) {
          var d = dados[i][1];
          var mes = ("0" + (d.getMonth() + 1)).slice(-2);
          var dia = ("0" + d.getDate()).slice(-2);
          dataLinha = d.getFullYear() + "-" + mes + "-" + dia; // Formato padrão HTML
          
          // Se a sua data estiver no formato BR, ele tenta também:
          var dataLinhaBR = dia + "/" + mes + "/" + d.getFullYear();
          if (dataBusca === dataLinhaBR) dataLinha = dataLinhaBR; 
       }

       // Se a Data E o Tipo baterem certo, a linha é eliminada!
       if ((dataLinha === dataBusca || dataLinha.includes(dataBusca)) && tipoLinha === tipoBusca) {
          aba.deleteRow(i + 1);
          excluiuAlgo = true;
       }
    }

    if (excluiuAlgo) {
      return { sucesso: true, mensagem: "✅ Escala de " + tipoBusca + " do dia " + dataBusca + " excluída com sucesso!" };
    } else {
      return { sucesso: false, mensagem: "⚠️ Nenhuma escala encontrada com esses dados para excluir." };
    }
  } catch (e) {
    return { sucesso: false, mensagem: "Erro interno ao excluir: " + e.toString() };
  }
}

function listarEscalasAgrupadas() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var abaEscala = ss.getSheetByName("ESCALA TRABALHADO"); var abaFunc = ss.getSheetByName("REF FUNCIONÁRIO");
  if (!abaEscala) return [];
  
  var mapDadosFunc = {};
  if (abaFunc) {
    var dadosFunc = abaFunc.getDataRange().getValues();
    for(var i=1; i<dadosFunc.length; i++) { 
      mapDadosFunc[String(dadosFunc[i][1]).trim().toUpperCase()] = {
        assinatura: String(dadosFunc[i][11]),
        matricula: String(dadosFunc[i][5])
      };
    }
  }

  var dadosEscala = abaEscala.getDataRange().getValues();
  var agrupado = {};
  for (var i = 1; i < dadosEscala.length; i++) {
    var rawData = dadosEscala[i][1];
    var dataStr = Object.prototype.toString.call(rawData) === '[object Date]' ? Utilities.formatDate(rawData, Session.getScriptTimeZone(), "dd/MM/yyyy") : String(rawData).trim();
    var tipo = String(dadosEscala[i][4]).trim();
    if(!dataStr || !tipo) continue;

    var chave = dataStr + "|" + tipo;
    if(!agrupado[chave]) agrupado[chave] = { data: dataStr, tipo: tipo, funcionarios: [] };
    
    var nome = String(dadosEscala[i][2]).trim();
    var dadosComplementares = mapDadosFunc[nome.toUpperCase()] || {assinatura: "", matricula: "-"};
    var folga = String(dadosEscala[i][5]).trim();
    if (Object.prototype.toString.call(dadosEscala[i][5]) === '[object Date]') folga = Utilities.formatDate(dadosEscala[i][5], Session.getScriptTimeZone(), "dd/MM/yyyy");

    agrupado[chave].funcionarios.push({ 
      nome: nome, 
      funcao: String(dadosEscala[i][3]), 
      folga: folga,
      matricula: dadosComplementares.matricula,
      assinatura: dadosComplementares.assinatura
    });
  }
  return Object.values(agrupado).reverse();
}
function listarOcorrenciasTotaisBackend(tipo) {
  try {
    var aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(tipo);
    if(!aba) return [];
    var dados = aba.getDataRange().getDisplayValues();
    var lista = [];
    
    // Pega as últimas 200 linhas de baixo para cima para ficar rápido
    for (var i = dados.length - 1; i >= 1; i--) {
       if (dados[i][0] && dados[i][0] !== "") {
          if (tipo === "ATESTADOS") {
              lista.push({ linha: i + 1, id: dados[i][0], funcionario: dados[i][1], motivo: dados[i][2], data: dados[i][3], dias: dados[i][4] });
          } else {
              lista.push({ linha: i + 1, id: dados[i][0], data: dados[i][1], funcionario: dados[i][2], ato: dados[i][4], pdf: dados[i][7] || "" });
          }
          if(lista.length >= 200) break;
       }
    }
    return lista; // Já vem do mais recente pro mais antigo
  } catch(e) { return []; }
}
