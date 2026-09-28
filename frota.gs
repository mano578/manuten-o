// ==============================================================
// GESTÃO DE FROTA (EDITAR E EXCLUIR)
// ==============================================================

// NOTA (revisão 7): esta função comparava `dados[i][0]` (Coluna A = "Data Cadastro") com o
// prefixo recebido — mas o prefixo do veículo está na Coluna C (índice 2), como confirma
// `listarVeiculosCadastrados` mais abaixo (`prefixo: String(dados[i][2])`). Por isso a busca
// NUNCA encontrava o veículo (comparava uma data com um prefixo) e sempre retornava "Veículo não
// encontrado na planilha.", mesmo com o veículo existindo. Corrigido para comparar a coluna certa.
function excluirVeiculoBackendPorPrefixo(prefixo) {
  try {
    var aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("CADASTRO_VEICULOS"); // <-- Mude se a sua aba tiver outro nome
    var dados = aba.getDataRange().getValues();
    
    for(var i = 1; i < dados.length; i++) {
       if(String(dados[i][2]).trim().toUpperCase() === String(prefixo).trim().toUpperCase()) {
          aba.deleteRow(i + 1);
          return { sucesso: true, mensagem: "🗑️ Veículo excluído com sucesso!" };
       }
    }
    return { sucesso: false, mensagem: "❌ Veículo não encontrado na planilha." };
  } catch(e) { return { sucesso: false, mensagem: "Erro: " + e.message }; }
}

// NOTA (revisão 7 — CORREÇÃO DO ERRO "Veículo não encontrado na planilha" AO EDITAR):
// a causa era esta linha comparar `sheetDados[i][0]` (Coluna A = "Data Cadastro", um valor de
// data) com o prefixo digitado (ex.: "UBC7C21") — nunca dava match, porque o prefixo real do
// veículo está na Coluna C (índice 2), exatamente como `listarVeiculosCadastrados` já lê mais
// abaixo (`prefixo: String(dados[i][2])`). Por isso TODA edição de veículo falhava com esse erro,
// mesmo o veículo existindo na planilha. Corrigido para comparar a coluna certa (índice 2).
//
// NOTA (revisão 7 — NOVO CAMPO "Status"): agora também grava `dados.status` (ATIVO/DEVOLVIDO/
// EMPRESTADO) na coluna 26 (Z), quando informado pelo front-end — ver `listarVeiculosCadastrados`
// e a nova função `atualizarStatusVeiculoBackend` para o detalhe de como essa coluna é lida.
function atualizarVeiculoBackendPorPrefixo(prefixoAntigo, dados) {
  try {
    // NOTA (revisão): a aba real de cadastro de veículos é "CADASTRO_VEICULOS" (usada em
    // excluirVeiculoBackendPorPrefixo acima e em várias outras partes do sistema). O nome "FROTA"
    // usado aqui originalmente não existe na planilha — esta função sempre lançava erro
    // ("Cannot read properties of null") ao tentar editar um veículo. Corrigido para usar a aba real.
    var aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("CADASTRO_VEICULOS") || SpreadsheetApp.getActiveSpreadsheet().getSheetByName("FROTA");
    var sheetDados = aba.getDataRange().getValues();
    
    for(var i = 1; i < sheetDados.length; i++) {
       // Procura a linha onde está o veículo original (Coluna C = índice 2 = Prefixo)
       if(String(sheetDados[i][2]).trim().toUpperCase() === String(prefixoAntigo).trim().toUpperCase()) {
          var linha = i + 1;
          
          // =========================================================================
          // ⚠️ ATENÇÃO INGRID: AJUSTE OS NÚMEROS DAS COLUNAS CONFORME A SUA PLANILHA!
          // Ex: getRange(linha, 1) significa que vai atualizar a Coluna A.
          // Ex: getRange(linha, 2) significa que vai atualizar a Coluna B.
          // =========================================================================
          
          aba.getRange(linha, 3).setValue(String(dados.prefixo).toUpperCase());
          aba.getRange(linha, 4).setValue(String(dados.placa).toUpperCase());
          aba.getRange(linha, 7).setValue(String(dados.marcaModelo).toUpperCase());
          aba.getRange(linha, 6).setValue(String(dados.ano).toUpperCase());
          // CORREÇÃO (revisão 8): esta linha gravava `dados.categoria` incondicionalmente — se o
          // campo chegasse vazio do front-end (ex.: <select> sem opção correspondente ao valor
          // salvo do veículo), a Categoria era apagada da planilha mesmo que o usuário só quisesse
          // mudar o Status. Agora só sobrescreve quando o front-end de fato enviou um valor.
          if (dados.categoria) aba.getRange(linha, 2).setValue(String(dados.categoria).toUpperCase());
          aba.getRange(linha, 8).setValue(String(dados.chassi).toUpperCase());
          aba.getRange(linha, 9).setValue(String(dados.renavam).toUpperCase());
          aba.getRange(linha, 10).setValue(String(dados.motor).toUpperCase());
          // Coluna 26 (Z) = Status do veículo (ATIVO/DEVOLVIDO/EMPRESTADO) — só grava se o
          // front-end mandou o campo, para não sobrescrever com vazio chamadas antigas.
          if (dados.status) aba.getRange(linha, 26).setValue(String(dados.status).toUpperCase());

          // Se a sua planilha não tiver alguma destas colunas, basta apagar a linha acima correspondente.

          return { sucesso: true, mensagem: "✅ Veículo atualizado com sucesso!" };
       }
    }
    return { sucesso: false, mensagem: "❌ Veículo não encontrado na planilha." };
  } catch(e) { return { sucesso: false, mensagem: "Erro: " + e.message }; }
}

// NOVA FUNÇÃO (adicionada nesta revisão, não existia antes): muda só o status do veículo
// (ATIVO, DEVOLVIDO ou EMPRESTADO) sem precisar abrir o formulário completo de edição. Reaproveita
// a mesma coluna 26 (Z) que `atualizarVeiculoBackendPorPrefixo` já grava.
function atualizarStatusVeiculoBackend(prefixo, novoStatus) {
  try {
    var statusValido = ["ATIVO", "DEVOLVIDO", "EMPRESTADO"];
    var statusNormalizado = String(novoStatus || "").trim().toUpperCase();
    if (statusValido.indexOf(statusNormalizado) === -1) {
      return { sucesso: false, mensagem: "❌ Status inválido. Use ATIVO, DEVOLVIDO ou EMPRESTADO." };
    }
    var aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("CADASTRO_VEICULOS");
    var dados = aba.getDataRange().getValues();
    for (var i = 1; i < dados.length; i++) {
      if (String(dados[i][2]).trim().toUpperCase() === String(prefixo).trim().toUpperCase()) {
        aba.getRange(i + 1, 26).setValue(statusNormalizado);
        return { sucesso: true, mensagem: "✅ Status atualizado para " + statusNormalizado + "!", status: statusNormalizado };
      }
    }
    return { sucesso: false, mensagem: "❌ Veículo não encontrado na planilha." };
  } catch (e) { return { sucesso: false, mensagem: "Erro: " + e.message }; }
}
function obterResumoDaFrota() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var planilha = ss.getSheetByName("MANUTENÇÃO");
    if (!planilha) return [];
    var valores = planilha.getDataRange().getValues();

    // 🆕 MAPA DE DEFEITOS (O.S. ABERTAS) — usado para mostrar o balão/tooltip no Painel Geral
    var mapaDefeitosOS = {};
    var abaOS = ss.getSheetByName("ORDEM_SERVICO_FROTA");
    if (abaOS && abaOS.getLastRow() > 1) {
      var dadosOS = abaOS.getRange(2, 1, abaOS.getLastRow() - 1, 15).getValues();
      for (var o = 0; o < dadosOS.length; o++) {
        var statusOS = String(dadosOS[o][2] || "").trim().toUpperCase();
        var prefixoOS = String(dadosOS[o][7] || "").trim().toUpperCase();
        if (prefixoOS && statusOS === "ABERTO") {
          // Mantém sempre o relato mais recente daquele prefixo (a planilha cresce para baixo)
          mapaDefeitosOS[prefixoOS] = String(dadosOS[o][14] || "").trim();
        }
      }
    }

    // 🆕 MAPA DE DESCRIÇÕES SOS (CHAMADOS EM ABERTO) — usado para mostrar o balão/tooltip no Painel Geral
    var mapaDescricoesSOS = {};
    var abaSOS = ss.getSheetByName("SOS_FROTA");
    if (abaSOS && abaSOS.getLastRow() > 1) {
      var dadosSOS = abaSOS.getRange(2, 1, abaSOS.getLastRow() - 1, 6).getValues();
      for (var s = 0; s < dadosSOS.length; s++) {
        var statusSOS = String(dadosSOS[s][5] || "EM ABERTO").trim().toUpperCase();
        var prefixoSOS = String(dadosSOS[s][1] || "").trim().toUpperCase();
        if (prefixoSOS && statusSOS === "EM ABERTO") {
          mapaDescricoesSOS[prefixoSOS] = String(dadosSOS[s][4] || "").trim();
        }
      }
    }

    var resumo = {
      "COLETA DOMICILIAR": { nome: "Coleta Domiciliar", liberados: [], retidosOS: [], retidosSOS: [] },
      "ENTULHO MECANIZADO": { nome: "Entulho Mecanizado", liberados: [], retidosOS: [], retidosSOS: [] },
      "ENTULHO MANUAL": { nome: "Entulho Manual", liberados: [], retidosOS: [], retidosSOS: [] },
      "LIMPEZA MECANIZADA": { nome: "Limpeza Mecânizada", liberados: [], retidosOS: [], retidosSOS: [] },
      "MOTO FISCAL": { nome: "Moto Fiscal", liberados: [], retidosOS: [], retidosSOS: [] },
      "PINTURA MECÂNIZADA ": { nome: "Pintura Mecânizada", liberados: [], retidosOS: [], retidosSOS: [] },
      "PODA MANUAL": { nome: "Poda Manual", liberados: [], retidosOS: [], retidosSOS: [] },
      "PODA TRITURADA": { nome: "Poda Triturada", liberados: [], retidosOS: [], retidosSOS: [] },
      "REMOÇÃO MOTORIZADA": { nome: "Remoção Motorizada", liberados: [], retidosOS: [], retidosSOS: [] },
      "ROÇAGEM MECANIZADA": { nome: "Roçagem Mecânizada", liberados: [], retidosOS: [], retidosSOS: [] },
      "TRANSPORTE DE EQUIPES": { nome: "Transporte de Equipes", liberados: [], retidosOS: [], retidosSOS: [] },
      "VARRIÇÃO MECANIZADA": { nome: "Varrição Mecânizada", liberados: [], retidosOS: [], retidosSOS: [] },
      "ADMINISTRAÇÃO": { nome: "Administração", liberados: [], retidosOS: [], retidosSOS: [] },
      "PINTURA MECÂNIZADA": { nome: "Pintura Mecânizada", liberados: [], retidosOS: [], retidosSOS: [] }, 
    };

    for (var i = 1; i < valores.length; i++) {
      var prefixo = String(valores[i][1] || "").trim();
      var placa = String(valores[i][2] || "").trim();
      var status = String(valores[i][3] || "").trim().toUpperCase();
      var centroCusto = String(valores[i][37] || "").trim().toUpperCase();

      if (!prefixo) continue;
      var prefixoChave = prefixo.toUpperCase();

      if (resumo[centroCusto]) {
        // 🔥 FILTRO ESTREITO: Só entra se bater EXATAMENTE com a palavra!
        if (status === "LIBERADO") {
           resumo[centroCusto].liberados.push({ prefixo: prefixo, placa: placa });
        }
        else if (status === "RETIDO - O.S." || status === "RETIDO") {
           // 🆕 Anexa o defeito/relato da O.S. aberta para exibir no balão/tooltip do Painel
           resumo[centroCusto].retidosOS.push({ prefixo: prefixo, placa: placa, defeito: mapaDefeitosOS[prefixoChave] || "" });
        }
        else if (status === "S.O.S" || status === "SOS" || status === "S.O.S (MANUTENÇÃO RÁPIDA)") {
           // 🆕 Anexa a descrição do chamado SOS em aberto para exibir no balão/tooltip do Painel
           resumo[centroCusto].retidosSOS.push({ prefixo: prefixo, placa: placa, descricao: mapaDescricoesSOS[prefixoChave] || "" });
        }
      }
    }
    return [resumo["COLETA DOMICILIAR"], resumo["ENTULHO MECANIZADO"], resumo["ENTULHO MANUAL"], resumo["LIMPEZA MECANIZADA"], resumo["MOTO FISCAL"], resumo["PINTURA MECÂNIZADA"], resumo["PODA MANUAL"], resumo["PODA TRITURADA"], resumo["REMOÇÃO MOTORIZADA"], resumo["ROÇAGEM MECANIZADA"], resumo["TRANSPORTE DE EQUIPES"], resumo["VARRIÇÃO MECANIZADAL"], resumo["ADMINISTRAÇÃO"], resumo["PINTURA MECÂNIZADA"]];
  } catch(e) { return []; }
}
// ================= 2. LOCALIZADOR INTELIGENTE DE COLUNAS =================
// Esta função impede que o sistema quebre se você mudar as colunas de lugar nas planilhas
function encontrarColuna(dadosAba, nomeBuscado1, nomeBuscado2) {
  if(!dadosAba || dadosAba.length === 0) return -1;
  var cabecalho = dadosAba[0];
  for(var i=0; i<cabecalho.length; i++) {
    var col = String(cabecalho[i]).toUpperCase();
    if(col.includes(nomeBuscado1) || (nomeBuscado2 && col.includes(nomeBuscado2))) return i;
  }
  return -1;
}


// NOTA (revisão): analisarPendenciasVeiculoBackend e efetivarLiberacaoVeiculoBackend
// estavam duplicadas neste arquivo e em Painel_Backend.gs. Mantidas apenas as versões
// de Painel_Backend.gs, que usam getDisplayValues()/SpreadsheetApp.flush() e são mais
// robustas contra erros de coluna. As definições aqui foram removidas para eliminar a
// sobrescrita silenciosa de função duplicada.

function salvarVeiculoForm(formulario) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet(); var abaCadastros = ss.getSheetByName("CADASTRO_VEICULOS");
    var cabecalho = ["Data Cadastro", "Categoria", "Prefixo", "Placa", "Ano Fab.", "Ano Mod.", "Marca/Modelo", "Chassi", "Renavam", "Motor", "Eixos", "Rodado", "Tipo Equip.", "Capacidade", "Valor Aluguel", "Fornecedor", "Início de Operação", "Link Foto Dianteira", "Link Foto Lat. Esquerda", "Link Foto Lat. Direita", "Link Foto Traseira", "Link Tacógrafo", "Link Documentos", "Centro de Custo", "Informação"];
    
    if (!abaCadastros) { 
      abaCadastros = ss.insertSheet("CADASTRO_VEICULOS"); 
      abaCadastros.appendRow(cabecalho); 
      abaCadastros.getRange("A1:Y1").setFontWeight("bold").setBackground("#d9ead3"); 
    }
    var pastas = DriveApp.getFoldersByName("FROTA_ANEXOS"); var pastaDestino = pastas.hasNext() ? pastas.next() : DriveApp.createFolder("FROTA_ANEXOS");

    function salvarArquivo(blob, nomeBase) {
      if (blob && typeof blob.getName === 'function' && blob.getName() !== "") {
        var arquivo = pastaDestino.createFile(blob); arquivo.setName(formulario.placa + "_" + nomeBase + "_" + blob.getName());
        arquivo.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); return arquivo.getUrl();
      }
      return "";
    }

    var dataInicio = formulario.inicioOperacao ? formatarDataInputHtml(formulario.inicioOperacao).split(" ")[0] : "";
    var novaLinha = [ 
      new Date(), formulario.categoria, formulario.prefixo, formulario.placa, 
      formulario.anoFabricacao, formulario.anoModelo, formulario.marcaModelo, 
      formulario.chassi, formulario.renavam, formulario.nrMotor, formulario.qtdEixo, 
      formulario.tipoRodado, formulario.tipoEquipamento || "", formulario.capacidadeEquipamento || "", 
      formulario.valorAluguel || "", formulario.fornecedor || "", dataInicio, 
      salvarArquivo(formulario.fileDianteira, "Dianteira"), salvarArquivo(formulario.fileLateralEsquerda, "Lat_Esq"), 
      salvarArquivo(formulario.fileLateralDireita, "Lat_Dir"), salvarArquivo(formulario.fileTraseira, "Traseira"), 
      salvarArquivo(formulario.fileTacografo, "Tacografo"), salvarArquivo(formulario.fileDocumentos, "Documentos"),
      formulario.centroCusto || "", formulario.informacao || ""
    ];
    abaCadastros.appendRow(novaLinha);

    // 🔥 CHAMA O ROBÔ PARA SINCRONIZAR COM A MANUTENÇÃO 🔥
    sincronizarVeiculoManutencao(formulario.prefixo, formulario.placa, formulario.centroCusto || "");

    return { sucesso: true, mensagem: "Veículo cadastrado com sucesso!" };
  } catch (erro) { return { sucesso: false, mensagem: "Erro ao cadastrar: " + erro.toString() }; }
}

function listarVeiculosCadastrados() {
  var abaCadastros = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("CADASTRO_VEICULOS");
  if (!abaCadastros) return [];
  var dados = abaCadastros.getDataRange().getValues();
  var veiculos = [];
  for (var i = 1; i < dados.length; i++) {
    veiculos.push({
      originalIndex: i - 1, prefixo: String(dados[i][2]), placa: String(dados[i][3]), marcaModelo: String(dados[i][6]),
      ano: String(dados[i][4]) + "/" + String(dados[i][5]), categoria: String(dados[i][1]), chassi: String(dados[i][7]),
      renavam: String(dados[i][8]), motor: String(dados[i][9]), eixos: String(dados[i][10]), rodado: String(dados[i][11]),
      tipoEquipamento: String(dados[i][12]), capacidade: String(dados[i][13]), fornecedor: String(dados[i][15]),
      inicioOperacao: dados[i][16] ? Utilities.formatDate(new Date(dados[i][16]), Session.getScriptTimeZone(), "dd/MM/yyyy") : "-",
      linkDianteira: String(dados[i][17]), linkLatEsq: String(dados[i][18]), linkLatDir: String(dados[i][19]),
      linkTraseira: String(dados[i][20]), linkTacografo: String(dados[i][21]), linkDocs: String(dados[i][22]),
      centroCusto: String(dados[i][23] || ""), informacao: String(dados[i][24] || ""),
      // Coluna 26 (Z) = índice 25 = Status do veículo (ATIVO/DEVOLVIDO/EMPRESTADO), gravado por
      // atualizarVeiculoBackendPorPrefixo/atualizarStatusVeiculoBackend. Sem valor ainda -> ATIVO.
      status: String(dados[i][25] || "ATIVO").toUpperCase()
    });
  }
  return veiculos.reverse(); 
}
// ================= GERENCIADOR DE ANEXOS DA FROTA (NOVO) =================
function atualizarAnexoVeiculo(indexOriginal, placa, tipoAnexo, base64Data, mimeType, fileName) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var aba = ss.getSheetByName("CADASTRO_VEICULOS");
    
    // Mapeamento exato das colunas de anexos na sua planilha CADASTRO_VEICULOS
    var columnMap = {
      'Dianteira': 18,
      'Lat_Esq': 19,
      'Lat_Dir': 20,
      'Traseira': 21,
      'Tacografo': 22,
      'Documentos': 23
    };
    
    var col = columnMap[tipoAnexo];
    if(!col) throw "Tipo de anexo inválido para atualização.";

    var pastas = DriveApp.getFoldersByName("FROTA_ANEXOS");
    var pastaDestino = pastas.hasNext() ? pastas.next() : DriveApp.createFolder("FROTA_ANEXOS");

    // Decodifica e salva o novo arquivo no Drive
    var decoded = Utilities.base64Decode(base64Data);
    var nomeArquivo = placa + "_ATUALIZADO_" + tipoAnexo + "_" + new Date().getTime();
    var blob = Utilities.newBlob(decoded, mimeType, nomeArquivo);
    var arquivo = pastaDestino.createFile(blob);
    arquivo.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    var urlNova = arquivo.getUrl();
    
    // Atualiza a célula exata na linha do veículo
    aba.getRange(indexOriginal + 2, col).setValue(urlNova); 

    return { sucesso: true, mensagem: "Anexo do veículo atualizado com sucesso!", url: urlNova };
  } catch (e) {
    return { sucesso: false, mensagem: "Erro ao atualizar anexo: " + e.toString() };
  }
}
// ================= FUNÇÃO PARA EXCLUIR ANEXO DO VEÍCULO =================
function excluirAnexoVeiculo(indexOriginal, tipoAnexo) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var aba = ss.getSheetByName("CADASTRO_VEICULOS");
    
    // Mapeamento das colunas (Igual ao de atualização)
    var columnMap = {
      'Dianteira': 18,
      'Lat_Esq': 19,
      'Lat_Dir': 20,
      'Traseira': 21,
      'Tacografo': 22,
      'Documentos': 23
    };
    
    var col = columnMap[tipoAnexo];
    if(!col) throw "Tipo de anexo inválido.";

    // Limpa a célula correspondente ao anexo apagando o link
    aba.getRange(indexOriginal + 2, col).setValue(""); 

    return { sucesso: true, mensagem: "Anexo excluído com sucesso!" };
  } catch (e) {
    return { sucesso: false, mensagem: "Erro ao excluir anexo: " + e.toString() };
  }
}
// ================= ROBÔ DE SINCRONIZAÇÃO AUTOMÁTICA =================
// Cria a linha na aba MANUTENÇÃO automaticamente quando um veículo é cadastrado
function sincronizarVeiculoManutencao(prefixo, placa, centroCusto) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var abaManu = ss.getSheetByName("MANUTENÇÃO");

    if (!abaManu) return; // Se a aba não existir, aborta silenciosamente

    var dados = abaManu.getDataRange().getValues();
    var novoId = 1;

    // 1. Descobre o último ID para continuar a sequência matemática
    if (dados.length > 1) {
      var maxId = 0;
      for (var i = 1; i < dados.length; i++) {
        var idAtual = parseInt(dados[i][0], 10);
        if (!isNaN(idAtual) && idAtual > maxId) {
          maxId = idAtual;
        }
      }
      novoId = maxId + 1;
    }

    // 2. Prepara uma linha em branco até à coluna AL (que é a coluna 38)
    var novaLinhaArray = new Array(38).fill("");

    // 3. Preenche as colunas cirurgicamente (lembrando que Array começa no 0)
    novaLinhaArray[0] = novoId;                  // Coluna A (ID)
    novaLinhaArray[1] = prefixo.toUpperCase();   // Coluna B (PREFIXO)
    novaLinhaArray[2] = placa.toUpperCase();     // Coluna C (PLACA)
    novaLinhaArray[3] = "LIBERADO";              // Coluna D (SITUAÇÃO)
    
    // As colunas E até AK ficam vazias
    
    novaLinhaArray[37] = centroCusto.toUpperCase(); // Coluna AL (CENTRO DE CUSTO)

    // 4. Injeta a linha inteira de uma só vez no fundo da planilha
    abaManu.appendRow(novaLinhaArray);

  } catch (e) {
    Logger.log("Erro ao sincronizar veículo na Manutenção: " + e.message);
  }
}
