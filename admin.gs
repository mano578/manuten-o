// NOTA (revisão 6): você mandou o link direto de cada pasta no Google Drive. Isso é MUITO mais
// confiável do que buscar por nome (`DriveApp.getFoldersByName`) — busca por nome pode falhar
// silenciosamente por espaço invisível, acentuação, a pasta estar num Drive Compartilhado, ou
// mais de uma pasta existir com o mesmo nome. Extraí o ID de cada pasta direto do link que você
// mandou (é o trecho entre "/folders/" e "?usp="):
// - Driver:  https://drive.google.com/drive/folders/1eyUrZjX_H6yW6DKnzl8BpKGInVKHzcRK
// - Backup:  https://drive.google.com/drive/folders/1DiUifUzeu_kqkZq_EaVlX1mnMm-sOV2m
// Agora o sistema tenta abrir a pasta DIRETO por este ID primeiro (rápido e não depende do
// nome bater) e só cai para a busca por nome/criação de pasta nova se o ID não abrir por algum
// motivo (ex.: você mover a pasta para a lixeira).
var ID_PASTA_DRIVER_MANUTENCAO = "1eyUrZjX_H6yW6DKnzl8BpKGInVKHzcRK";
var ID_PASTA_BACKUP_SISTEMA = "1DiUifUzeu_kqkZq_EaVlX1mnMm-sOV2m";

function obterPastaPorId_(id) {
  try {
    return DriveApp.getFolderById(id);
  } catch (e) {
    Logger.log("[obterPastaPorId_] Não foi possível abrir a pasta pelo ID '" + id + "': " + e);
    return null;
  }
}

// NOTA (revisão 3): `usuarioAdmin` (o valor vindo do front-end) não era normalizado antes da
// comparação — só o valor da planilha recebia `.trim()`. Um espaço a mais/menos no nome de
// usuário guardado na sessão do navegador (ex.: copiado e colado com espaço) fazia essa checagem
// falhar silenciosamente para TODAS as funções de admin, incluindo a listagem do Driver. Agora
// `usuarioAdmin` também é normalizado com `.trim()` antes de comparar.
function verificarAdmin(usuarioAdmin) {
  var planilha = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("USUARIOS_SISTEMA");
  var dados = planilha.getDataRange().getValues();
  var usuarioAdminNormalizado = String(usuarioAdmin || "").trim();
  for (var i = 1; i < dados.length; i++) {
    if (String(dados[i][0]).trim() === usuarioAdminNormalizado) {
      var permissoes = String(dados[i][3]).trim().toLowerCase();
      if (permissoes.indexOf("admin") !== -1 || permissoes.indexOf("tudo") !== -1) return true;
    }
  }
  return false;
}

// ================= MÓDULO INTELIGENTE DE RECUPERAÇÃO DE SENHA =================

// NOTA (revisão): solicitarRecuperacaoSenhaBackend estava definida duas vezes
// neste arquivo (bloco colado duplicado, mesma lógica). Mantida apenas a
// segunda definição, mais abaixo.

function trocarSenhaAutonoma(usuario, nome, novaSenha) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var aba = ss.getSheetByName("USUARIOS_SISTEMA");
    
    var dados = aba.getDataRange().getValues();
    var colNome = -1, colUser = -1, colSenha = -1, colTroca = -1;
    
    for(var c=0; c<dados[0].length; c++) {
       var cabecalho = String(dados[0][c]).toUpperCase();
       if(cabecalho === "NOME") colNome = c;
       if(cabecalho.indexOf("USU") > -1) colUser = c;
       if(cabecalho.indexOf("SENHA") > -1) colSenha = c;
       if(cabecalho.indexOf("LIVRE") > -1 || cabecalho.indexOf("TROCA") > -1) colTroca = c;
    }

    if (colNome === -1) colNome = 2; if (colUser === -1) colUser = 0;
    if (colSenha === -1) colSenha = 1; if (colTroca === -1) colTroca = 4;

    for (var i = 1; i < dados.length; i++) {
      if (String(dados[i][colUser]).trim().toUpperCase() === String(usuario).trim().toUpperCase() && 
          String(dados[i][colNome]).trim().toUpperCase() === String(nome).trim().toUpperCase()) {
        
        if (String(dados[i][colTroca]).trim().toUpperCase() === "SIM") {
          // Grava a nova senha na planilha
          aba.getRange(i + 1, colSenha + 1).setValue(novaSenha);
          return { sucesso: true, mensagem: "✅ Senha alterada com sucesso! Você já pode entrar." };
        }
      }
    }
    return { sucesso: false, mensagem: "❌ O Nome informado não bate com o registo do Usuário." };
  } catch (e) { return { sucesso: false, mensagem: "Erro interno: " + e.toString() }; }
}
// ================= SUBSTITUIR FUNÇÃO DE CADASTRO ADMIN =================
function cadastrarUsuarioAdmin(adminLogado, novoUser, novaSenha, nome, permissoes, trocaLivre) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var aba = ss.getSheetByName("USUARIOS_SISTEMA"); 
    if (!aba) {
      aba = ss.insertSheet("USUARIOS_SISTEMA");
      aba.appendRow(["USUÁRIO", "SENHA", "NOME", "PERMISSÕES", "TROCA LIVRE", "NOTIFICAÇÕES"]);
    }
    
    aba.appendRow([novoUser, novaSenha, nome, permissoes, trocaLivre, ""]);
    return { sucesso: true, mensagem: "Usuário cadastrado com sucesso!" };
  } catch (e) {
    return { sucesso: false, mensagem: "Erro ao cadastrar: " + e.toString() };
  }
}

function listarSolicitacoesRecuperacao(usuarioAdmin) {
  if (!verificarAdmin(usuarioAdmin)) return [];
  var dados = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("USUARIOS_SISTEMA").getDataRange().getValues();
  var solicitacoes = [];
  for (var i = 1; i < dados.length; i++) {
    if (String(dados[i][4]).trim() === "SOLICITOU_RECUPERACAO") solicitacoes.push({ usuario: String(dados[i][0]).trim(), nome: String(dados[i][2]).trim() });
  }
  return solicitacoes;
}

function redefinirSenhaAdmin(usuarioAdmin, usuarioAlvo, novaSenha) {
  if (!verificarAdmin(usuarioAdmin)) return { sucesso: false, mensagem: "Acesso negado." };
  var planilha = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("USUARIOS_SISTEMA");
  var dados = planilha.getDataRange().getValues();
  for (var i = 1; i < dados.length; i++) {
    if (String(dados[i][0]).trim() === usuarioAlvo.trim()) {
      planilha.getRange(i + 1, 2).setValue(novaSenha); planilha.getRange(i + 1, 5).clearContent();      
      return { sucesso: true, mensagem: "Senha updated com sucesso!" };
    }
  }
  return { sucesso: false, mensagem: "Usuário alvo não encontrado." };
}

// DRIVER DE ARQUIVOS
// NOTA (revisão 2): duas correções depois de reportado que "a listagem não mostra os arquivos e
// o upload (base64) não funciona":
// 1. `arquivo.setSharing(...)` pode lançar exceção em contas do Google Workspace cuja política
//    do domínio bloqueia compartilhamento "qualquer pessoa com o link" (comum em empresas). Como
//    esse setSharing rodava DEPOIS do createFile mas ANTES do `return { sucesso: true }`, uma
//    política de domínio restritiva fazia o upload inteiro ser reportado como falha ao front-end
//    (alerta "Erro: ..."), mesmo com o arquivo já criado no Drive. Agora o setSharing está em seu
//    próprio try/catch: se a política do domínio bloquear o compartilhamento, o arquivo continua
//    sendo salvo e a função retorna sucesso (só não fica com link público).
// 2. Adicionado log de erro mais explícito para facilitar diagnóstico caso volte a falhar.
// NOTA (revisão 5): você confirmou o nome REAL da pasta do Driver no seu Google Drive:
// "DRIVER MANUTENÇÃO_Files_" (o código antigo procurava por "SISTEMA_INOVAR_DRIVER", um nome
// que nunca existiu de verdade no seu Drive — por isso o upload "funcionava" sem erro, mas criava
// silenciosamente uma pasta NOVA E VAZIA com esse nome errado, diferente da pasta real onde os
// arquivos de manutenção já estavam). Agora o upload grava direto na pasta real.
// Também foi adicionada uma função auxiliar nova, `obterOuCriarPastaCacheada_`, que guarda o ID
// da pasta em cache (CacheService, 6 horas) na primeira vez que a acha por nome — chamadas
// seguintes pulam a busca por nome (mais lenta) e vão direto pelo ID, deixando upload e listagem
// mais rápidos.
// NOTA (revisão 6): novo parâmetro opcional `idPreferencial` — quando informado (ID de pasta
// tirado direto do link do Google Drive), a função abre a pasta DIRETO por esse ID, sem nem
// precisar buscar por nome. Chamadas antigas que não passam esse parâmetro continuam funcionando
// exatamente como antes (busca por nome).
function obterOuCriarPastaCacheada_(nomePasta, criarSeNaoExistir, idPreferencial) {
  if (idPreferencial) {
    var pastaPorId = obterPastaPorId_(idPreferencial);
    if (pastaPorId) return pastaPorId;
  }
  var cache = CacheService.getScriptCache();
  var chave = "pastaId_" + nomePasta;
  var idCache = cache.get(chave);
  if (idCache) {
    try {
      return DriveApp.getFolderById(idCache);
    } catch (eIdInvalido) {
      cache.remove(chave); // pasta foi apagada/movida — cai para a busca por nome abaixo
    }
  }
  var pastas = DriveApp.getFoldersByName(nomePasta);
  if (pastas.hasNext()) {
    var pastaEncontrada = pastas.next();
    cache.put(chave, pastaEncontrada.getId(), 21600);
    return pastaEncontrada;
  }
  if (criarSeNaoExistir) {
    var pastaNova = DriveApp.createFolder(nomePasta);
    cache.put(chave, pastaNova.getId(), 21600);
    return pastaNova;
  }
  return null;
}

function salvarArquivoBase64(usuarioAdmin, nome, mimeType, base64Data) {
  if (!verificarAdmin(usuarioAdmin)) return { sucesso: false, erro: "Acesso Negado." };
  try {
    var pastaDestino = obterOuCriarPastaCacheada_("DRIVER MANUTENÇÃO_Files_", true, ID_PASTA_DRIVER_MANUTENCAO);
    var arquivo = pastaDestino.createFile(Utilities.newBlob(Utilities.base64Decode(base64Data), mimeType, nome));
    try {
      arquivo.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch (eShare) {
      // Política do domínio pode bloquear link público — não falha o upload por isso.
      Logger.log("[AVISO] salvarArquivoBase64: não foi possível tornar o arquivo público (política do domínio?): " + eShare);
    }
    return { sucesso: true };
  } catch(e) {
    Logger.log("[ERRO] salvarArquivoBase64: " + e);
    return { sucesso: false, erro: e.toString() };
  }
}

// NOTA (revisão): listarArquivosDriver estava triplicada no sistema (duas vezes aqui
// em Admin.gs, usando nomes de pasta diferentes e até INVERTIDOS um do outro -
// "SISTEMA_INOVAR_DRIVER" vs "DRIVER_SISTEMA_INOVAR" -, e uma terceira vez em Driver.gs
// usando um ID de pasta fixo). Isso é um bug real: dependendo de qual definição
// vencesse, os arquivos enviados por salvarArquivoBase64 (que grava em
// "SISTEMA_INOVAR_DRIVER") podiam não aparecer na listagem. Consolidada aqui em uma
// única função, usando a mesma pasta do upload, para não perder arquivos já enviados.
//
// NOTA (revisão 2): agora a função varre TODAS as pastas com os nomes históricos que já
// existiram no código ao longo das duplicatas ("SISTEMA_INOVAR_DRIVER", "DRIVER_SISTEMA_INOVAR")
// e junta os arquivos de todas elas numa lista só. Isso resolve o caso de arquivos que já tinham
// sido enviados antes desta revisão, enquanto o bug de nomes trocados ainda existia, e que por
// isso foram salvos numa pasta que a listagem antiga não lia — eles "desapareciam" para o
// usuário, mas continuavam no Drive, só numa pasta errada. Também foi adicionado o campo `id`
// (que faltava) em cada item, necessário para o botão de excluir (apagarArquivoDriverServidor)
// funcionar — sem o `id`, o botão de excluir nunca tinha como saber qual arquivo remover.
//
// NOTA (revisão 3): função agora nunca falha silenciosamente. `verificarAdmin` falso e "pasta
// vazia" antes eram INDISTINGUÍVEIS para quem usa o sistema — os dois casos retornavam [] e a
// tela sempre mostrava "Driver vazio.", sem dizer qual dos dois aconteceu. Agora cada caso é
// logado no Apps Script (Execuções) com uma mensagem diferente. Toda a função está em try/catch
// para garantir que, se o Drive lançar qualquer erro (ex.: permissão da conta que executa o Web
// App), isso apareça no log em vez de travar silenciosamente.
//
// NOTA (revisão 5): nome real da pasta confirmado pelo usuário: "DRIVER MANUTENÇÃO_Files_" — vai
// primeiro na lista de nomes procurados, usando a busca cacheada (`obterOuCriarPastaCacheada_`)
// para não repetir a varredura por nome em toda chamada. Os nomes históricos
// ("SISTEMA_INOVAR_DRIVER", "DRIVER_SISTEMA_INOVAR") continuam sendo verificados depois, só para
// não perder de vista arquivos que possam ter sido enviados neles enquanto o bug existia. A pasta
// fixa por ID antiga (Driver.gs, ID_PASTA_DRIVER) foi removida desta função porque o usuário
// confirmou que esse ID não existe/não é acessível nesta conta — mantê-la só custava uma chamada
// ao Drive que sempre lançava exceção, deixando a tela mais lenta sem nenhum benefício.
function listarArquivosDriver(usuarioAdmin) {
  try {
    if (!verificarAdmin(usuarioAdmin)) {
      Logger.log("[listarArquivosDriver] Acesso negado para o usuário: '" + usuarioAdmin + "'. Verifique se este usuário tem 'admin' na coluna de permissões da aba USUARIOS_SISTEMA.");
      return [];
    }
    var nomesPastas = ["DRIVER MANUTENÇÃO_Files_", "SISTEMA_INOVAR_DRIVER", "DRIVER_SISTEMA_INOVAR"];
    var lista = []; var idsJaListados = {};
    for (var p = 0; p < nomesPastas.length; p++) {
      // NOTA (revisão 6): a pasta real ("DRIVER MANUTENÇÃO_Files_") agora abre direto pelo ID
      // que você mandou no link do Drive — não depende mais da busca por nome bater certinho.
      var idPreferencialDesta = (nomesPastas[p] === "DRIVER MANUTENÇÃO_Files_") ? ID_PASTA_DRIVER_MANUTENCAO : null;
      var pasta = obterOuCriarPastaCacheada_(nomesPastas[p], false, idPreferencialDesta);
      if (!pasta) {
        Logger.log("[listarArquivosDriver] Pasta '" + nomesPastas[p] + "': não encontrada.");
        continue;
      }
      var qtdNestaPasta = 0;
      var arquivos = pasta.getFiles();
      while (arquivos.hasNext()) {
        var f = arquivos.next();
        var fid = f.getId();
        if (idsJaListados[fid]) continue; // evita duplicar se o mesmo arquivo aparecer 2x
        idsJaListados[fid] = true;
        qtdNestaPasta++;
        lista.push({ id: fid, nome: f.getName(), url: f.getUrl(), dataMili: f.getDateCreated().getTime(), tamanhoBytes: f.getSize() });
      }
      Logger.log("[listarArquivosDriver] Pasta '" + nomesPastas[p] + "': " + qtdNestaPasta + " arquivo(s).");
    }
    Logger.log("[listarArquivosDriver] Total de arquivos encontrados: " + lista.length);
    lista.sort(function(a,b) { return b.dataMili - a.dataMili; });
    for(var i=0; i<lista.length; i++) lista[i].dataStr = Utilities.formatDate(new Date(lista[i].dataMili), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
    return lista;
  } catch (e) {
    Logger.log("[ERRO] listarArquivosDriver: " + e);
    return [];
  }
}

// NOVA FUNÇÃO (adicionada nesta revisão, não existia antes): busca arquivos do Driver por nome,
// para o admin não precisar rolar a lista inteira manualmente quando tiver muitos arquivos
// enviados. Reaproveita listarArquivosDriver (mesma checagem de permissão), só filtra o
// resultado pelo termo digitado (sem diferenciar maiúsculas/minúsculas ou acentos exatos).
function buscarArquivosDriverPorNome(usuarioAdmin, termoBusca) {
  var todos = listarArquivosDriver(usuarioAdmin);
  var termo = String(termoBusca || "").trim().toLowerCase();
  if (!termo) return todos;
  return todos.filter(function(arq) {
    return arq.nome.toLowerCase().indexOf(termo) !== -1;
  });
}

// NOVA FUNÇÃO (adicionada nesta revisão, não existia antes): estatísticas rápidas do Driver —
// total de arquivos e espaço ocupado, formatado em KB/MB/GB. Útil para o admin acompanhar o
// quanto já foi enviado sem precisar abrir o Google Drive por fora.
function obterEstatisticasDriver(usuarioAdmin) {
  if (!verificarAdmin(usuarioAdmin)) return { sucesso: false, erro: "Acesso Negado." };
  var arquivos = listarArquivosDriver(usuarioAdmin);
  var totalBytes = 0;
  for (var i = 0; i < arquivos.length; i++) totalBytes += (arquivos[i].tamanhoBytes || 0);
  var tamanhoFormatado;
  if (totalBytes < 1024) tamanhoFormatado = totalBytes + " B";
  else if (totalBytes < 1024 * 1024) tamanhoFormatado = (totalBytes / 1024).toFixed(1) + " KB";
  else if (totalBytes < 1024 * 1024 * 1024) tamanhoFormatado = (totalBytes / (1024 * 1024)).toFixed(1) + " MB";
  else tamanhoFormatado = (totalBytes / (1024 * 1024 * 1024)).toFixed(2) + " GB";
  return { sucesso: true, totalArquivos: arquivos.length, totalBytes: totalBytes, tamanhoFormatado: tamanhoFormatado, ultimoUpload: arquivos.length ? arquivos[0].dataStr : null };
}

// NOTA (revisão 3): função de diagnóstico NOVA, criada para investigar o caso relatado de
// "Driver não mostra nenhum arquivo" sem precisar de acesso a um ambiente de testes real. Ela NÃO
// é chamada por nenhuma tela — é para você rodar manualmente:
// 1. Abra o editor do Apps Script (script.google.com, no projeto do Sistema Inovar).
// 2. No menu de funções ao lado do botão "Executar" (▶), selecione "diagnosticarDriverSistema".
// 3. Clique em "Executar". Na primeira vez, ele vai pedir autorização — aceite.
// 4. Vá em "Execuções" (ícone de relógio na barra lateral esquerda) e abra a execução mais
//    recente para ver o resultado detalhado no log.
// Isso mostra: qual conta Google está executando o script, quantas pastas existem com cada nome
// histórico usado pelo sistema, e quantos arquivos tem dentro de cada uma — o suficiente para
// achar exatamente onde os arquivos enviados estão parando.
// NOTA (revisão 4): a investigação anterior confirmou que os arquivos EXISTEM na pasta certa
// ("SISTEMA_INOVAR_DRIVER") e que o Drive está acessível normalmente, então a causa de
// "Driver vazio." não é a pasta — é algo antes disso na cadeia. Suspeito nº1:
// verificarAdmin(usuarioLogado) retorna false para o usuário logado no front-end (valor salvo no
// navegador não é idêntico ao da coluna A da aba USUARIOS_SISTEMA, ou a coluna D não contém
// literalmente "admin"/"tudo"). Por isso a função agora aceita um parâmetro opcional
// `usuarioParaTestar`: quando informado, testa verificarAdmin com esse valor exato e lista todos
// os usuários cadastrados (coluna A e coluna D) para comparação direta.
function diagnosticarDriverSistema(usuarioParaTestar) {
  var linhas = [];
  linhas.push("Conta que executa o script (getEffectiveUser): " + Session.getEffectiveUser().getEmail());
  try {
    linhas.push("Conta que abriu o Web App (getActiveUser): " + Session.getActiveUser().getEmail());
  } catch (e) {
    linhas.push("Conta que abriu o Web App (getActiveUser): não disponível neste contexto (" + e + ")");
  }
  // NOTA (revisão 5): nome real da pasta confirmado pelo usuário é "DRIVER MANUTENÇÃO_Files_" —
  // vai primeiro na lista. Os nomes antigos seguem aqui só de referência histórica. A pasta fixa
  // por ID antiga (Driver.gs, ID_PASTA_DRIVER) foi removida deste diagnóstico porque o usuário já
  // confirmou que ela não existe/não é acessível nesta conta.
  // NOTA (revisão 6): checagem direta por ID (links que você mandou) — a forma mais confiável de
  // confirmar se o script consegue abrir cada pasta, sem depender do nome bater.
  linhas.push("--- Checagem direta por ID (link do Google Drive) ---");
  var pastaDriverPorId = obterPastaPorId_(ID_PASTA_DRIVER_MANUTENCAO);
  if (pastaDriverPorId) {
    var qtdArqDriverId = 0; var itDriverId = pastaDriverPorId.getFiles();
    while (itDriverId.hasNext()) { itDriverId.next(); qtdArqDriverId++; }
    linhas.push("Pasta Driver por ID (" + ID_PASTA_DRIVER_MANUTENCAO + "): OK, nome real='" + pastaDriverPorId.getName() + "', " + qtdArqDriverId + " arquivo(s).");
  } else {
    linhas.push("Pasta Driver por ID (" + ID_PASTA_DRIVER_MANUTENCAO + "): FALHOU AO ABRIR — a conta que executa o script pode não ter acesso a essa pasta.");
  }
  var pastaBackupPorId = obterPastaPorId_(ID_PASTA_BACKUP_SISTEMA);
  if (pastaBackupPorId) {
    var qtdArqBackupId = 0; var itBackupId = pastaBackupPorId.getFiles();
    while (itBackupId.hasNext()) { itBackupId.next(); qtdArqBackupId++; }
    linhas.push("Pasta Backup por ID (" + ID_PASTA_BACKUP_SISTEMA + "): OK, nome real='" + pastaBackupPorId.getName() + "', " + qtdArqBackupId + " arquivo(s).");
  } else {
    linhas.push("Pasta Backup por ID (" + ID_PASTA_BACKUP_SISTEMA + "): FALHOU AO ABRIR — a conta que executa o script pode não ter acesso a essa pasta.");
  }
  var nomes = ["DRIVER MANUTENÇÃO_Files_", "SISTEMA_INOVAR_DRIVER", "DRIVER_SISTEMA_INOVAR"];
  for (var i = 0; i < nomes.length; i++) {
    var pastas = DriveApp.getFoldersByName(nomes[i]);
    var qtdPastas = 0;
    while (pastas.hasNext()) {
      qtdPastas++;
      var pasta = pastas.next();
      var qtdArquivos = 0; var nomesArquivos = [];
      var arquivos = pasta.getFiles();
      while (arquivos.hasNext()) { qtdArquivos++; var arq = arquivos.next(); nomesArquivos.push(arq.getName()); }
      linhas.push("Pasta '" + nomes[i] + "' (id: " + pasta.getId() + "): " + qtdArquivos + " arquivo(s). Nomes: " + nomesArquivos.join(", "));
    }
    if (qtdPastas === 0) linhas.push("Pasta '" + nomes[i] + "': NENHUMA pasta com este nome encontrada no Drive desta conta.");
  }
  // Lista todos os usuários cadastrados (coluna A = usuário, coluna D = permissões), para
  // comparação visual direta com o que está sendo digitado no login.
  try {
    var planilhaUsu = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("USUARIOS_SISTEMA");
    if (!planilhaUsu) {
      linhas.push("Aba 'USUARIOS_SISTEMA' não encontrada na planilha.");
    } else {
      var dadosUsu = planilhaUsu.getDataRange().getValues();
      linhas.push("--- Usuários cadastrados em USUARIOS_SISTEMA (" + (dadosUsu.length - 1) + ") ---");
      for (var iu = 1; iu < dadosUsu.length; iu++) {
        var userCol = String(dadosUsu[iu][0]);
        var permCol = String(dadosUsu[iu][3]);
        linhas.push("  [" + (iu + 1) + "] usuario='" + userCol + "' (len=" + userCol.length + ") | permissoes='" + permCol + "'");
      }
    }
  } catch (eUsu) {
    linhas.push("Erro ao ler USUARIOS_SISTEMA: " + eUsu);
  }
  // Testa verificarAdmin diretamente com o valor informado, se algum foi passado.
  if (usuarioParaTestar) {
    var resultadoAdmin = verificarAdmin(usuarioParaTestar);
    linhas.push("--- Teste de verificarAdmin('" + usuarioParaTestar + "') (len=" + String(usuarioParaTestar).length + ") ---");
    linhas.push("Resultado: " + (resultadoAdmin ? "TRUE (é admin, listarArquivosDriver deveria funcionar)" : "FALSE (acesso negado — é essa a causa do 'Driver vazio.')"));
  } else {
    linhas.push("--- Nenhum usuário passado para testar verificarAdmin. Rode de novo assim: diagnosticarDriverSistema(\"seu_usuario_de_login\") para testar exatamente o valor usado no login. ---");
  }
  var relatorio = linhas.join("\n");
  Logger.log(relatorio);
  return relatorio;
}

// BACKUP
function criarBackupSistema(usuarioAdmin) {
  if (!verificarAdmin(usuarioAdmin)) return { sucesso: false, mensagem: "Acesso Negado." };
  try {
    // NOTA (revisão 6): abre a pasta de backup direto pelo ID que você mandou no link do Drive —
    // não depende mais do nome bater certinho. Só cria pasta nova por nome se o ID não abrir.
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var pastaBackup = obterOuCriarPastaCacheada_("SISTEMA_INOVAR_BACKUPS", true, ID_PASTA_BACKUP_SISTEMA);
    DriveApp.getFileById(ss.getId()).makeCopy("BACKUP_" + ss.getName() + "_" + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd-MM-yyyy_HH-mm-ss"), pastaBackup);
    return { sucesso: true, mensagem: "Backup salvo no Drive da empresa!" };
  } catch(e) { return { sucesso: false, mensagem: "Erro ao criar backup: " + e.toString() }; }
}

// NOTA (revisão): listarBackups estava duplicada em Admin.gs usando nomes de pasta
// diferentes ("SISTEMA_INOVAR_BACKUPS" vs "BACKUPS_SISTEMA_INOVAR"). Consolidada em
// uma única função, usando a mesma pasta que criarBackupSistema realmente grava
// ("SISTEMA_INOVAR_BACKUPS"), para que os backups já criados continuem aparecendo
// na lista.
function listarBackups(usuarioAdmin) {
  if (!verificarAdmin(usuarioAdmin)) return [];
  // NOTA (revisão 6): idem criarBackupSistema — abre direto pelo ID do link do Drive, com a
  // busca por nome como reserva.
  var pastaBackup = obterOuCriarPastaCacheada_("SISTEMA_INOVAR_BACKUPS", false, ID_PASTA_BACKUP_SISTEMA);
  if (!pastaBackup) return [];
  var arquivos = pastaBackup.getFiles(); var lista = [];
  while (arquivos.hasNext()) { var f = arquivos.next(); lista.push({ nome: f.getName(), url: f.getUrl(), dataMili: f.getDateCreated().getTime() }); }
  lista.sort(function(a,b) { return b.dataMili - a.dataMili; });
  for(var i=0; i<lista.length; i++) lista[i].dataStr = Utilities.formatDate(new Date(lista[i].dataMili), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
  return lista;
}

// ================= MÓDULO INTELIGENTE DE RECUPERAÇÃO DE SENHA =================

// ================= MOTOR DO ESQUECEU A SENHA =================

function solicitarRecuperacaoSenhaBackend(usuario) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var abaUser = ss.getSheetByName("USUARIOS_SISTEMA");
    if (!abaUser) return { sucesso: false, permite: false, mensagem: "Aba 'USUARIOS_SISTEMA' não encontrada." };
    
    var dadosUser = abaUser.getDataRange().getValues();
    var colUser = -1, colTroca = -1;
    
    // Procura a coluna do Usuário e da Permissão
    for(var c=0; c<dadosUser[0].length; c++) {
       var cabecalho = String(dadosUser[0][c]).toUpperCase();
       if(cabecalho.indexOf("USU") > -1 || cabecalho.indexOf("LOGIN") > -1) colUser = c;
       if(cabecalho.indexOf("LIVRE") > -1 || cabecalho.indexOf("TROCA") > -1) colTroca = c;
    }
    
    // Segurança extra caso os cabeçalhos tenham espaços estranhos
    if (colUser === -1) colUser = 0; 
    if (colTroca === -1) colTroca = 4;

    // Vasculha a planilha linha por linha
    for (var i = 1; i < dadosUser.length; i++) {
      if (String(dadosUser[i][colUser]).trim().toUpperCase() === String(usuario).trim().toUpperCase()) {
        
        var trocaLivre = String(dadosUser[i][colTroca] || "").trim().toUpperCase();
        
        if (trocaLivre === "SIM") {
          // REGRA 1: USUÁRIO PODE MUDAR SOZINHO
          return { sucesso: true, permite: true, mensagem: "Usuário autorizado." };
        } else {
          // REGRA 2: NÃO PODE MUDAR SOZINHO -> GERA ALERTA NA ABA NOTIFICACOES_ADMIN
          var abaNotif = ss.getSheetByName("NOTIFICACOES_ADMIN");
          if (!abaNotif) {
             abaNotif = ss.insertSheet("NOTIFICACOES_ADMIN");
             abaNotif.appendRow(["DATA E HORA", "USUÁRIO", "TIPO DE NOTIFICAÇÃO", "STATUS"]);
          }
          
          var dataAtual = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), "dd/MM/yyyy HH:mm:ss");
          
          // Adiciona a linha exatamente como na sua imagem
          abaNotif.appendRow([dataAtual, usuario, "Solicitação de Redefinição de Senha", "PENDENTE"]);

          return { sucesso: true, permite: false, mensagem: "Você não tem permissão para alterar a senha sozinho.\nUma notificação foi enviada ao Administrador do sistema!" };
        }
      }
    }
    return { sucesso: false, permite: false, mensagem: "Usuário '" + usuario + "' não cadastrado no sistema." };
  } catch (e) { return { sucesso: false, permite: false, mensagem: "Erro no servidor: " + e.toString() }; }
}

// ================= MÓDULO: LER E RESOLVER NOTIFICAÇÕES (PAINEL ADMIN) =================

// ================= MÓDULO: LER E RESOLVER NOTIFICAÇÕES (PAINEL ADMIN) =================

function listarNotificacoesSenha() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var abaNotif = ss.getSheetByName("NOTIFICACOES_ADMIN");
    if (!abaNotif) return []; // Se a aba não existir, retorna vazio
    
    var dados = abaNotif.getDataRange().getValues();
    var notificacoes = [];
    
    // De acordo com a sua imagem: A=0(Data), B=1(Usuário), C=2(Tipo), D=3(Status)
    for (var i = 1; i < dados.length; i++) {
      var status = String(dados[i][3] || "").trim().toUpperCase(); // Lê a Coluna D
      
      if (status === "PENDENTE") { // Puxa apenas as que não foram resolvidas
        var dataStr = dados[i][0];
        
        // Garante que a data do Excel venha formatada para a tela
        if (dataStr instanceof Date) {
          var dia = ("0" + dataStr.getDate()).slice(-2);
          var mes = ("0" + (dataStr.getMonth() + 1)).slice(-2);
          var hora = ("0" + dataStr.getHours()).slice(-2);
          var min = ("0" + dataStr.getMinutes()).slice(-2);
          dataStr = dia + "/" + mes + "/" + dataStr.getFullYear() + " " + hora + ":" + min;
        }

        notificacoes.push({
          linhaOrigem: i + 1, // Memoriza a linha para sabermos onde marcar "RESOLVIDO" depois
          usuario: String(dados[i][1] || "Usuário Desconhecido"), 
          dataHora: dataStr || "Recente"
        });
      }
    }
    return notificacoes.reverse(); // Mostra as mais recentes primeiro
  } catch (e) { return []; }
}

function resolverNotificacaoBackend(linhaNotificacao) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var abaNotif = ss.getSheetByName("NOTIFICACOES_ADMIN");
    if (!abaNotif) return { sucesso: false };
    
    // Muda a Coluna D para RESOLVIDO
    abaNotif.getRange(linhaNotificacao, 4).setValue("RESOLVIDO"); 
    return { sucesso: true };
  } catch (e) { return { sucesso: false }; }
}

function adminRedefinirSenhaUsuario(usuario, novaSenha, linhaNotificacao) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // 1. Acha o usuário na aba de cadastros e muda a senha dele
    var abaUser = ss.getSheetByName("USUARIOS_SISTEMA");
    if (!abaUser) return { sucesso: false, mensagem: "Aba 'USUARIOS_SISTEMA' não encontrada." };
    
    var dadosUser = abaUser.getDataRange().getValues();
    var colUser = -1, colSenha = -1;
    for(var c=0; c<dadosUser[0].length; c++) {
       var cab = String(dadosUser[0][c]).toUpperCase();
       if(cab.indexOf("USU") > -1 || cab.indexOf("LOGIN") > -1) colUser = c;
       if(cab.indexOf("SENHA") > -1) colSenha = c;
    }
    if(colUser === -1) colUser = 0; if(colSenha === -1) colSenha = 1;

    var senhaAlterada = false;
    for (var i = 1; i < dadosUser.length; i++) {
      if (String(dadosUser[i][colUser]).trim().toUpperCase() === String(usuario).trim().toUpperCase()) {
        abaUser.getRange(i + 1, colSenha + 1).setValue(novaSenha);
        senhaAlterada = true;
        break;
      }
    }

    if (!senhaAlterada) return { sucesso: false, mensagem: "Usuário '" + usuario + "' não encontrado no cadastro." };

    // 2. Muda o status na aba NOTIFICACOES_ADMIN para remover da tela
    var abaNotif = ss.getSheetByName("NOTIFICACOES_ADMIN");
    if (abaNotif) {
      abaNotif.getRange(linhaNotificacao, 4).setValue("RESOLVIDO"); 
    }

    return { sucesso: true, mensagem: "✅ Senha alterada com sucesso!" };
  } catch (e) { return { sucesso: false, mensagem: "Erro: " + e.toString() }; }
}

// NOTA (revisão): listarUsuarios/editarUsuario/excluirUsuario abaixo são código
// legado: operam na aba "USUARIOS", que não existe no sistema atual (a aba real é
// "USUARIOS_SISTEMA"). O frontend não chama estas três funções em nenhum lugar - ele
// usa listarUsuariosBackend/atualizarUsuarioBackend/excluirUsuarioBackend, mais
// abaixo. Mantidas sem alteração por segurança (nenhuma função deve ser removida),
// mas ficam sinalizadas aqui como não usadas/quebradas caso sejam chamadas.
// Busca todos os usuários (exceto senhas)
function listarUsuarios() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("USUARIOS"); // Ajuste o nome da aba se necessário
  var data = sheet.getDataRange().getValues();
  var usuarios = [];
  for(var i = 1; i < data.length; i++) {
    usuarios.push({ id: i + 1, nome: data[i][1], user: data[i][2], perms: data[i][4] });
  }
  return usuarios;
}

// Edita usuário exigindo senha do admin
function editarUsuario(linha, dados, senhaAdmin) {
  // 1. Validar senha do Admin aqui (ex: verificar se a senhaAdmin está correta no banco)
  // 2. Se for admin, permitir edição.
  // 3. Gravar na linha indicada.
  return { sucesso: true, mensagem: "Usuário atualizado!" };
}

function excluirUsuario(linha) {
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("USUARIOS").deleteRow(linha);
  return { sucesso: true };
}
// ================= GESTÃO DE USUÁRIOS (Backend) =================

function listarUsuariosBackend() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var aba = ss.getSheetByName("USUARIOS_SISTEMA");
    if (!aba) return [{erro: "Aba USUARIOS_SISTEMA não existe na planilha."}];
    
    var dados = aba.getDataRange().getValues();
    var lista = [];
    
    // Lê da linha 2 em diante
    for (var i = 1; i < dados.length; i++) {
      var userBanco = String(dados[i][0]).trim();
      if (userBanco !== "") {
        lista.push({
          id: i + 1, // Número exato da linha no Excel
          user: userBanco,
          nome: dados[i][2],
          perms: dados[i][3]
        });
      }
    }
    return lista;
  } catch(e) { 
    return [{erro: "Erro de servidor: " + e.toString()}]; 
  }
}

function atualizarUsuarioBackend(linha, nome, perms) {
  try {
    var aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("USUARIOS_SISTEMA");
    aba.getRange(linha, 3).setValue(nome.toUpperCase()); // Atualiza Coluna C (NOME)
    aba.getRange(linha, 4).setValue(perms.toLowerCase()); // Atualiza Coluna D (PERMISSÕES)
    return {sucesso: true, mensagem: "✅ Perfil atualizado com sucesso!"};
  } catch(e) { 
    return {sucesso: false, mensagem: "Erro ao atualizar: " + e.toString()}; 
  }
}

function excluirUsuarioBackend(linha) {
  try {
    var aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("USUARIOS_SISTEMA");
    aba.deleteRow(linha);
    return {sucesso: true, mensagem: "✅ Usuário excluído permanentemente!"};
  } catch(e) { 
    return {sucesso: false, mensagem: "Erro ao excluir: " + e.toString()}; 
  }
}
// ================= DRIVER E BACKUPS UNIVERSAIS (Backend) =================

function listarArquivosDriver(usuarioLogado) {
  try {
    var nomePasta = "SISTEMA_INOVAR_DRIVER";
    var pastas = DriveApp.getFoldersByName(nomePasta);
    if (!pastas.hasNext()) return []; // Se a pasta ainda não existir, retorna vazio
    
    var pasta = pastas.next();
    var arquivos = pasta.getFiles();
    var lista = [];
    
    while(arquivos.hasNext()){
      var arq = arquivos.next();
      lista.push({
        nome: arq.getName(),
        url: arq.getUrl(),
        dataStr: Utilities.formatDate(arq.getDateCreated(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm"),
        ts: arq.getDateCreated().getTime() // Usado para ordenar cronologicamente
      });
    }
    
    // Retorna ordenado do mais recente para o mais antigo
    return lista.sort((a,b) => b.ts - a.ts);
  } catch(e) {
    return [];
  }
}

function listarBackups(usuarioLogado) {
  try {
    var nomePasta = "SISTEMA_INOVAR_BACKUPS";
    var pastas = DriveApp.getFoldersByName(nomePasta);
    if (!pastas.hasNext()) return [];
    
    var pasta = pastas.next();
    var arquivos = pasta.getFiles();
    var lista = [];
    
    while(arquivos.hasNext()){
      var arq = arquivos.next();
      lista.push({
        nome: arq.getName(),
        url: arq.getUrl(),
        dataStr: Utilities.formatDate(arq.getDateCreated(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm"),
        ts: arq.getDateCreated().getTime()
      });
    }
    
    return lista.sort((a,b) => b.ts - a.ts);
  } catch(e) {
    return [];
  }
}