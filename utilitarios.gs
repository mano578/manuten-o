// ==============================================================
// UTILITARIOS.GS — Funções auxiliares novas, adicionadas na revisão
// ==============================================================
// NOTA (revisão): Este arquivo é uma ADIÇÃO da revisão de código — não existia no sistema
// original. Nenhuma função existente foi removida ou substituída por este arquivo; ele apenas
// oferece helpers reutilizáveis para reduzir duplicação de código e centralizar validações de
// segurança básicas. Nenhum outro arquivo foi obrigado a usá-los — foram aplicados apenas onde
// fazia sentido sem arriscar quebrar comportamento existente (sem ambiente de teste real
// disponível para validar mudanças mais profundas).
//
// Se quiser adotar estes helpers em mais lugares do sistema no futuro, basta chamar as funções
// abaixo a partir dos outros arquivos .gs (todas as funções em Apps Script compartilham o mesmo
// escopo global do projeto).

/**
 * Normaliza uma string para comparação: remove espaços nas pontas e converte para maiúsculas.
 * Reduz duplicação do padrão `String(x).trim().toUpperCase()` repetido centenas de vezes no
 * sistema original.
 */
function normalizarTexto(valor) {
  return String(valor === null || valor === undefined ? "" : valor).trim().toUpperCase();
}

/**
 * Sanitiza uma string simples que será usada em HTML dinâmico (innerHTML) no front-end,
 * escapando caracteres que poderiam gerar HTML/script indesejado. Uso recomendado para qualquer
 * texto vindo da planilha e exibido via innerHTML no front-end (o sistema original insere muitos
 * valores da planilha diretamente em innerHTML sem escapar).
 */
function escaparHtml(valor) {
  var s = String(valor === null || valor === undefined ? "" : valor);
  return s
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

/**
 * Busca a primeira aba existente entre uma lista de nomes possíveis. Útil porque o sistema
 * original tem várias abas com mais de um nome histórico usado em partes diferentes do código
 * (ex.: "SOS_FROTA" vs "CENTRAL_SOS", "ORDEM_SERVICO_FROTA" vs "ORDEM_SERVICO"). Centraliza esse
 * padrão de fallback em vez de repetir `getSheetByName(a) || getSheetByName(b)` em cada função.
 * @param {Spreadsheet} ss - planilha ativa (SpreadsheetApp.getActiveSpreadsheet())
 * @param {string[]} nomes - lista de nomes possíveis, em ordem de prioridade
 * @return {Sheet|null}
 */
function encontrarAbaPorNomes(ss, nomes) {
  for (var i = 0; i < nomes.length; i++) {
    var aba = ss.getSheetByName(nomes[i]);
    if (aba) return aba;
  }
  return null;
}

/**
 * Verifica de forma centralizada se um objeto de sessão/usuário logado tem permissão de admin.
 * Vários arquivos (Admin.gs, Driver.gs, etc.) repetem a checagem de admin de formas ligeiramente
 * diferentes. Esta função não substitui nenhuma checagem existente automaticamente (para não
 * arriscar mudar comportamento sem poder testar), mas pode ser adotada progressivamente.
 * @param {Object} permissoes - array de permissões do usuário (já em minúsculas, como retornado
 *   por validarLogin em App.gs)
 */
function usuarioEhAdmin(permissoes) {
  if (!permissoes) return false;
  var lista = Array.isArray(permissoes) ? permissoes : String(permissoes).split(",");
  for (var i = 0; i < lista.length; i++) {
    if (normalizarTexto(lista[i]) === "ADMIN") return true;
  }
  return false;
}

/**
 * Wrapper de log central para erros — facilita encontrar todos os pontos de falha registrados
 * pelo sistema em um único lugar (visualizações > execuções, no editor do Apps Script), com um
 * prefixo padronizado por módulo.
 */
function registrarErro(modulo, funcao, erro) {
  try {
    Logger.log("[ERRO] " + modulo + "." + funcao + ": " + (erro && erro.message ? erro.message : erro));
  } catch (e) {
    // Nunca deixa o log quebrar a função que o chamou.
  }
}
