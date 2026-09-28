// ================= MÓDULO BACKEND: CHECKLIST DIÁRIO DOS MOTORISTAS =================
// Nova aba de planilha usada por este módulo: "CHECKLIST_DIARIO_MOTORISTAS" (criada automaticamente
// na primeira gravação, igual ao padrão já usado em Custos.gs/Abastecimento.gs).
//
// 🆕 PARA ADICIONAR/EDITAR ITENS DO CHECKLIST NO FUTURO: edite SOMENTE a função
// `obterDefinicaoChecklistPorTipo` abaixo. Ela é a ÚNICA fonte dos itens (básicos + específicos por
// tipo de equipamento) usada tanto para montar o formulário quanto para conferir a gravidade ao salvar.
// Não é necessário tocar em nenhum arquivo .html — o formulário é montado dinamicamente a partir
// do que essa função devolve.

// Lista fixa dos itens BÁSICOS, comuns a qualquer veículo/equipamento.
function _itensBasicosChecklist() {
  return [
    { id: "nivel_oleo", texto: "Nível de óleo do motor" },
    { id: "nivel_agua", texto: "Nível de água/arrefecimento" },
    { id: "freios", texto: "Freios (pedal e freio de mão)" },
    { id: "pneus", texto: "Pneus (calibragem e desgaste)" },
    { id: "iluminacao", texto: "Faróis, lanternas e setas" },
    { id: "buzina", texto: "Buzina" },
    { id: "cinto", texto: "Cinto de segurança" },
    { id: "extintor", texto: "Extintor de incêndio (válido)" },
    { id: "limpeza", texto: "Limpeza geral do veículo" },
    { id: "documentacao", texto: "Documentação (CRLV/checklist) em dia" },
    { id: "combustivel", texto: "Combustível suficiente para o turno" },
    { id: "retrovisores", texto: "Retrovisores e para-brisa" }
  ];
}

// Mapeia o texto livre do campo "Informação/Tipo de Equipamento" (cadastro do veículo, coluna Y)
// para um dos grupos de checklist específico abaixo. Usa busca por palavra-chave (case-insensitive)
// para tolerar variações de cadastro (ex.: "COMPACTADOR 19 M³" cai no grupo "COMPACTADOR").
function mapearGrupoChecklist(informacaoTexto) {
  var t = String(informacaoTexto || "").toUpperCase();

  // Compactador e variações de coleta usam o mesmo formulário físico "INSPEÇÃO E CONTROLE DE
  // SAÍDA DA FROTA" (mesmos blocos/itens), por isso caem todos no grupo COMPACTADOR.
  if (t.indexOf("COMPACTADOR") > -1) return "COMPACTADOR";
  if (t.indexOf("CAÇAMBA") > -1 || t.indexOf("CACAMBA") > -1 || t.indexOf("CARROCERIA") > -1) return "COMPACTADOR";
  if (t.indexOf("ENTULHO") > -1) return "COMPACTADOR";
  if (t.indexOf("PODA") > -1) return "COMPACTADOR";
  if (t.indexOf("TRITURAD") > -1) return "COMPACTADOR";
  if (t.indexOf("RETROESCAVADEIRA") > -1 || t.indexOf("SANEADORA") > -1 || t.indexOf("TRATOR") > -1) return "MAQUINA_PESADA";
  if (t.indexOf("VARREDEIRA") > -1) return "VARREDEIRA";
  if (t.indexOf("MOTO") > -1) return "MOTO";
  if (t.indexOf("ÔNIBUS") > -1 || t.indexOf("ONIBUS") > -1) return "ONIBUS";
  if (t.indexOf("PINTURA") > -1) return "MAQUINA_PINTURA";
  if (t.indexOf("REBOQUE") > -1) return "REBOQUE";
  // Saveiro e demais veículos leves/administrativos usam o formulário físico "CHECKLIST DE
  // VEÍCULO LEVE — VISTORIA DE SAÍDA DA GARAGEM".
  if (t.indexOf("SAVEIRO") > -1) return "VEICULO_LEVE";
  if (t.indexOf("CARRO") > -1) return "CARRO_ADMINISTRATIVO";

  return "GERAL";
}

// Itens do formulário físico "INSPEÇÃO E CONTROLE DE SAÍDA DA FROTA" (compactadores e demais
// veículos de coleta: caçamba de entulho manual, poda manual, poda triturada, entulho mecanizado),
// organizados por bloco igual ao papel usado hoje na garagem.
function _itensCompactadorPorBloco() {
  return [
    { id: "cab_documentos", texto: "Documentos obrigatórios ok?", bloco: "Cabine/Interior" },
    { id: "cab_velocimetro", texto: "Velocímetro funcionando?", bloco: "Cabine/Interior" },
    { id: "cab_tacografo", texto: "Tacógrafo funcionando?", bloco: "Cabine/Interior" },
    { id: "cab_tacografo_disco", texto: "Tacógrafo com disco e lacre ok?", bloco: "Cabine/Interior" },
    { id: "cab_ar_radio", texto: "Ar-condicionado e rádio ok?", bloco: "Cabine/Interior" },
    { id: "cab_carenagem", texto: "Carenagens do painel da cabine ok?", bloco: "Cabine/Interior" },
    { id: "cab_volante_buzina", texto: "Volante e tampa da buzina ok?", bloco: "Cabine/Interior" },
    { id: "cab_porta_luva", texto: "Tampa do porta-luva ok?", bloco: "Cabine/Interior" },
    { id: "cab_porta_fusivel", texto: "Tampa do porta-fusível ok?", bloco: "Cabine/Interior" },
    { id: "cab_cinzeiro", texto: "Cinzeiro e isqueiro ok?", bloco: "Cabine/Interior" },
    { id: "cab_parabrisa", texto: "Para-brisas dianteiro e traseiro ok?", bloco: "Cabine/Interior" },
    { id: "cab_vidros", texto: "Vidros das portas ok?", bloco: "Cabine/Interior" },
    { id: "cab_elevador", texto: "Elevador das portas ok?", bloco: "Cabine/Interior" },
    { id: "cab_limpador", texto: "Limpador de para-brisa ok?", bloco: "Cabine/Interior" },
    { id: "cab_lampadas_salao", texto: "Lâmpadas do salão ok?", bloco: "Cabine/Interior" },
    { id: "cab_cintos", texto: "Cintos de segurança ok?", bloco: "Cabine/Interior" },
    { id: "cab_bancos", texto: "Bancos e forração interior ok?", bloco: "Cabine/Interior" },
    { id: "cab_macanetas", texto: "Maçanetas interna e externa ok?", bloco: "Cabine/Interior" },
    { id: "cab_retrovisores", texto: "Retrovisores e hastes ok?", bloco: "Cabine/Interior" },
    { id: "cab_interior_limpo", texto: "Interior da cabine ok e limpo?", bloco: "Cabine/Interior" },

    { id: "susp_molas", texto: "Molas ok?", bloco: "Suspensão" },
    { id: "susp_amort_diant", texto: "Amortecedores dianteiros ok?", bloco: "Suspensão" },
    { id: "susp_amort_tras", texto: "Amortecedores traseiros ok?", bloco: "Suspensão" },

    { id: "comp_comando_tras", texto: "Comandos traseiro HD ok?", bloco: "Compactador" },
    { id: "comp_comando_diant", texto: "Comandos dianteiro HD ok?", bloco: "Compactador" },
    { id: "comp_tem_chorume", texto: "Tem caixa de chorume?", bloco: "Compactador" },
    { id: "comp_chorume_limpa", texto: "Caixa de chorume limpa?", bloco: "Compactador" },
    { id: "comp_acel_automatica", texto: "Aceleração automática ok?", bloco: "Compactador" },
    { id: "comp_plataforma_gari", texto: "Plataforma do gari ok?", bloco: "Compactador" },
    { id: "comp_alcas_gari", texto: "Alças do gari ok?", bloco: "Compactador" },

    { id: "hid_agua_radiador", texto: "Nível de água do radiador ok?", bloco: "Hidráulica e Lubrificação" },
    { id: "hid_oleo_direcao", texto: "Óleo da direção hidráulica ok?", bloco: "Hidráulica e Lubrificação" },
    { id: "hid_oleo_motor", texto: "Nível de óleo do motor ok?", bloco: "Hidráulica e Lubrificação" },
    { id: "hid_oleo_sistema", texto: "Óleo do sistema hidráulico ok?", bloco: "Hidráulica e Lubrificação" },
    { id: "hid_abastecido", texto: "Veículo está abastecido?", bloco: "Hidráulica e Lubrificação" },

    { id: "tras_com_residuo", texto: "Veículo está com resíduo?", bloco: "Traseira/Carga" },
    { id: "tras_lavado", texto: "Veículo lavado e lubrificado?", bloco: "Traseira/Carga" },
    { id: "tras_placas", texto: "Placa dianteira e traseira ok?", bloco: "Traseira/Carga" },
    { id: "tras_placa_lacre", texto: "Placa traseira com lacre ok?", bloco: "Traseira/Carga" },

    { id: "lant_parachoque", texto: "Para-choque dianteiro e traseiro ok?", bloco: "Lanternagem" },
    { id: "lant_paralamas", texto: "Para-lamas dianteiros e traseiros ok?", bloco: "Lanternagem" },
    { id: "lant_estribo", texto: "Estribo e ponteira (LD/LE) ok?", bloco: "Lanternagem" },

    { id: "eletr_buzina", texto: "Buzina funcionando?", bloco: "Parte Elétrica" },
    { id: "eletr_instrumentos", texto: "Instrumentos do painel ok?", bloco: "Parte Elétrica" },
    { id: "eletr_combustivel", texto: "Marcador de combustível ok?", bloco: "Parte Elétrica" },
    { id: "eletr_farois", texto: "Faróis baixo e alto funcionando?", bloco: "Parte Elétrica" },
    { id: "eletr_pisca_alerta", texto: "Pisca-alerta dianteiro/traseiro ok?", bloco: "Parte Elétrica" },
    { id: "eletr_setas", texto: "Setas dianteiras e traseiras ok?", bloco: "Parte Elétrica" },
    { id: "eletr_lanternas_tras", texto: "Lanternas traseiras ok?", bloco: "Parte Elétrica" },
    { id: "eletr_lampada_re", texto: "Lâmpadas de ré ok?", bloco: "Parte Elétrica" },
    { id: "eletr_giroflex", texto: "Giroflex ok?", bloco: "Parte Elétrica" },
    { id: "eletr_alerta_sonoro", texto: "Alerta sonoro de ré ok?", bloco: "Parte Elétrica" },

    { id: "pneu_careca", texto: "Algum pneu careca? Qual?", bloco: "Pneus" },
    { id: "pneu_avariado", texto: "Algum pneu avariado? Qual?", bloco: "Pneus" },

    { id: "manut_preventiva", texto: "Necessita recolher para manutenção preventiva?", bloco: "Recolher Veículo p/ Manutenção" },
    { id: "manut_corretiva", texto: "Necessita recolher para manutenção corretiva?", bloco: "Recolher Veículo p/ Manutenção" }
  ];
}

// Itens do formulário físico "CHECKLIST DE VEÍCULO LEVE — VISTORIA DE SAÍDA DA GARAGEM" (Saveiro e
// demais veículos leves/administrativos).
function _itensVeiculoLevePorBloco() {
  return [
    { id: "doc_crlv", texto: "CRLV do veículo ok?", bloco: "Documentação" },
    { id: "doc_cnh", texto: "CNH do condutor válida?", bloco: "Documentação" },
    { id: "doc_autorizacao", texto: "Autorização de uso ok?", bloco: "Documentação" },
    { id: "doc_seguro", texto: "Comprovante de seguro ok?", bloco: "Documentação" },

    { id: "pneu_pressao", texto: "Pressão dos pneus ok?", bloco: "Pneus" },
    { id: "pneu_conservacao", texto: "Estado de conservação dos pneus ok?", bloco: "Pneus" },
    { id: "pneu_estepe", texto: "Estepe calibrado?", bloco: "Pneus" },
    { id: "pneu_ferramentas", texto: "Ferramentas e macaco ok?", bloco: "Pneus" },

    { id: "abast_nivel", texto: "Nível de combustível ok?", bloco: "Abastecimento" },
    { id: "abast_tampa", texto: "Tampa do tanque ok?", bloco: "Abastecimento" },
    { id: "abast_vazamento", texto: "Sem vazamentos de combustível?", bloco: "Abastecimento" },

    { id: "motor_oleo", texto: "Nível de óleo do motor ok?", bloco: "Motor" },
    { id: "motor_agua", texto: "Nível de água do radiador ok?", bloco: "Motor" },
    { id: "motor_freio", texto: "Nível do fluido de freio ok?", bloco: "Motor" },
    { id: "motor_direcao", texto: "Nível do fluido de direção hidráulica ok?", bloco: "Motor" },
    { id: "motor_vazamento", texto: "Sem vazamentos aparentes no motor?", bloco: "Motor" },

    { id: "ilum_farois", texto: "Faróis alto e baixo ok?", bloco: "Iluminação e Sinalização" },
    { id: "ilum_posicao", texto: "Luzes de posição ok?", bloco: "Iluminação e Sinalização" },
    { id: "ilum_freio", texto: "Luzes de freio ok?", bloco: "Iluminação e Sinalização" },
    { id: "ilum_setas", texto: "Setas dianteiras e traseiras ok?", bloco: "Iluminação e Sinalização" },
    { id: "ilum_re", texto: "Luz de ré ok?", bloco: "Iluminação e Sinalização" },
    { id: "ilum_placa", texto: "Luz de placa ok?", bloco: "Iluminação e Sinalização" },

    { id: "freio_pedal", texto: "Funcionamento do pedal de freio ok?", bloco: "Freios" },
    { id: "freio_mao", texto: "Freio de mão ok?", bloco: "Freios" },
    { id: "freio_vazamento", texto: "Sem vazamentos no sistema de freio?", bloco: "Freios" },

    { id: "limp_palhetas", texto: "Palhetas do limpador ok?", bloco: "Limpadores e Vidros" },
    { id: "limp_reservatorio", texto: "Reservatório de água do limpador ok?", bloco: "Limpadores e Vidros" },
    { id: "limp_parabrisa", texto: "Para-brisa, vidros e retrovisores ok?", bloco: "Limpadores e Vidros" },
    { id: "limp_retrovisores_ajuste", texto: "Retrovisores ajustados e firmes?", bloco: "Limpadores e Vidros" },

    { id: "int_cintos", texto: "Cintos de segurança ok?", bloco: "Interior e Exterior" },
    { id: "int_painel", texto: "Painel e luzes do painel ok?", bloco: "Interior e Exterior" },
    { id: "int_buzina", texto: "Buzina ok?", bloco: "Interior e Exterior" },
    { id: "int_banco", texto: "Banco e ajustes ok?", bloco: "Interior e Exterior" },
    { id: "int_carroceria", texto: "Carroceria e pintura ok?", bloco: "Interior e Exterior" },
    { id: "int_odor", texto: "Sem odor ou sinais de problemas?", bloco: "Interior e Exterior" }
  ];
}

// Devolve { grupoChave, grupoNome, itensBasicos, itensEspecificos } prontos para montar o formulário.
function obterDefinicaoChecklistPorTipo(informacaoTexto) {
  var grupoChave = mapearGrupoChecklist(informacaoTexto);

  var grupos = {
    CARRO_ADMINISTRATIVO: {
      nome: "🚗 Carro Administrativo (modelo Veículo Leve)",
      itens: _itensVeiculoLevePorBloco()
    },
    VEICULO_LEVE: {
      nome: "🚐 Veículo Leve (Saveiro)",
      itens: _itensVeiculoLevePorBloco()
    },
    COMPACTADOR: {
      nome: "🗑️ Compactador / Caminhão de Coleta",
      itens: _itensCompactadorPorBloco()
    },
    CAMINHAO_ENTULHO: {
      nome: "🚛 Caminhão de Entulho / Caçamba",
      itens: [
        { id: "basculamento", texto: "Sistema de basculamento (PTO)" },
        { id: "lona_cobertura", texto: "Lona / cobertura de carga" },
        { id: "travas_cacamba", texto: "Travas da caçamba" },
        { id: "vazamento_hidraulico2", texto: "Vazamento hidráulico" },
        { id: "cinta_amarracao", texto: "Corrente / cinta de amarração" }
      ]
    },
    TRITURADOR: {
      nome: "🌳 Triturador",
      itens: [
        { id: "laminas_corte", texto: "Lâminas / facas de corte" },
        { id: "embreagem_triturador", texto: "Sistema de embreagem do triturador" },
        { id: "protetor_seguranca", texto: "Protetor de segurança" },
        { id: "oleo_triturador", texto: "Nível de óleo do triturador" }
      ]
    },
    MAQUINA_PESADA: {
      nome: "🚜 Máquina Pesada / Trator",
      itens: [
        { id: "lanca_cacamba", texto: "Lança e caçamba" },
        { id: "estabilizadores", texto: "Estabilizadores" },
        { id: "sistema_hidraulico", texto: "Sistema hidráulico" },
        { id: "rodado_esteira", texto: "Rodado / esteira" },
        { id: "implemento", texto: "Implemento (roçadeira/lâmina)" }
      ]
    },
    VARREDEIRA: {
      nome: "🧹 Varredeira",
      itens: [
        { id: "escovas", texto: "Escovas de varrição" },
        { id: "aspersao_agua", texto: "Sistema de aspersão de água" },
        { id: "reservatorio_residuos", texto: "Reservatório de resíduos" },
        { id: "mangueiras_bicos", texto: "Mangueiras e bicos aspersores" }
      ]
    },
    MOTO: {
      nome: "🏍️ Moto",
      itens: [
        { id: "corrente_transmissao", texto: "Corrente e transmissão" },
        { id: "freios_moto", texto: "Freios dianteiro e traseiro" },
        { id: "espelhos", texto: "Espelhos retrovisores" },
        { id: "capacete", texto: "Capacete disponível e em bom estado" },
        { id: "suspensao", texto: "Suspensão" }
      ]
    },
    ONIBUS: {
      nome: "🚌 Ônibus / Transporte de Equipe",
      itens: [
        { id: "portas_escadas", texto: "Portas e escadas" },
        { id: "bancos_cintos", texto: "Bancos e cintos dos passageiros" },
        { id: "iluminacao_interna", texto: "Iluminação interna" },
        { id: "saida_emergencia", texto: "Saída de emergência" }
      ]
    },
    MAQUINA_PINTURA: {
      nome: "🎨 Máquina de Pintura",
      itens: [
        { id: "compressor_bomba", texto: "Compressor / bomba de pintura" },
        { id: "mangueiras_bicos_tinta", texto: "Mangueiras e bicos" },
        { id: "reservatorio_tinta", texto: "Nível do reservatório de tinta" },
        { id: "epis", texto: "EPIs disponíveis" }
      ]
    },
    REBOQUE: {
      nome: "🔗 Reboque",
      itens: [
        { id: "engate_pino", texto: "Engate e pino de segurança" },
        { id: "sinalizacao_traseira", texto: "Sinalização traseira" },
        { id: "pneus_reboque", texto: "Pneus do reboque" },
        { id: "estrutura_chassi", texto: "Estrutura / chassi" }
      ]
    },
    GERAL: {
      nome: "🔧 Equipamento Geral",
      itens: [
        { id: "estado_carroceria", texto: "Estado geral da carroceria" },
        { id: "vazamentos_visiveis", texto: "Vazamentos visíveis" },
        { id: "ruidos_anormais", texto: "Ruídos anormais" }
      ]
    }
  };

  var grupo = grupos[grupoChave] || grupos.GERAL;

  return {
    grupoChave: grupoChave,
    grupoNome: grupo.nome,
    itensBasicos: _itensBasicosChecklist(),
    itensEspecificos: grupo.itens
  };
}

// Normaliza a data/hora gravada na planilha para exibir sempre "dd/MM/yyyy" (sem hora), mesmo quando
// o Google Sheets autoconverte a célula de texto para um objeto Date ao ser lida via getValues().
function _normalizarDataChecklistBR(valor, tz) {
  if (Object.prototype.toString.call(valor) === "[object Date]") {
    return Utilities.formatDate(valor, tz, "dd/MM/yyyy");
  }
  var texto = String(valor || "");
  var m = texto.match(/^(\d{2}\/\d{2}\/\d{4})/);
  if (m) return m[1];
  return texto;
}

// Calcula a gravidade geral do checklist a partir das respostas ("OK" | "ATENCAO" | "CRITICO").
function _calcularGravidadeGeral(itensRespondidos) {
  var temCritico = false;
  var temAtencao = false;
  for (var i = 0; i < itensRespondidos.length; i++) {
    var r = String(itensRespondidos[i].resposta || "").toUpperCase();
    if (r === "CRITICO") temCritico = true;
    if (r === "ATENCAO") temAtencao = true;
  }
  if (temCritico) return "CRITICO";
  if (temAtencao) return "ATENCAO";
  return "OK";
}

function salvarChecklistDiario(dados) {
  try {
    if (!dados || !dados.prefixo || !dados.motorista || !dados.itens || !dados.itens.length) {
      return { sucesso: false, mensagem: "Erro: dados incompletos (veículo, motorista e itens são obrigatórios)." };
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var aba = ss.getSheetByName("CHECKLIST_DIARIO_MOTORISTAS");

    if (!aba) {
      aba = ss.insertSheet("CHECKLIST_DIARIO_MOTORISTAS");
      aba.appendRow(["ID", "DATA/HORA", "PREFIXO", "PLACA", "MOTORISTA", "GRUPO EQUIPAMENTO", "GRAVIDADE GERAL", "ITENS (JSON)", "OBSERVAÇÃO GERAL", "KM DE SAÍDA", "FOTOS AVARIAS (JSON)"]);
      aba.getRange("A1:K1").setFontWeight("bold").setBackground("#0f766e").setFontColor("#ffffff");
    } else {
      // Garante que a coluna FOTOS AVARIAS (JSON) existe em planilhas já criadas antes desta atualização
      if (aba.getRange("K1").getValue() === "") {
        aba.getRange("K1").setValue("FOTOS AVARIAS (JSON)").setFontWeight("bold").setBackground("#0f766e").setFontColor("#ffffff");
      }
    }

    var ultimaLinha = aba.getLastRow();
    var maxId = 0;
    if (ultimaLinha > 1) {
      var valoresId = aba.getRange(2, 1, ultimaLinha - 1, 1).getValues();
      for (var k = 0; k < valoresId.length; k++) {
        var idNum = parseInt(valoresId[k][0], 10);
        if (!isNaN(idNum) && idNum > maxId) maxId = idNum;
      }
    }
    var nextId = maxId > 0 ? maxId + 1 : 1;

    var gravidadeGeral = _calcularGravidadeGeral(dados.itens);
    var dataHora = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), "dd/MM/yyyy HH:mm");

    // Upload das fotos de avarias (se houver) para uma pasta dedicada no Drive
    var urlsFotosAvarias = [];
    if (dados.fotosAvariasBase64 && dados.fotosAvariasBase64.length) {
      var pastasAvarias = DriveApp.getFoldersByName("CHECKLIST_AVARIAS");
      var pastaAvarias = pastasAvarias.hasNext() ? pastasAvarias.next() : DriveApp.createFolder("CHECKLIST_AVARIAS");
      for (var f = 0; f < dados.fotosAvariasBase64.length; f++) {
        try {
          var fotoInfo = dados.fotosAvariasBase64[f];
          var blobFoto = Utilities.newBlob(Utilities.base64Decode(fotoInfo.base64), fotoInfo.mimeType, fotoInfo.nome);
          var fileFoto = pastaAvarias.createFile(blobFoto);
          fileFoto.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
          urlsFotosAvarias.push(fileFoto.getUrl());
        } catch (eFoto) {
          // Ignora falha isolada de uma foto e continua salvando o restante
        }
      }
    }

    aba.appendRow([
      nextId,
      dataHora,
      String(dados.prefixo).toUpperCase(),
      String(dados.placa || ""),
      String(dados.motorista),
      String(dados.grupoNome || ""),
      gravidadeGeral,
      JSON.stringify(dados.itens),
      String(dados.observacaoGeral || ""),
      String(dados.km || ""),
      JSON.stringify(urlsFotosAvarias)
    ]);

    return { sucesso: true, mensagem: "✅ Checklist salvo com sucesso!", gravidadeGeral: gravidadeGeral, id: nextId };
  } catch (erro) {
    return { sucesso: false, mensagem: "Erro inesperado ao salvar checklist: " + erro.message };
  }
}

function listarChecklistsDiarios() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var aba = ss.getSheetByName("CHECKLIST_DIARIO_MOTORISTAS");
    if (!aba) return [];

    var tz = ss.getSpreadsheetTimeZone();
    var dados = aba.getDataRange().getValues();
    var lista = [];
    for (var i = 1; i < dados.length; i++) {
      var itensJson = [];
      try { itensJson = JSON.parse(dados[i][7] || "[]"); } catch (eParse) { itensJson = []; }

      var fotosAvariasArr = [];
      try { fotosAvariasArr = JSON.parse(dados[i][10] || "[]"); } catch (eParseFotos) { fotosAvariasArr = []; }

      lista.push({
        id: dados[i][0],
        dataHora: _normalizarDataChecklistBR(dados[i][1], tz),
        prefixo: String(dados[i][2]),
        placa: String(dados[i][3]),
        motorista: String(dados[i][4]),
        grupoNome: String(dados[i][5]),
        gravidadeGeral: String(dados[i][6]),
        itens: itensJson,
        observacaoGeral: String(dados[i][8] || ""),
        km: String(dados[i][9] || ""),
        fotosAvarias: fotosAvariasArr
      });
    }
    return lista.reverse();
  } catch (erro) {
    return [];
  }
}
