// ================= MOTOR DE ATESTADOS E DISCIPLINA (RH) =================

function salvarOcorrenciaRHBackend(tipo, dados) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var aba = ss.getSheetByName(tipo);
    if (!aba) return { sucesso: false, mensagem: "Aba " + tipo + " não encontrada." };

    var dataAtual = new Date();
    var idFmt = tipo.substring(0,3) + "-" + dataAtual.getTime().toString().slice(-6);
    var urlPdfOuImagem = "";

    // 1. UPLOAD DE ATESTADO (IMAGEM/PDF) SE EXISTIR
    if (tipo === "ATESTADOS" && dados.arquivoBase64) {
       var blobArquivo = Utilities.newBlob(Utilities.base64Decode(dados.arquivoBase64), dados.arquivoMimeType, dados.arquivoNome);
       var pastas = DriveApp.getFoldersByName("RH_DOCUMENTOS");
       var pastaDocs = pastas.hasNext() ? pastas.next() : DriveApp.createFolder("RH_DOCUMENTOS");
       var fileCriado = pastaDocs.createFile(blobArquivo);
       fileCriado.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
       urlPdfOuImagem = fileCriado.getUrl();
    }

    // 2. GERAR PDF OFICIAL (ADVERTÊNCIA / SUSPENSÃO)
    if (tipo === "ADVERTÊNCIA" || tipo === "SUSPENSÃO") {
       urlPdfOuImagem = gerarESalvarPdfDisciplinar(tipo, idFmt, dados);
    }

    // 3. PREPARAR A LINHA PARA A PLANILHA ESPECÍFICA
    var novaLinha = [];
    if (tipo === "ATESTADOS") {
      // Colunas: ID, FUNCIONÁRIO, CID/MOTIVO, DATA ATESTADO, DIAS, DATA RETORNO, IMAGEM, AFASTAMENTO, RETORNO AFAST., DIAS FALTA
      novaLinha = [
        idFmt, dados.funcionario.toUpperCase(), dados.motivo.toUpperCase(), 
        dados.dataAtestado, dados.diasAtestado, dados.dataRetornoAtestado, 
        urlPdfOuImagem, 
        dados.dataAfastamento, dados.dataRetornoAfastamento, dados.diasFalta
      ];
    } 
    else if (tipo === "ADVERTÊNCIA" || tipo === "SUSPENSÃO") {
      // Colunas: ID, DATA, FUNCIONÁRIO, CPF, ATO, TIPO (Adv/Susp), APLICADA POR, PDF
      var textoAtoCompleto = dados.motivoLegal + "\n" + dados.ato;
      novaLinha = [
        idFmt, Utilities.formatDate(dataAtual, Session.getScriptTimeZone(), "dd/MM/yyyy"),
        dados.funcionario.toUpperCase(), dados.cpf, textoAtoCompleto, tipo, 
        dados.aplicadaPor.toUpperCase(), urlPdfOuImagem
      ];
    }

    aba.appendRow(novaLinha);
    return { sucesso: true, mensagem: tipo + " registrado com sucesso!", urlPdf: urlPdfOuImagem };
  } catch (e) {
    return { sucesso: false, mensagem: "Erro ao salvar: " + e.toString() };
  }
}

// BUSCAR HISTÓRICO DO FUNCIONÁRIO PARA A TELA
function listarHistoricoFuncionario(nomeFuncionario) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var historico = { atestados: [], advertencias: [], suspensoes: [] };
  var nomeRef = String(nomeFuncionario).trim().toUpperCase();

  function buscarAba(nomeAba, arrayDestino, isAtestado) {
    var aba = ss.getSheetByName(nomeAba);
    if(!aba) return;
    var dados = aba.getDataRange().getValues();
    for(var i=1; i<dados.length; i++) {
      if(String(dados[i][1] || dados[i][2]).trim().toUpperCase() === nomeRef) {
        if(isAtestado) {
          arrayDestino.push({ linha: i+1, id: dados[i][0], motivo: dados[i][2], data: dados[i][3], dias: dados[i][4] });
        } else {
          arrayDestino.push({ linha: i+1, id: dados[i][0], data: dados[i][1], ato: dados[i][4], pdf: dados[i][7] });
        }
      }
    }
  }

  buscarAba("ATESTADOS", historico.atestados, true);
  buscarAba("ADVERTÊNCIA", historico.advertencias, false);
  buscarAba("SUSPENSÃO", historico.suspensoes, false);

  return historico;
}

// EXCLUIR REGISTRO DE FORMA SEGURA (PELO ID ÚNICO E NÃO PELA LINHA)
function excluirOcorrenciaRHBackend(tipo, idRegistro) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var aba = ss.getSheetByName(tipo);
    if (!aba) return { sucesso: false, mensagem: "Aba não encontrada." };

    var dados = aba.getDataRange().getValues();
    
    // Procura de baixo para cima (mais rápido) para encontrar o ID Exato
    for (var i = dados.length - 1; i >= 1; i--) {
      if (String(dados[i][0]).trim() === String(idRegistro).trim()) {
        aba.deleteRow(i + 1); // Exclui a linha exata onde o ID foi encontrado
        return { sucesso: true, mensagem: "🗑️ Registro " + idRegistro + " excluído permanentemente!" };
      }
    }
    
    return { sucesso: false, mensagem: "O registro já não existe na planilha." };
  } catch(e) {
    return { sucesso: false, mensagem: "Erro ao excluir: " + e.toString() };
  }
}

// GERADOR DE PDF (NORMA ABNT, 1 PÁGINA GARANTIDA, LINHAS CORRIGIDAS)
function gerarESalvarPdfDisciplinar(tipo, id, dados) {
  // Formatar a data em Português
  var dataObj = new Date();
  var meses = ["janeiro", "fevereiro", "março", "abril", "maio", "junho", "julho", "agosto", "setembro", "outubro", "novembro", "dezembro"];
  var dataPtBr = dataObj.getDate() + " de " + meses[dataObj.getMonth()] + " de " + dataObj.getFullYear();
  
  // Formatar o CPF com a máscara
  var cpfLimpo = String(dados.cpf).replace(/\D/g, '');
  var cpfFormatado = cpfLimpo.length === 11 ? cpfLimpo.replace(/(\d{3})(\d{3})(\d{3})(\d{2})/, "$1.$2.$3-$4") : dados.cpf;

  var titulo = tipo === "ADVERTÊNCIA" ? "AVISO DE ADVERTÊNCIA" : "AVISO DE SUSPENSÃO";
  var acaoTxt = tipo === "ADVERTÊNCIA" ? "Advertido" : "Suspenso";
  
  // Fonte ajustada para 13px e margens controladas para caber exatamente em 1 folha
  var html = `
    <div style="font-family: Arial, sans-serif; font-size: 13px; padding: 15px 40px; color: #000;">
        
        <div style="text-align: center; margin-bottom: 35px;">
            <div style="font-size: 48px; font-weight: 900; font-family: 'Arial Black', Impact, sans-serif; color: #000; line-height: 0.9; letter-spacing: -2px;">Inovar</div>
            <div style="font-size: 10px; font-weight: bold; font-family: Arial, sans-serif; color: #84cc16; letter-spacing: 1px;">SOLUÇÕES AMBIENTAIS LTDA</div>
        </div>
        
        <h3 style="text-align: center; margin-bottom: 25px; font-size: 15px; text-decoration: underline;">${titulo}</h3>
        
        <p style="margin-bottom: 4px;"><strong>Ao</strong></p>
        <p style="margin-bottom: 4px;"><strong>Senhor(a):</strong> ${dados.funcionario.toUpperCase()}</p>
        <p style="margin-bottom: 20px;"><strong>CPF:</strong> ${cpfFormatado}</p>
        
        <p style="text-align: justify; line-height: 1.4; margin-bottom: 15px;">
           Em conformidade com a Consolidação das Leis Trabalhistas (CLT). Fica V. As. ${acaoTxt} pela falta abaixo discriminada:
        </p>
        
        <p style="text-align: justify; line-height: 1.4; margin-bottom: 15px;">
           <strong>${dados.motivoLegal || 'Art. 482 da CLT'}</strong>
        </p>
        
        <p style="text-align: justify; line-height: 1.4; margin-bottom: 20px;">
           <strong>Ato:</strong> ${dados.ato}
        </p>
        
        <p style="text-align: justify; line-height: 1.4; margin-bottom: 30px;">
           Esperamos que tome as necessárias precauções, a fim de que não se repita a irregularidade acima discriminada, como também aproveitamos para esclarecer-lhe que a repetição, ou prática de outras faltas previstas em nossos regulamentos poderão contribuir desfavoravelmente em seu progresso nesta empresa, além de poder acarretar-lhe penalidades mais severas, conforme preceituam as disposições do artigo 482 e suas alíneas da Consolidação das Leis Trabalhistas.
        </p>
        
        <p style="margin-bottom: 45px;">João Pessoa, ${dataPtBr}.</p>
        
        <div style="border-top: 1px solid #000; width: 400px; padding-top: 5px; margin-bottom: 30px; font-weight: bold;">
            Inovar Soluções Ambientais LTDA
        </div>
        
        <div>
            <p style="margin-bottom: 40px;">Ciente em: ___/___/______.</p>
            <div style="border-top: 1px solid #000; width: 400px; padding-top: 5px; margin-bottom: 30px;">
                ${dados.funcionario.toUpperCase()}
            </div>
        </div>
        
        <div>
            <p style="font-weight: bold; margin-bottom: 40px;">Testemunhas:</p>
            <div style="border-top: 1px solid #000; width: 400px; padding-top: 5px; margin-bottom: 40px;">1. CPF: </div>
            <div style="border-top: 1px solid #000; width: 400px; padding-top: 5px;">2. CPF: </div>
        </div>

    </div>
  `;

  var blob = HtmlService.createHtmlOutput(html).getAs('application/pdf');
  blob.setName(tipo + "_" + dados.funcionario.split(" ")[0] + "_" + id + ".pdf");
  
  var pastas = DriveApp.getFoldersByName("RH_DOCUMENTOS");
  var pasta = pastas.hasNext() ? pastas.next() : DriveApp.createFolder("RH_DOCUMENTOS");
  var arquivo = pasta.createFile(blob);
  arquivo.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  
  return arquivo.getUrl();
}
