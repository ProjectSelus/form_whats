// ================= CONFIGURAÇÕES =================
const PASTA_ORIGEM_ID = '131-qWXESEjRe7MdvF3BLyI_7eVaBCgxW';
const PASTA_DESTINO_ID = '1m1SaPeq9n2_quznRe8-dj9xybGD1J08W';
// =================================================


// 1. CRIA O MENU PERSONALIZADO NA BARRA DO GOOGLE SHEETS
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('⚡ Automação PM')
    .addItem('▶️ Gerar Escala (10 Dias)', 'gerarEscalaPatrulhamento')
    .addItem('⏱️ Subtrair 20 Minutos da Seleção', 'subtrair20Minutos')
    .addItem('▶️ Processar Arquivos Pendentes Agora', 'executarManualDaPlanilha')
    .addToUi();
}

// 2. FUNÇÃO PRINCIPAL: GERAR ESCALA
function gerarEscalaPatrulhamento() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var abaOrigem = ss.getSheetByName("Página2");
  var abaDestino = ss.getSheetByName("Página1");
  
  var ultimaLinhaOrigem = abaOrigem.getLastRow();

    // USA getDisplayValues() PARA PEGAR O TEXTO EXATO QUE ESTÁ NA TELA (Ex: "17:15")

  var dados = abaOrigem.getRange(2, 2, ultimaLinhaOrigem - 1, 17).getDisplayValues();
  
  var ultimaLinhaDestino = abaDestino.getLastRow();
  var dataAtual = new Date();
  
  // CONTINUAÇÃO DA ÚLTIMA LINHA DA PÁGINA 1
  if (ultimaLinhaDestino > 1) {
    var ultimaData = abaDestino.getRange(ultimaLinhaDestino, 1).getValue(); 
    if (ultimaData instanceof Date && !isNaN(ultimaData.getTime())) {
      dataAtual = new Date(ultimaData.getTime());
      dataAtual.setDate(dataAtual.getDate() + 1);
    }
  }
  
  var diasGerados = 0;
  var novasLinhas = [];

  while (diasGerados < 10) {
    var diaSemana = dataAtual.getDay();
    
    // Pula Sábado (6) e Domingo (0)
    if (diaSemana === 0 || diaSemana === 6) {
      dataAtual.setDate(dataAtual.getDate() + 1);
      continue;
    }
    
    var dataParaInserir = new Date(dataAtual.getTime()); 
    var isSexta = (diaSemana === 5);
    
    // ESCOLA FIXA DAS 06:30
    var escolasManha = [];
    for (var i = 0; i < dados.length; i++) {
      var escola = dados[i][0];
      var horaManha = dados[i][1];
      if (escola !== "" && horaManha.trim() === "06:30") {
        escolasManha.push(escola);
      }
    }
    
    var escolaFixa = (escolasManha.length > 0) 
      ? escolasManha[Math.floor(Math.random() * escolasManha.length)]
      : "MARECHAL RONDON";
      
    novasLinhas.push([dataParaInserir, "06:30", "POLICIAMENTO ESCOLAR " + escolaFixa, "", "=E2"]);

    // ESCOLAS ALEATÓRIAS (POR PESO)
    var eventosPossiveis = [];
    for (var i = 0; i < dados.length; i++) {
      var escola = dados[i][0];
      if (!escola) continue;
      
         // Colunas dos horários na matriz (C=1, E=3, G=5, I=7, K=9, M=11, O=13, Q=15)
      var colunasHora = [3, 5, 7, 9, 11, 13, 15]; 
      for (var c = 0; c < colunasHora.length; c++) {
        var idxHora = colunasHora[c];
        var idxPeso = idxHora + 1;
        var horaStr = dados[i][idxHora].trim();
        var peso = Number(dados[i][idxPeso]);
        
        if (horaStr !== "" && peso > 0) {
          // EXCEÇÃO: Adventista na Sexta-feira na coluna K (saida tarde 1, índice 9)
          if (escola.toUpperCase().indexOf("ADVENTISTA") !== -1 && isSexta && idxHora === 9) { 
            horaStr = "16:50";
          }
          
          for (var p = 0; p < peso; p++) {
            eventosPossiveis.push({ escola: escola, hora: horaStr });
          }
        }
      }
    }

      // Sorteia 2 eventos diferentes para o dia
    var sorteadosDoDia = [];
    for (var s = 0; s < 2; s++) {
      if (eventosPossiveis.length > 0) {
        var indexSorteado = Math.floor(Math.random() * eventosPossiveis.length);
        var eventoSorteado = eventosPossiveis[indexSorteado];
        
        sorteadosDoDia.push(eventoSorteado);

          // Impede de repetir a mesma escola e o mesmo horário no mesmo dia
        eventosPossiveis = eventosPossiveis.filter(function(item) {
          return !(item.escola === eventoSorteado.escola && item.hora === eventoSorteado.hora);
        });
      }
    }
    
    for (var s = 0; s < sorteadosDoDia.length; s++) {
      novasLinhas.push([dataParaInserir, sorteadosDoDia[s].hora, "POLICIAMENTO ESCOLAR " + sorteadosDoDia[s].escola, "", "=E2"]);
    }
    
    dataAtual.setDate(dataAtual.getDate() + 1);
    diasGerados++;
  }
    // 4. GRAVAÇÃO DOS DADOS
  if (novasLinhas.length > 0) {
    var rangeDestino = abaDestino.getRange(ultimaLinhaDestino + 1, 1, novasLinhas.length, 5); // Formata a coluna B como TEXTO antes de inserir os horários (garante visualização limpa)
    abaDestino.getRange(ultimaLinhaDestino + 1, 2, novasLinhas.length, 1).setNumberFormat("@"); // Insere os dados
    rangeDestino.setValues(novasLinhas);
      // Formata a coluna A como DATA
    abaDestino.getRange(ultimaLinhaDestino + 1, 1, novasLinhas.length, 1).setNumberFormat("dd/MM/yyyy");
  }
}

// 3. FUNÇÃO PARA SUBTRAIR 20 MINUTOS DAS CÉLULAS SELECIONADAS
function subtrair20Minutos() {
  var range = SpreadsheetApp.getActiveRange();
  if (!range) {
    SpreadsheetApp.getUi().alert("Selecione pelo menos uma célula com horário primeiro!");
    return;
  }
  
  var displayValues = range.getDisplayValues();
  var novosValores = [];
  var alterouAlgo = false;
  
  for (var r = 0; r < displayValues.length; r++) {
    var linha = [];
    for (var c = 0; c < displayValues[r].length; c++) {
      var valStr = displayValues[r][c].trim();
      
      // Procura formato de hora (ex: 06:30, 17:15, 6:30)
      var match = valStr.match(/^(\d{1,2}):(\d{2})$/);
      
      if (match) {
        var horas = parseInt(match[1], 10);
        var minutos = parseInt(match[2], 10);
        
        // Converte tudo para minutos totais e subtrai 20 minutos
        var totalMinutos = (horas * 60 + minutos) - 20;
        
        // Trata virada da meia-noite (ex: 00:10 vira 23:50)
        if (totalMinutos < 0) {
          totalMinutos += 1440;
        }
        
        var novasHoras = Math.floor(totalMinutos / 60);
        var novosMinutos = totalMinutos % 60;
        
        // Formata com zero à esquerda (HH:mm)
        var horaFormatada = (novasHoras < 10 ? "0" + novasHoras : novasHoras) + ":" + 
                            (novosMinutos < 10 ? "0" + novosMinutos : novosMinutos);
        
        linha.push(horaFormatada);
        alterouAlgo = true;
      } else {
        // Se não for horário, mantém o valor original
        linha.push(valStr);
      }
    }
    novosValores.push(linha);
  }
  
  if (alterouAlgo) {
    range.setNumberFormat("@"); // Garante formato de texto para não distorcer o horário
    range.setValues(novosValores);
  }
}

// Função chamada ao clicar no menu da Planilha
function executarManualDaPlanilha() {
  let ui = SpreadsheetApp.getUi();
  
  try {
    let resultado = processarArquivosTXT();
    ui.alert('Resultado da Automação', resultado, ui.ButtonSet.OK);
  } catch (erro) {
    ui.alert('Erro', 'Ocorreu uma falha: ' + erro.message, ui.ButtonSet.OK);
  }
}

// Função principal de processamento executada dentro da planilha
function processarArquivosTXT() {
  let pastaOrigem = DriveApp.getFolderById(PASTA_ORIGEM_ID);
  let pastaDestino = DriveApp.getFolderById(PASTA_DESTINO_ID);
  let planilha = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  // Busca os arquivos param_*.txt na raiz ou pastas
  let arquivosTxt = DriveApp.searchFiles("title contains 'param_' and mimeType = 'text/plain' and trashed = false");
  let processadosCount = 0;
  let logResultados = [];

  while (arquivosTxt.hasNext()) {
    let txtFile = arquivosTxt.next();
    let conteudo = txtFile.getBlob().getDataAsString();

    let oldNameMatch = conteudo.match(/OLD_NAME:\s*(.+)/);
    let newNameMatch = conteudo.match(/NEW_NAME:\s*(.+)/);
    let summaryMatch = conteudo.match(/SUMMARY:\s*(.+)/);

    if (oldNameMatch && newNameMatch && summaryMatch) {
      let oldName = oldNameMatch[1].trim();
      let newName = newNameMatch[1].trim();
      let summary = summaryMatch[1].trim();

      let documentosOriginais = pastaOrigem.getFilesByName(oldName);
      
      if (documentosOriginais.hasNext()) {
        let docOriginal = documentosOriginais.next();
        
        // 1. Renomeia e move
        docOriginal.setName(newName);
        docOriginal.moveTo(pastaDestino);
        let linkArquivo = docOriginal.getUrl();

        // 2. Data e Hora
        let dataAtual = new Date();
        let fuso = Session.getScriptTimeZone();
        let stringData = Utilities.formatDate(dataAtual, fuso, "dd/MM/yyyy");
        let stringHora = Utilities.formatDate(dataAtual, fuso, "HH:mm:ss");

        // 3. Salva nas Colunas: A: DATA | B: HORA | C: NOME/MENSAGEM | D: LINK | E: RESUMO/ATIVIDADE
        planilha.appendRow([stringData, stringHora, newName, linkArquivo, summary]);

        // 4. Manda o TXT para a lixeira
        txtFile.setTrashed(true);
        
        logResultados.push(`✅ '${oldName}' renomeado para '${newName}' e movido.`);
        processadosCount++;
      } else {
        logResultados.push(`⚠️ Arquivo '${oldName}' não foi encontrado na pasta de origem.`);
      }
    }
  }

  if (processadosCount === 0) {
    return "Nenhum arquivo TXT de automação pendente foi encontrado.";
  }

  return `Automação concluída! ${processadosCount} arquivo(s) processado(s):\n\n` + logResultados.join("\n");
}

// 4. SERVE OS DADOS PARA O SITE (API JSON) OU EXIBE A PÁGINA WEB
function doGet(e) {
  // Se for solicitado explicitamente o HTML (ex: no link do Apps Script com ?formato=html)
  if (e && e.parameter && e.parameter.formato === 'html') {
    return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('PMMS - 14ª CIPM | Solicitações de Serviços (SS)')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
  }

  // Padrão: Retorna os dados em JSON da planilha para o site GitHub Pages (https://relatorio.zerog.com.br/eventos)
  var dados = obterOcorrencias();
  return ContentService.createTextOutput(JSON.stringify(dados))
    .setMimeType(ContentService.MimeType.JSON);
}

// Função para ler o HTML direto do Google Drive (evita ter que reimplantar a cada alteração visual)
function lerArquivoDoDrive(nomeArquivo) {
  var arquivos = DriveApp.getFilesByName(nomeArquivo);
  while (arquivos.hasNext()) {
    var f = arquivos.next();
    if (!f.isTrashed()) {
      return f.getBlob().getDataAsString('UTF-8');
    }
  }
  throw new Error('Arquivo não encontrado no Google Drive: ' + nomeArquivo);
}

// 5. RETORNA AS OCORRÊNCIAS DA PLANILHA PARA O HTML
function obterOcorrencias() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var aba = ss.getSheetByName("Página1");
    if (!aba) {
      return { erro: "Aba 'Página1' não encontrada." };
    }
    
    var ultimaLinha = aba.getLastRow();
    if (ultimaLinha < 2) {
      return { ocorrencias: [] };
    }
    
    // Obtém Coluna A até E (Data, Hora, Mensagem, Arquivo, Atividade)
    var dados = aba.getRange(2, 1, ultimaLinha - 1, 5).getDisplayValues();
    var lista = [];
    
    for (var i = 0; i < dados.length; i++) {
      var dataStr = dados[i][0] ? dados[i][0].trim() : "";
      var horaStr = dados[i][1] ? dados[i][1].trim() : "";
      var mensagemStr = dados[i][2] ? dados[i][2].trim() : "";
      var arquivoStr = dados[i][3] ? dados[i][3].trim() : "";
      var atividadeStr = dados[i][4] ? dados[i][4].trim() : "";
      
      // Ignora linhas totalmente em branco
      if (!dataStr && !horaStr && !mensagemStr && !atividadeStr) {
        continue;
      }
      
      lista.push({
        data: dataStr,
        hora: horaStr,
        mensagem: mensagemStr,
        arquivo: arquivoStr,
        atividade: atividadeStr
      });
    }
    
    return { ocorrencias: lista };
  } catch (e) {
    return { erro: e.message };
  }
}