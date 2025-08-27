/**
 * @OnlyCurrentDoc
 *
 * SCRIPT PARA ENVIAR ALERTAS AUTOMÁTICOS POR E-MAIL
 *
 * Este script foi projetado para ler uma planilha do Google e enviar um
 * e-mail de alerta quando encontrar datas que estão próximas de um vencimento
 * pré-configurado (por exemplo, 30 ou 15 dias).
 *
 * A lógica e as funções foram mantidas conforme a estrutura original.
 * Para configurar, altere apenas as variáveis na seção "CONFIGURAÇÕES DO USUÁRIO".
 *
 */

// ===================================================================================
// FUNÇÃO PRINCIPAL
// Esta é a função que realiza todo o trabalho de verificação e envio de e-mail.
// ===================================================================================
function enviarAlertaNC() {

  // --- CONFIGURAÇÕES DO USUÁRIO (ALTERE OS VALORES ABAIXO) ---

  // 1. Informe o nome exato da aba que o script deve monitorar.
  var NOME_DA_ABA = "AE - Não Conformidades";

  // 2. Informe a partir de qual linha começam os seus dados (ignorando o cabeçalho).
  var LINHA_INICIAL = 5;

  // 3. Defina os dias de antecedência para o envio do alerta.
  // Ex: [15, 30] enviará um alerta quando faltarem exatamente 15 ou 30 dias.
  var DIAS_PARA_ALERTA = [15, 30];

  // 4. Configure o destinatário e o assunto do e-mail.
  var EMAIL_DESTINATARIO = "seu_email@exemplo.com";
  var ASSUNTO_DO_EMAIL = "Alerta de Itens - Vencendo em 30 ou 15 Dias";
  var TITULO_DO_EMAIL = "📌 Alerta de Não Conformidades"; // Título que aparece dentro do corpo do e-mail

  // 5. Mapeie as colunas da sua planilha.
  // Lembre-se que a contagem começa em 0 (Coluna A = 0, Coluna B = 1, C = 2, ...).
  var colNumero = 0;       // Coluna que contém o ID ou número do item.
  var colDesvio = 1;       // Coluna com a descrição do item/desvio.
  var colAcao = 10;        // Coluna com a descrição da ação corretiva.
  var colResponsavel = 13; // Coluna que informa o responsável.

  // 6. Mapeie as COLUNAS DE DATA que você deseja monitorar.
  var colPrazoAcao = 11;   // Coluna com a data do "Prazo para Ação".
  var colAbrangencia = 14; // Coluna com a data da "Abrangência".
  // Se tiver mais colunas de data, crie novas variáveis para elas aqui.

  // --- FIM DAS CONFIGURAÇÕES ---


  // O script começa acessando a planilha ativa.
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  // Em seguida, acessa a aba específica que foi configurada acima.
  var sheet = ss.getSheetByName(NOME_DA_ABA);

  // Se a aba não for encontrada, o script registra um erro e para a execução.
  if (!sheet) {
    Logger.log("Aba não encontrada: '" + NOME_DA_ABA + "'");
    return;
  }

  // O script pega todos os dados da aba e os armazena na variável 'dados'.
  var dados = sheet.getDataRange().getValues();

  // Verifica se a planilha tem dados suficientes para analisar, baseado na linha inicial configurada.
  if (dados.length < LINHA_INICIAL) {
    Logger.log("Planilha sem dados para analisar além do cabeçalho.");
    return;
  }

  // Cria um objeto de data com o dia de hoje, zerando as horas, para garantir uma comparação precisa de dias inteiros.
  var hoje = new Date();
  hoje.setHours(0, 0, 0, 0);

  // Inicializa variáveis para armazenar as mensagens de alerta, as datas inválidas e um contador.
  var mensagens = [];
  var invalidDates = [];
  var matchesCount = 0;

  // Inicia um laço (loop) para percorrer cada linha da planilha, começando da 'LINHA_INICIAL'.
  // O "-1" ajusta o número da linha para o índice do array (que começa em 0).
  for (var i = LINHA_INICIAL - 1; i < dados.length; i++) {
    var linha = dados[i]; // Pega a linha atual que está sendo processada.

    // Extrai os valores das células da linha atual, com base nas colunas configuradas.
    var numero = linha[colNumero];
    var desvio = linha[colDesvio];
    var acao = linha[colAcao];
    var prazoAcaoRaw = linha[colPrazoAcao]; // Pega o valor "bruto" da célula de data.
    var abrangenciaRaw = linha[colAbrangencia];
    var responsavel = linha[colResponsavel];

    // Tenta converter os valores "brutos" das células em datas válidas usando a função auxiliar 'parseDateFromCell'.
    var prazoAcao = parseDateFromCell(prazoAcaoRaw);
    var abrangencia = parseDateFromCell(abrangenciaRaw);

    // Se a célula de data tinha um valor mas não pôde ser convertida, armazena como data inválida.
    if (prazoAcaoRaw && !prazoAcao) invalidDates.push({ row: i + 1, col: colPrazoAcao + 1, value: prazoAcaoRaw });
    if (abrangenciaRaw && !abrangencia) invalidDates.push({ row: i + 1, col: colAbrangencia + 1, value: abrangenciaRaw });

    // --- Início das Verificações de Data ---

    // 1. Verifica a data de "Prazo de Ação".
    // Se a data for válida e estiver dentro do intervalo de dias para alerta...
    if (prazoAcao && isWithinDaysRange(prazoAcao, hoje, DIAS_PARA_ALERTA)) {
      // Cria a mensagem formatada para o e-mail e a adiciona na lista de mensagens.
      mensagens.push(formatMensagem(numero, desvio, acao, prazoAcao, "Prazo ação", responsavel, prazoAcao));
      matchesCount++; // Incrementa o contador de alertas encontrados.
    }

    // 2. Verifica a data de "Abrangência".
    // Se a data for válida e estiver dentro do intervalo de dias para alerta...
    if (abrangencia && isWithinDaysRange(abrangencia, hoje, DIAS_PARA_ALERTA)) {
      // Cria a mensagem formatada para o e-mail e a adiciona na lista de mensagens.
      mensagens.push(formatMensagem(numero, desvio, acao, abrangencia, "Abrangência", responsavel, abrangencia));
      matchesCount++; // Incrementa o contador de alertas encontrados.
    }

    // Para adicionar novas verificações de data, copie um dos blocos "if" acima e ajuste as variáveis.
  }

  // Após percorrer todas as linhas, o script registra um resumo da execução nos logs.
  Logger.log("Linhas processadas: " + (dados.length - (LINHA_INICIAL - 1)));
  Logger.log("Alertas encontrados: " + matchesCount);
  if (invalidDates.length > 0) {
    Logger.log("Datas inválidas encontradas (linha, coluna, valor):");
    invalidDates.forEach(x => Logger.log(x.row + " / " + x.col + " -> " + x.value));
  }

  // Se houver pelo menos uma mensagem de alerta para enviar...
  if (mensagens.length > 0) {
    // Monta o corpo do e-mail usando HTML para criar uma tabela bem formatada.
    var corpoEmail = `
    <!DOCTYPE html>
    <html lang="pt-br">
    <head>
      <meta charset="UTF-8">
      <meta name="viewport" content="width=device-width, initial-scale=1.0">
      <title>Alertas de Vencimento</title>
      <style>
        body { font-family: Arial, sans-serif; margin: 20px; background-color: #f4f4f9; }
        table { width: 100%; border-collapse: collapse; margin-top: 20px; }
        table, th, td { border: 1px solid #ccc; }
        th, td { padding: 10px; text-align: center; }
        th { background-color: #005860; color: white; }
        tr:nth-child(even) { background-color: #f2f2f2; }
        tr:hover { background-color: #ddd; }
      </style>
    </head>
    <body>
      <h1>${TITULO_DO_EMAIL}</h1>
      <p>Os itens a seguir estão próximos do vencimento (${DIAS_PARA_ALERTA.join(' ou ')} dias):</p>
      <table>
        <thead>
          <tr>
            <th>NC</th>
            <th>Desvio</th>
            <th>Ação</th>
            <th>Status</th>
            <th>Data</th>
            <th>Data Limite</th>
            <th>Responsável</th>
          </tr>
        </thead>
        <tbody>
          ${mensagens.join('')}
        </tbody>
      </table>
    </body>
    </html>
    `;

    // Envia o e-mail usando o serviço MailApp do Google.
    MailApp.sendEmail({
      to: EMAIL_DESTINATARIO,
      subject: ASSUNTO_DO_EMAIL,
      htmlBody: corpoEmail
    });

    Logger.log("E-mail enviado para: " + EMAIL_DESTINATARIO);
  } else {
    // Caso nenhum alerta seja encontrado, apenas registra uma mensagem nos logs.
    Logger.log("Nenhum alerta para enviar hoje.");
  }
}

// ===================================================================================
// FUNÇÕES AUXILIARES
// Estas funções dão suporte à função principal. Não é necessário alterá-las.
// ===================================================================================

/**
 * Converte o conteúdo de uma célula em um objeto de Data do JavaScript.
 * Funciona com datas já formatadas, números seriais do Google Sheets e texto (dd/mm/yyyy ou yyyy-mm-dd).
 * @param {any} cell O valor da célula.
 * @return {Date|null} Um objeto de Data ou null se a conversão falhar.
 */
function parseDateFromCell(cell) {
  if (!cell && cell !== 0) return null;
  var d = null;
  if (Object.prototype.toString.call(cell) === "[object Date]" && !isNaN(cell.getTime())) d = new Date(cell);
  else if (typeof cell === 'number') d = new Date(Math.round((cell - 25569) * 86400 * 1000));
  else if (typeof cell === 'string') {
    var s = cell.trim();
    var match = s.match(/^(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})$/);
    if (match) d = new Date(+match[1], +match[2] - 1, +match[3]);
    else {
      match = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/);
      if (match) d = new Date(+match[3], +match[2] - 1, +match[1]);
      else {
        var parsed = Date.parse(s);
        if (!isNaN(parsed)) d = new Date(parsed);
      }
    }
  }
  if (d) d.setHours(0, 0, 0, 0);
  return d && !isNaN(d.getTime()) ? d : null;
}

/**
 * Verifica se a diferença de dias entre duas datas está contida em um array.
 * @param {Date} dateAlvo A data futura a ser verificada.
 * @param {Date} hoje A data de hoje.
 * @param {Array<number>} diasArray Um array de números, ex: [15, 30].
 * @return {boolean} Verdadeiro se a diferença de dias for uma das opções do array.
 */
function isWithinDaysRange(dateAlvo, hoje, diasArray) {
  if (!Array.isArray(diasArray)) return false;
  var diff = Math.round((dateAlvo - hoje) / (1000 * 60 * 60 * 24));
  return diasArray.includes(diff);
}

/**
 * Formata um objeto de Data para uma string no formato "dd/MM/yyyy".
 * @param {Date} d O objeto de data a ser formatado.
 * @return {string} A data formatada como texto.
 */
function formatDateStr(d) {
  var tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
  return Utilities.formatDate(new Date(d), tz, "dd/MM/yyyy");
}

/**
 * Garante que o valor a ser inserido no HTML seja um texto, evitando erros com valores nulos.
 * @param {any} v O valor a ser convertido para texto.
 * @return {string} O valor como texto ou uma string vazia.
 */
function safeText(v) {
  return (v === null || v === undefined) ? "" : String(v);
}

/**
 * Monta uma linha (<tr>) da tabela HTML para o corpo do e-mail com os dados do alerta.
 * @return {string} Uma string contendo a linha HTML completa da tabela.
 */
function formatMensagem(numero, desvio, acao, data, tipo, responsavel, dataAlerta) {
  return `<tr>
            <td>${safeText(numero)}</td>
            <td>${safeText(desvio)}</td>
            <td>${safeText(acao)}</td>
            <td>${tipo}</td>
            <td>${formatDateStr(data)}</td>
            <td>${dataAlerta ? formatDateStr(dataAlerta) : ""}</td>
            <td>${safeText(responsavel)}</td>
          </tr>`;
}

// ===================================================================================
// FUNÇÕES DE TESTE E VALIDAÇÃO
// Use estas funções no editor do Apps Script para executar testes manuais.
// ===================================================================================

/**
 * Função para executar o script manualmente a partir do editor.
 * Limpa os logs antigos antes de rodar a função principal.
 */
function enviarAlertaNC_manual() {
  Logger.clear();
  enviarAlertaNC();
}

/**
 * Função auxiliar para validar se os valores em uma coluna são datas válidas.
 * Ao executar, ela perguntará qual coluna da aba ATIVA você deseja verificar.
 */
function validarDatas_manual() {
  // Pede ao usuário para inserir o número da coluna a ser validada.
  var ui = SpreadsheetApp.getUi();
  var result = ui.prompt(
    'Validador de Datas',
    'Por favor, insira o NÚMERO da coluna que você quer validar (ex: A=1, B=2):',
    ui.ButtonSet.OK_CANCEL);

  var button = result.getSelectedButton();
  var text = result.getResponseText();

  // Se o usuário clicou OK e inseriu um número...
  if (button == ui.Button.OK) {
    var colunaDatas = parseInt(text);
    if (isNaN(colunaDatas) || colunaDatas <= 0) {
      ui.alert("Número de coluna inválido.");
      return;
    }
    
    // Pega a aba atualmente aberta pelo usuário.
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const lastRow = sheet.getLastRow();
    const inicio = 2; // Assume que os dados começam na linha 2 por padrão para validação.
    let erros = [];

    // Percorre a coluna especificada e verifica cada célula.
    for (let i = inicio; i <= lastRow; i++) {
      const valor = sheet.getRange(i, colunaDatas).getValue();
      // Se a célula não estiver vazia e não for uma data válida...
      if (valor && !parseDateFromCell(valor)) {
        erros.push(`Linha ${i}, valor inválido: ${valor}`);
      }
    }

    // Exibe o resultado para o usuário.
    if (erros.length > 0) {
      Logger.log("Datas inválidas encontradas:");
      Logger.log(erros.join("\n"));
      ui.alert("Problemas Encontrados!", "Foram encontradas datas em formato inválido. Veja os detalhes no menu 'Registros de execução'.", ui.ButtonSet.OK);
    } else {
      Logger.log("Todas as datas estão válidas ✅");
      ui.alert("Sucesso!", "Todas as datas na coluna " + colunaDatas + " parecem ser válidas.", ui.ButtonSet.OK);
    }
  }
}