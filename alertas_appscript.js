/**************************************************************************************************
 * @fileoverview Scripts para automatizar o envio de e-mails de alerta com base em datas de 
 * vencimento em uma planilha Google Sheets. Ideal para gestão de tarefas, conformidade (compliance)
 * e controle de documentos.
 *
 * @version 1.1
 * @author [Maikon_Silva]
 * @license MIT
 **************************************************************************************************/

// ================================================
// CONFIGURAÇÃO GLOBAL
// ================================================
// Coloque aqui o ID da sua planilha. Ex: "1a2b3c4d5e6f7g8h9i0j"
const SPREADSHEET_ID = "ID_DA_SUA_PLANILHA_AQUI"; 

// E-mails para onde os alertas gerais serão enviados.
const DEFAULT_DESTINATION_EMAIL = "email1@exemplo.com, email2@exemplo.com";
const DEFAULT_BCC_EMAIL = "seu_email_para_copia@exemplo.com";


/**************************************************************************************************
 * ALERTA 1: MONITORAMENTO DE NÃO CONFORMIDADES (TIPO A)
 **************************************************************************************************/
const CONFIG_NC_A = {
  SHEET_NAME: "NaoConformidades_TipoA", // Nome da aba na planilha
  DESTINATION_EMAIL: DEFAULT_DESTINATION_EMAIL,
  BCC_EMAIL: DEFAULT_BCC_EMAIL,
  ALERT_DAYS: [15, 30], // Dias de antecedência para alertar
  START_ROW: 5, // A partir de qual linha os dados começam (Linha 5 = 5)
  COLUMNS: {
    id: 0,        // Coluna A
    description: 2, // Coluna C
    action: 11,     // Coluna L
    responsible: 14,// Coluna O
    // Datas a serem verificadas
    dates_to_check: [
      { name: "Prazo Ação", index: 12 }, // Coluna M
      { name: "Abrangência", index: 15 },// Coluna P
      { name: "Eficácia", index: 18 }   // Coluna S
    ],
    // Coluna para verificar se a tarefa de eficácia foi concluída
    effectiveness_completed_check: {
      column_index: 19, // Coluna T
      completed_marker: 'X' // Texto/marcador que indica a conclusão
    }
  }
};

/**
 * Função principal que dispara os alertas para Não Conformidades do Tipo A.
 */
function runNCTypeA_Alerts() {
  Logger.log("Iniciando verificação de alertas para Não Conformidades Tipo A...");
  const alertGenerator = new AlertGenerator(SPREADSHEET_ID, CONFIG_NC_A);
  alertGenerator.processAndSendEmail(
    "Alerta de Não Conformidades (Tipo A)",
    "As seguintes não conformidades estão próximas do vencimento:"
  );
}


/**************************************************************************************************
 * ALERTA 2: MONITORAMENTO DE NÃO CONFORMIDADES (TIPO B)
 **************************************************************************************************/
const CONFIG_NC_B = {
  SHEET_NAME: "NaoConformidades_TipoB",
  DESTINATION_EMAIL: DEFAULT_DESTINATION_EMAIL,
  BCC_EMAIL: DEFAULT_BCC_EMAIL,
  ALERT_DAYS: [15, 30],
  START_ROW: 5,
  COLUMNS: {
    id: 0,        // Coluna A
    description: 1, // Coluna B
    action: 10,     // Coluna K
    responsible: 13,// Coluna N
    dates_to_check: [
      { name: "Prazo Ação", index: 11 }, // Coluna L
      { name: "Abrangência", index: 14 },// Coluna O
      { name: "Eficácia", index: 17 }   // Coluna R
    ],
    effectiveness_completed_check: {
      column_index: 18, // Coluna S
      completed_marker: 'X'
    }
  }
};

/**
 * Função principal que dispara os alertas para Não Conformidades do Tipo B.
 */
function runNCTypeB_Alerts() {
  Logger.log("Iniciando verificação de alertas para Não Conformidades Tipo B...");
  const alertGenerator = new AlertGenerator(SPREADSHEET_ID, CONFIG_NC_B);
  alertGenerator.processAndSendEmail(
    "Alerta de Não Conformidades (Tipo B)",
    "As seguintes não conformidades estão próximas do vencimento:"
  );
}


/**************************************************************************************************
 * ALERTA 3: MONITORAMENTO DE DOCUMENTOS (LISTA MESTRA)
 **************************************************************************************************/
const CONFIG_DOCS = {
  SHEET_NAME: "ListaMestra_Documentos",
  DESTINATION_EMAIL: DEFAULT_DESTINATION_EMAIL,
  BCC_EMAIL: DEFAULT_BCC_EMAIL,
  ALERT_DAYS: [7, 14, 30], // Alerta com 7, 14 ou 30 dias de antecedência
  START_ROW: 2,
  COLUMNS: {
    id: 0,          // Coluna A
    description: 3,   // Coluna D
    responsible: null, // Não há coluna de responsável neste exemplo
    action: 5,        // Coluna F (usando como "Documento")
    dates_to_check: [
      { name: "Vencimento", index: 6 } // Coluna G
    ],
    effectiveness_completed_check: null // Sem verificação de conclusão para este alerta
  }
};

/**
 * Função principal que dispara os alertas para a Lista Mestra de Documentos.
 */
function runDocs_Alerts() {
  Logger.log("Iniciando verificação de alertas para Documentos...");
  const alertGenerator = new AlertGenerator(SPREADSHEET_ID, CONFIG_DOCS);
  alertGenerator.processAndSendEmail(
    "Alerta de Vencimento de Documentos",
    "Os seguintes documentos estão próximos do vencimento:"
  );
}


/**************************************************************************************************
 * MOTOR DE GERAÇÃO DE ALERTAS (CLASSE REUTILIZÁVEL)
 **************************************************************************************************/
class AlertGenerator {
  constructor(spreadsheetId, config) {
    this.spreadsheetId = spreadsheetId;
    this.config = config;
    this.today = new Date();
    this.today.setHours(0, 0, 0, 0);
  }

  /**
   * Processa a planilha e envia o e-mail se encontrar alertas.
   * @param {string} emailSubject - O assunto do e-mail.
   * @param {string} emailIntro - A frase introdutória no corpo do e-mail.
   */
  processAndSendEmail(emailSubject, emailIntro) {
    try {
      const sheet = SpreadsheetApp.openById(this.spreadsheetId).getSheetByName(this.config.SHEET_NAME);
      if (!sheet) {
        throw new Error(`Aba "${this.config.SHEET_NAME}" não encontrada.`);
      }

      const data = sheet.getDataRange().getValues();
      const startIndex = this.config.START_ROW - 1;

      if (data.length <= startIndex) {
        Logger.log("Nenhuma linha de dados para processar.");
        return;
      }

      const htmlMessages = this.generateAlerts(data, startIndex);

      Logger.log(`Processamento concluído. ${htmlMessages.length} alertas encontrados para "${this.config.SHEET_NAME}".`);

      if (htmlMessages.length > 0) {
        const fullHtml = this.createEmailBody(emailIntro, htmlMessages.join(''));
        MailApp.sendEmail({
          to: this.config.DESTINATION_EMAIL,
          bcc: this.config.BCC_EMAIL,
          subject: emailSubject,
          htmlBody: fullHtml
        });
        Logger.log(`E-mail de alerta para "${this.config.SHEET_NAME}" enviado com sucesso.`);
      } else {
        Logger.log(`Nenhum alerta para enviar hoje para "${this.config.SHEET_NAME}".`);
      }
    } catch (e) {
      Logger.log(`ERRO ao processar "${this.config.SHEET_NAME}": ${e.message}`);
    }
  }

  /**
   * Itera sobre os dados da planilha e gera as mensagens de alerta.
   * @param {Array<Array<string>>} data - Os dados da planilha.
   * @param {number} startIndex - O índice da linha para começar a verificação.
   * @returns {Array<string>} Um array de linhas HTML para a tabela de alertas.
   */
  generateAlerts(data, startIndex) {
    const htmlMessages = [];
    const cols = this.config.COLUMNS;

    for (let i = startIndex; i < data.length; i++) {
      const row = data[i];
      if (row[cols.id] === '') continue; // Pula linhas vazias

      cols.dates_to_check.forEach(dateInfo => {
        const dateValue = this.parseDate(row[dateInfo.index]);
        if (!dateValue) return;

        // Regra de validação para eficácia
        if (dateInfo.name === "Eficácia" && cols.effectiveness_completed_check) {
          const marker = row[cols.effectiveness_completed_check.column_index];
          if (String(marker).trim().toUpperCase() === cols.effectiveness_completed_check.completed_marker) {
            return; // Pula se estiver marcado como concluído
          }
        }

        const daysRemaining = Math.round((dateValue.getTime() - this.today.getTime()) / (1000 * 60 * 60 * 24));

        if (this.config.ALERT_DAYS.includes(daysRemaining)) {
          htmlMessages.push(this.formatHtmlRow(row, dateInfo, dateValue, daysRemaining));
        }
      });
    }
    return htmlMessages;
  }

  /**
   * Formata uma linha de alerta em HTML para a tabela do e-mail.
   */
  formatHtmlRow(row, dateInfo, dateValue, daysRemaining) {
    const cols = this.config.COLUMNS;
    const safe = (v) => v === null || v === undefined ? "" : String(v);

    const responsible = cols.responsible !== null ? safe(row[cols.responsible]) : "N/A";
    const formattedDate = Utilities.formatDate(dateValue, "America/Sao_Paulo", "dd/MM/yyyy");

    return `
      <tr>
        <td>${safe(row[cols.id])}</td>
        <td>${safe(row[cols.description])}</td>
        <td>${safe(row[cols.action])}</td>
        <td>${formattedDate}</td>
        <td>${safe(dateInfo.name)}</td>
        <td>${daysRemaining}</td>
        <td>${responsible}</td>
      </tr>`;
  }

  /**
   * Cria o corpo HTML completo do e-mail.
   */
  createEmailBody(intro, htmlRows) {
    return `
      <!DOCTYPE html><html><head><style>
        body { font-family: Arial, sans-serif; margin: 20px; background-color: #f4f4f9; color: #333; }
        table { width: 100%; border-collapse: collapse; margin-top: 20px; box-shadow: 0 2px 5px rgba(0,0,0,0.1); }
        th, td { padding: 12px; text-align: left; border-bottom: 1px solid #ddd; }
        th { background-color: #005860; color: white; text-align: center; }
        tr:nth-child(even) { background-color: #f2f2f2; }
        h1 { color: #005860; }
      </style></head><body>
        <h1>📌 Alerta de Vencimentos</h1>
        <p>${intro}</p>
        <table>
          <thead><tr>
            <th>ID/NC</th><th>Descrição/Desvio</th><th>Ação/Documento</th><th>Data Venc.</th>
            <th>Status</th><th>Dias Restantes</th><th>Responsável</th>
          </tr></thead>
          <tbody>${htmlRows}</tbody>
        </table>
      </body></html>`;
  }

  /**
   * Converte um valor da célula (data, número ou texto) em um objeto Date.
   */
  parseDate(value) {
    if (!value) return null;
    try {
      const date = new Date(value);
      if (date && !isNaN(date.getTime())) {
        date.setHours(0, 0, 0, 0);
        return date;
      }
    } catch (e) {}
    return null;
  }
}
