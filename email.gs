// Traz a data atual e a formatação do google
const today = new Date();
const formatToday = Utilities.formatDate(today, Session.getScriptTimeZone(), "dd/MM/yyyy");

// Retorna dados da tabela e conf e-mail
const SHEET_NAME = "" // Aqui insira o nome da sua Planilha (nome da folha);
const SHEET_NAME_EMAIL = "E-MAIL";  // Crie uma nova folha para diusparar de forma dinamica
const SUBJECT = `Alerta! Seu plano de e-mails venceu - ${formatToday}` // Titulo do e-mail com a data formatada;

// Consulta dados JSON e filtra contratos vencidos
function sendAlertByEmail() {
  const planilha = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME) // Traz os dados da planilha de SHEET_NAME;
  const sheetEmail = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME_EMAIL) // Traz os dados da planilha de SHEET_NAME_EMAIL;
  const dadosEmail = sheetEmail.getDataRange().getValues() // retorna os dados da planilha em array;
  const dados = planilha.getDataRange().getValues() // retorna os dados da planilha em array;
  const indiceColunaData = 10;

  // Data de 30 dias atrás
  const trintaDiasAtras = new Date();
  trintaDiasAtras.setDate(hoje.getDate() - 30);

  // Filtra contratos vencidos nos últimos 30 dias
  const dadosFiltrados = dados.filter((linha, index) => {
    if (index === 0) return false; // Ignorar o cabeçalho
    const dataCelula = new Date(linha[indiceColunaData]); // Coluna K
    return dataCelula < today && dataCelula >= trintaDiasAtras; // Vencidos nos últimos 30 dias
  });

  enviaEmail(dadosFiltrados, indiceColunaData, dadosEmail);
}

function enviaEmail(dadosFiltrados, indiceColunaData, dadosEmail) {
  // Construir mensagens para os contratos vencidos
  let mensagens = dadosFiltrados.map(linha => {
    const nomeColaborador = linha[4];
    const nomeContrato = linha[2];
    const egNumeroVencido = linha[0]; // EG do contrato
    const dataVencimento = new Date(linha[indiceColunaData]);
    const dataFormatada = Utilities.formatDate(dataVencimento, Session.getScriptTimeZone(), "dd/MM/yyyy");

    return `Olá ${nomeColaborador}, o contrato ${nomeContrato} referente ao ${egNumeroVencido} venceu na data ${dataFormatada}.`;
  }).join("\n\n");

  // Enviar e-mail se houver mensagens
  if (mensagens) {
    let emailDestino = "contratos@engeplus.eng.br"; // Default para quando não houver correspondência
    dadosEmail.forEach((email) => {
      dadosFiltrados.forEach(linha => {
        const nomeColaborador = linha[4];
        if (nomeColaborador === email[1]) {
          emailDestino = email[0];
        }
      });
    });

    GmailApp.sendEmail(emailDestino, SUBJECT, mensagens);
    Logger.log("E-mail enviado com sucesso!");
  } else {
    Logger.log("Nenhum contrato vencido nos últimos 30 dias para enviar alerta.");
  }
}
