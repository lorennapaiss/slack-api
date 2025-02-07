/**
 * Script para envio automatizado de mensagens no Slack com base em uma planilha Excel.
 * Agrupa indicadores por responsável e envia uma mensagem única via API do Slack.
 */

const { WebClient } = require('@slack/web-api');
const xlsx = require('xlsx');

// Token de autenticação do Slack (substitua pelo seu token válido)
const token = 'XXXX-XXXXXXXXXXXXXXXXXXXXXXXXXXX';
const web = new WebClient(token);

/**
 * Obtém o ID do usuário no Slack a partir do email cadastrado.
 * @param {string} email - Email do usuário.
 * @returns {Promise<string|null>} - ID do usuário no Slack ou null em caso de erro.
 */
async function getUserIdByEmail(email) {
  try {
    const response = await web.users.lookupByEmail({ email });
    return response.user.id;
  } catch (error) {
    console.error(`Erro ao obter ID para o email ${email}:`, error);
    return null;
  }
}

/**
 * Abre uma DM com um usuário no Slack.
 * @param {string} userId - ID do usuário no Slack.
 * @returns {Promise<string|null>} - ID do canal DM ou null em caso de erro.
 */
async function openDM(userId) {
  try {
    const response = await web.conversations.open({ users: userId });
    return response.channel.id;
  } catch (error) {
    console.error(`Erro ao abrir DM para ${userId}:`, error);
    return null;
  }
}

/**
 * Envia mensagens personalizadas para usuários no Slack.
 * @param {Array} userData - Lista de objetos contendo userId, nome e indicadores.
 */
async function sendMessageToUsers(userData) {
  for (const { userId, name, indicadores } of userData) {
    const dmChannel = await openDM(userId);
    if (!dmChannel) continue;

    try {
      // Formata a lista de indicadores como uma lista Markdown
      const indicadoresList = indicadores.map(indicador => `- ${indicador}`).join('\n');

      const message = `*_Olá ${name}! Tudo bem?_* 👋  \n\n` +
        `Estamos nos aproximando do *encerramento do ciclo de metas 2024*, e notamos que você ainda *não lançou os valores dos seguintes indicadores* que você é responsável:  \n\n` +
        `${indicadoresList}  \n\n` +
        `Por favor, acesse a plataforma *Qulture Rocks* e realize o lançamento dos indicadores até o dia *15/02*.  \n\n` +
        `Caso tenha alguma dúvida, entre em contato:  \n` +
        `📩 *E-mail:* metas@xyz.com  \n` +
        `💬 *Slack:* @Lore e @Thiago Brandão  \n\n` +
        `Estamos à disposição! 😊 Abraços.  \n\n` +
        `_*Obs:* Este canal não recebe mensagens. Por favor, entre em contato via e-mail ou Slack como mencionado acima._`;

      const result = await web.chat.postMessage({
        channel: dmChannel,
        text: message,
        mrkdwn: true // Habilita Markdown para formatação no Slack
      });

      console.log(`Mensagem enviada para ${name} (${userId}):`, result.ts);
    } catch (error) {
      console.error(`Erro ao enviar mensagem para ${name} (${userId}):`, error);
    }
  }
}

/**
 * Processa a planilha Excel, agrupa indicadores por responsável e envia mensagens no Slack.
 * @param {string} filePath - Caminho do arquivo Excel.
 */
async function processEmailsFromExcel(filePath) {
  const workbook = xlsx.readFile(filePath);
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const data = xlsx.utils.sheet_to_json(sheet);

  // Agrupar indicadores por responsável
  const responsaveis = {};
  for (const row of data) {
    const email = row.Email;
    const name = row.Nome;
    const indicador = row.Indicador; // Supondo que a coluna do indicador se chame "Indicador"

    if (!responsaveis[email]) {
      responsaveis[email] = { name, email, indicadores: [] };
    }
    responsaveis[email].indicadores.push(indicador);
  }

  // Buscar os IDs do Slack para cada responsável
  const userData = [];
  for (const email in responsaveis) {
    const userId = await getUserIdByEmail(email);
    if (userId) {
      userData.push({ userId, ...responsaveis[email] });
    }
  }

  // Enviar mensagens personalizadas
  await sendMessageToUsers(userData);
}

// Caminho do arquivo Excel
const filePath = './emails.xlsx';
processEmailsFromExcel(filePath);
