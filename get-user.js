const { WebClient } = require('@slack/web-api');

const token = 'XXXX-XXXXXXXXXXXXXXXXXXXXXXXXXXX';
const userEmail = 'email@email.com'; // Substitua pelo e-mail do usuário
const web = new WebClient(token);

async function getUserId() {
  try {
    const response = await web.users.lookupByEmail({ email: userEmail });
    console.log('User ID:', response.user.id); // Aqui você obtém o ID do usuário
  } catch (error) {
    console.error('Erro ao obter informações do usuário:', error);
  }
}

getUserId();
