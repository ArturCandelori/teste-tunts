// Implementação dos cálculos a partir da linha 85
// Processo de autenticação retirado de https://developers.google.com/sheets/api/quickstart/nodejs

const fs = require('fs');
const readline = require('readline');
const { google } = require('googleapis');

// If modifying these scopes, delete token.json.
const SCOPES = ['https://www.googleapis.com/auth/spreadsheets'];
// The file token.json stores the user's access and refresh tokens, and is
// created automatically when the authorization flow completes for the first
// time.
const TOKEN_PATH = 'token.json';

// Load client secrets from a local file.
fs.readFile('credentials.json', (err, content) => {
  if (err) return console.log('Error loading client secret file:', err);
  // Authorize a client with credentials, then call the Google Sheets API.
  authorize(JSON.parse(content), calcularMedias);
});

/**
 * Create an OAuth2 client with the given credentials, and then execute the
 * given callback function.
 * @param {Object} credentials The authorization client credentials.
 * @param {function} callback The callback to call with the authorized client.
 */
function authorize(credentials, callback) {
  const { client_secret, client_id, redirect_uris } = credentials.installed;
  const oAuth2Client = new google.auth.OAuth2(
    client_id,
    client_secret,
    redirect_uris[0]
  );

  // Check if we have previously stored a token.
  fs.readFile(TOKEN_PATH, (err, token) => {
    if (err) return getNewToken(oAuth2Client, callback);
    oAuth2Client.setCredentials(JSON.parse(token));
    callback(oAuth2Client);
  });
}

/**
 * Get and store new token after prompting for user authorization, and then
 * execute the given callback with the authorized OAuth2 client.
 * @param {google.auth.OAuth2} oAuth2Client The OAuth2 client to get token for.
 * @param {getEventsCallback} callback The callback for the authorized client.
 */
function getNewToken(oAuth2Client, callback) {
  const authUrl = oAuth2Client.generateAuthUrl({
    access_type: 'offline',
    scope: SCOPES,
  });
  console.log('Authorize this app by visiting this url:', authUrl);
  const rl = readline.createInterface({
    input: process.stdin,
    output: process.stdout,
  });
  rl.question('Enter the code from that page here: ', code => {
    rl.close();
    oAuth2Client.getToken(code, (err, token) => {
      if (err)
        return console.error(
          'Error while trying to retrieve access token',
          err
        );
      oAuth2Client.setCredentials(token);
      // Store the token to disk for later program executions
      fs.writeFile(TOKEN_PATH, JSON.stringify(token), err => {
        if (err) return console.error(err);
        console.log('Token stored to', TOKEN_PATH);
      });
      callback(oAuth2Client);
    });
  });
}

/**
 * Calcula as médias e a situação de cada aluno na seguinte tabela:
 * @see https://docs.google.com/spreadsheets/d/1Pcf20mTElugmzbYBhAE3garSIcg2La0cJTC30RYLEnM/edit#gid=0
 * @param {google.auth.OAuth2} auth The authenticated Google OAuth client.
 */

async function calcularMedias(auth) {
  try {
    const sheets = google.sheets({ version: 'v4', auth });

    // acessa os dados da tabela fornecida
    let tabela = await sheets.spreadsheets.values.get({
      spreadsheetId: '1Pcf20mTElugmzbYBhAE3garSIcg2La0cJTC30RYLEnM',
      range: 'engenharia_de_software!A4:F27',
    });

    // pega somente os dados dos alunos, sendo cada linha da tabela um array
    let alunos = tabela.data.values;
    //console.log(alunos)

    // lógica para calcular o resultado
    // resultado fica em um array com a situação e nota para aprovação final de cada aluno
    let resultados = alunos.map(aluno => {
      const f = Number(aluno[2]); //faltas
      const p1 = Number(aluno[3]);
      const p2 = Number(aluno[4]);
      const p3 = Number(aluno[5]);

      const m = Math.round((p1 + p2 + p3) / 3);
      let situacao = '';
      let naf = 0; // nota para aprovação final

      if (f > 15) {
        situacao = 'Reprovado por Falta';
      } else if (m > 70) {
        situacao = 'Aprovado';
      } else if (m < 50) {
        situacao = 'Reprovado por Nota';
      } else {
        situacao = 'Exame Final';
      }

      // caso aluno fique para final, essa será a nota mínima
      if (situacao === 'Exame Final') {
        naf = 100 - m;
      }

      return [situacao, naf];
    });

    // inserindo os resultados na tabela
    let res = await sheets.spreadsheets.values.update({
      spreadsheetId: '1Pcf20mTElugmzbYBhAE3garSIcg2La0cJTC30RYLEnM',
      range: 'engenharia_de_software!G4',
      valueInputOption: 'USER_ENTERED',
      resource: { values: resultados },
    });

    // verificando a resposta da api
    console.log(res);
  } catch (error) {
    console.log('Erro: ', error);
  }
}
