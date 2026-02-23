import { google } from 'googleapis';
import fs from 'fs/promises';
import path from 'path';
import http from 'http';
import { URL } from 'url';

const SCOPES = [
  'https://www.googleapis.com/auth/drive.file',
  'https://www.googleapis.com/auth/gmail.send',
  'https://www.googleapis.com/auth/documents',
  'https://www.googleapis.com/auth/spreadsheets',
  'https://www.googleapis.com/auth/presentations'
];

const TOKEN_PATH = path.join(process.cwd(), 'token.json');
const CREDENTIALS_PATH = path.join(process.cwd(), 'credentials.json');

async function loadCredentials() {
  try {
    const content = await fs.readFile(CREDENTIALS_PATH, 'utf-8');
    return JSON.parse(content);
  } catch (error) {
    console.error(`Error loading credentials.json from ${CREDENTIALS_PATH}`);
    console.error('Please download your OAuth 2.0 Client ID (Desktop App) from Google Cloud Console and save it as credentials.json');
    process.exit(1);
  }
}

async function authorize() {
  const credentials = await loadCredentials();
  const { client_secret, client_id, redirect_uris } = credentials.installed || credentials.web;
  
  // Important: We must force a specific redirect URI that matches what we configured in Google Cloud
  const redirectUri = 'http://localhost:3000';
  const oAuth2Client = new google.auth.OAuth2(client_id, client_secret, redirectUri);

  try {
    const token = await fs.readFile(TOKEN_PATH, 'utf-8');
    oAuth2Client.setCredentials(JSON.parse(token));
    console.log('Token already exists at', TOKEN_PATH);
    return oAuth2Client;
  } catch (err) {
    return getNewToken(client_id, client_secret);
  }
}

async function getNewToken(client_id: string, client_secret: string): Promise<any> {
  // Create a local server to receive the authorization code
  const server = http.createServer();
  
  return new Promise((resolve, reject) => {
    server.listen(0, '127.0.0.1', () => {
      const address = server.address() as any;
      const redirectUri = `http://127.0.0.1:${address.port}`;
      
      const oAuth2Client = new google.auth.OAuth2(client_id, client_secret, redirectUri);
      
      const authUrl = oAuth2Client.generateAuthUrl({
        access_type: 'offline',
        scope: SCOPES,
      });
      
      console.log('Authorize this app by visiting this url:\\n', authUrl);
      console.log(`\\nWaiting for authorization redirect on ${redirectUri}...`);

      server.on('request', async (req, res) => {
        try {
          if (!req.url) return;
          
          const url = new URL(req.url, redirectUri);
          const code = url.searchParams.get('code');
          
          if (code) {
            res.writeHead(200, { 'Content-Type': 'text/html' });
            res.end('<h1>Authorization successful!</h1><p>You can safely close this tab and return to your terminal.</p>');
            
            try {
              const { tokens } = await oAuth2Client.getToken(code);
              oAuth2Client.setCredentials(tokens);
              await fs.writeFile(TOKEN_PATH, JSON.stringify(tokens));
              console.log('Token stored to', TOKEN_PATH);
              server.close();
              resolve(oAuth2Client);
            } catch (err) {
              console.error('Error retrieving access token', err);
              server.close();
              reject(err);
            }
          } else if (url.searchParams.get('error')) {
            res.writeHead(400, { 'Content-Type': 'text/html' });
            res.end(`<h1>Authorization failed</h1><p>Error: ${url.searchParams.get('error')}</p>`);
            server.close();
            reject(new Error(url.searchParams.get('error') || 'Unknown error'));
          } else if (req.url !== '/favicon.ico') {
            res.writeHead(404, { 'Content-Type': 'text/plain' });
            res.end('Not found');
          }
        } catch (e) {
          console.error('Error handling request:', e);
          res.writeHead(500, { 'Content-Type': 'text/plain' });
          res.end('Internal Server Error');
        }
      });
    });
    
    server.on('error', (e: any) => {
      console.error('Server error:', e);
      reject(e);
    });
  });
}

authorize().then(() => {
  console.log('Authorization complete! You can now run the MCP server.');
  process.exit(0);
}).catch((error) => {
  console.error('Authorization failed:', error);
  process.exit(1);
});