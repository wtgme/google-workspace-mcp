import { Server } from '@modelcontextprotocol/sdk/server/index.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import { CallToolRequestSchema, ListToolsRequestSchema, Tool } from '@modelcontextprotocol/sdk/types.js';
import { google } from 'googleapis';
import fs from 'fs/promises';
import path from 'path';

const TOKEN_PATH = path.join(__dirname, '..', 'token.json');
const CREDENTIALS_PATH = path.join(__dirname, '..', 'credentials.json');

let authClient: any = null;

async function getAuthClient() {
  if (authClient) return authClient;

  try {
    const credContent = await fs.readFile(CREDENTIALS_PATH, 'utf-8');
    const credentials = JSON.parse(credContent);
    const { client_secret, client_id, redirect_uris } = credentials.installed || credentials.web;
    const redirectUri = redirect_uris && redirect_uris.length > 0 ? redirect_uris[0] : 'http://localhost';

    const oAuth2Client = new google.auth.OAuth2(client_id, client_secret, redirectUri);
    
    const tokenContent = await fs.readFile(TOKEN_PATH, 'utf-8');
    oAuth2Client.setCredentials(JSON.parse(tokenContent));
    
    authClient = oAuth2Client;
    return authClient;
  } catch (error) {
    console.error('Error loading Google API credentials. Please run authorize.ts first.');
    throw error;
  }
}

const create_file_tool: Tool = {
  name: 'create_drive_file',
  description: 'Create a new text or markdown file on Google Drive.',
  inputSchema: {
    type: 'object',
    properties: {
      name: { type: 'string', description: 'Name of the file to create (e.g. document.txt)' },
      content: { type: 'string', description: 'Content to write into the file' },
      mimeType: { type: 'string', description: 'MIME type of the file (e.g. text/plain)', default: 'text/plain' },
      folderId: { type: 'string', description: 'Optional. The Google Drive folder ID where the document should be created.' },
      folderPath: { type: 'string', description: 'Optional. A string path like "MyFolder/Subfolder" where the document should be created. (Ignored if folderId is provided)' }
    },
    required: ['name', 'content']
  }
};

const update_file_tool: Tool = {
  name: 'update_drive_file',
  description: 'Update the content of an existing file on Google Drive by its file ID.',
  inputSchema: {
    type: 'object',
    properties: {
      fileId: { type: 'string', description: 'The ID of the file to update on Google Drive' },
      content: { type: 'string', description: 'New content to write into the file' },
      mimeType: { type: 'string', description: 'MIME type of the file (e.g. text/plain)', default: 'text/plain' }
    },
    required: ['fileId', 'content']
  }
};

const send_gmail_tool: Tool = {
  name: 'send_gmail',
  description: 'Send an email via Gmail.',
  inputSchema: {
    type: 'object',
    properties: {
      to: { type: 'string', description: 'Recipient email address' },
      subject: { type: 'string', description: 'Email subject' },
      body: { type: 'string', description: 'Email body content (text/plain)' }
    },
    required: ['to', 'subject', 'body']
  }
};

const append_google_doc_tool: Tool = {
  name: 'append_google_doc',
  description: 'Append text to the end of a Google Doc.',
  inputSchema: {
    type: 'object',
    properties: {
      documentId: { type: 'string', description: 'The ID of the Google Doc' },
      text: { type: 'string', description: 'The text to append to the document' }
    },
    required: ['documentId', 'text']
  }
};

const append_google_sheet_tool: Tool = {
  name: 'append_google_sheet',
  description: 'Append rows of data to a Google Sheet.',
  inputSchema: {
    type: 'object',
    properties: {
      spreadsheetId: { type: 'string', description: 'The ID of the Google Sheet' },
      range: { type: 'string', description: 'The A1 notation of a range to search for a logical table of data (e.g., "Sheet1!A1").' },
      values: { 
        type: 'array', 
        description: 'A 2D array of values to append, e.g., [["row1col1", "row1col2"], ["row2col1", "row2col2"]]',
        items: { type: 'array', items: { type: 'string' } }
      }
    },
    required: ['spreadsheetId', 'range', 'values']
  }
};

const add_google_slide_tool: Tool = {
  name: 'add_google_slide',
  description: 'Add a new slide with text to a Google Slides presentation.',
  inputSchema: {
    type: 'object',
    properties: {
      presentationId: { type: 'string', description: 'The ID of the Google Slides presentation' },
      text: { type: 'string', description: 'The text to insert into the new slide' }
    },
    required: ['presentationId', 'text']
  }
};

const create_google_doc_tool: Tool = {
  name: 'create_google_doc',
  description: 'Create a new Google Doc, optionally with initial text content, in a specific folder or path.',
  inputSchema: {
    type: 'object',
    properties: {
      title: { type: 'string', description: 'The title of the new Google Doc' },
      content: { type: 'string', description: 'Optional initial text content to insert into the document' },
      folderId: { type: 'string', description: 'Optional. The Google Drive folder ID where the document should be created.' },
      folderPath: { type: 'string', description: 'Optional. A string path like "MyFolder/Subfolder" where the document should be created. (Ignored if folderId is provided)' }
    },
    required: ['title']
  }
};

const create_google_sheet_tool: Tool = {
  name: 'create_google_sheet',
  description: 'Create a new Google Sheet in a specific folder or path.',
  inputSchema: {
    type: 'object',
    properties: {
      title: { type: 'string', description: 'The title of the new Google Sheet' },
      folderId: { type: 'string', description: 'Optional. The Google Drive folder ID where the sheet should be created.' },
      folderPath: { type: 'string', description: 'Optional. A string path like "MyFolder/Subfolder" where the sheet should be created. (Ignored if folderId is provided)' }
    },
    required: ['title']
  }
};

const create_google_slide_tool: Tool = {
  name: 'create_google_slide',
  description: 'Create a new Google Slides presentation in a specific folder or path.',
  inputSchema: {
    type: 'object',
    properties: {
      title: { type: 'string', description: 'The title of the new Google Slides presentation' },
      folderId: { type: 'string', description: 'Optional. The Google Drive folder ID where the presentation should be created.' },
      folderPath: { type: 'string', description: 'Optional. A string path like "MyFolder/Subfolder" where the presentation should be created. (Ignored if folderId is provided)' }
    },
    required: ['title']
  }
};

const server = new Server({
  name: 'google-workspace-mcp',
  version: '1.0.0'
}, {
  capabilities: {
    tools: {}
  }
});

server.setRequestHandler(ListToolsRequestSchema, async () => {
  return {
    tools: [
      create_file_tool, 
      update_file_tool, 
      send_gmail_tool,
      append_google_doc_tool,
      append_google_sheet_tool,
      add_google_slide_tool,
      create_google_doc_tool,
      create_google_sheet_tool,
      create_google_slide_tool
    ]
  };
});

server.setRequestHandler(CallToolRequestSchema, async (request) => {
  const { name, arguments: args } = request.params;
  
  if (!args) {
    throw new Error('Missing arguments');
  }

  try {
    const auth = await getAuthClient();

    if (name === 'create_drive_file') {
      const drive = google.drive({ version: 'v3', auth });
      const { name: fileName, content, mimeType = 'text/plain', folderId, folderPath } = args as any;

      let parentId = folderId;
      if (!parentId && folderPath) {
        parentId = await getFolderIdByPath(drive, folderPath);
      }

      const res = await drive.files.create({
        requestBody: {
          name: fileName,
          mimeType: mimeType,
          parents: parentId ? [parentId] : undefined
        },
        media: {
          mimeType: mimeType,
          body: content
        }
      });

      return {
        content: [{ type: 'text', text: `Successfully created file ${fileName} on Drive.\nFile ID: ${res.data.id}\nFolder ID: ${parentId || 'root'}` }]
      };
    }

    if (name === 'update_drive_file') {
      const drive = google.drive({ version: 'v3', auth });
      const { fileId, content, mimeType = 'text/plain' } = args as any;

      const res = await drive.files.update({
        fileId: fileId,
        media: {
          mimeType: mimeType,
          body: content
        }
      });

      return {
        content: [{ type: 'text', text: `Successfully updated file on Drive.\nFile ID: ${res.data.id}` }]
      };
    }

    if (name === 'send_gmail') {
      const gmail = google.gmail({ version: 'v1', auth });
      const { to, subject, body } = args as any;

      const messageParts = [
        `To: ${to}`,
        'Content-Type: text/plain; charset=utf-8',
        'MIME-Version: 1.0',
        `Subject: ${subject}`,
        '',
        body
      ];
      const message = messageParts.join('\r\n');
      
      const encodedMessage = Buffer.from(message)
        .toString('base64')
        .replace(/\+/g, '-')
        .replace(/\//g, '_')
        .replace(/=+$/, '');

      const res = await gmail.users.messages.send({
        userId: 'me',
        requestBody: {
          raw: encodedMessage
        }
      });

      return {
        content: [{ type: 'text', text: `Successfully sent email to ${to}.\nMessage ID: ${res.data.id}` }]
      };
    }

    if (name === 'append_google_doc') {
      const docs = google.docs({ version: 'v1', auth });
      const { documentId, text } = args as any;

      const doc = await docs.documents.get({ documentId });
      const contentList = doc.data.body?.content;
      if (!contentList || contentList.length === 0) {
         throw new Error('Could not find document body content');
      }
      
      const lastElement = contentList[contentList.length - 1];
      const endIndex = (lastElement.endIndex || 2) - 1;

      await docs.documents.batchUpdate({
        documentId,
        requestBody: {
          requests: [
            {
              insertText: {
                location: { index: endIndex },
                text: '\n' + text
              }
            }
          ]
        }
      });

      return {
        content: [{ type: 'text', text: `Successfully appended text to Google Doc.\nDocument ID: ${documentId}` }]
      };
    }

    if (name === 'append_google_sheet') {
      const sheets = google.sheets({ version: 'v4', auth });
      const { spreadsheetId, range, values } = args as any;

      const res = await sheets.spreadsheets.values.append({
        spreadsheetId,
        range,
        valueInputOption: 'USER_ENTERED',
        requestBody: {
          values: values
        }
      });

      return {
        content: [{ type: 'text', text: `Successfully appended rows to Google Sheet.\nUpdates: ${JSON.stringify(res.data.updates)}` }]
      };
    }

    if (name === 'add_google_slide') {
      const slides = google.slides({ version: 'v1', auth });
      const { presentationId, text } = args as any;

      const pageId = 'page_' + Date.now();
      const shapeId = 'shape_' + Date.now();

      await slides.presentations.batchUpdate({
        presentationId,
        requestBody: {
          requests: [
            {
              createSlide: {
                objectId: pageId,
                slideLayoutReference: { predefinedLayout: 'BLANK' }
              }
            },
            {
              createShape: {
                objectId: shapeId,
                shapeType: 'TEXT_BOX',
                elementProperties: {
                  pageObjectId: pageId,
                  size: { width: { magnitude: 600, unit: 'PT' }, height: { magnitude: 300, unit: 'PT' } },
                  transform: { scaleX: 1, scaleY: 1, translateX: 50, translateY: 50, unit: 'PT' }
                }
              }
            },
            {
              insertText: {
                objectId: shapeId,
                text: text
              }
            }
          ]
        }
      });

      return {
        content: [{ type: 'text', text: `Successfully added slide to Google Presentation.\nPresentation ID: ${presentationId}` }]
      };
    }

    if (name === 'create_google_doc') {
      const drive = google.drive({ version: 'v3', auth });
      const docs = google.docs({ version: 'v1', auth });
      const { title, content, folderId, folderPath } = args as any;

      let parentId = folderId;
      if (!parentId && folderPath) {
        parentId = await getFolderIdByPath(drive, folderPath);
      }

      const res = await drive.files.create({
        requestBody: {
          name: title,
          mimeType: 'application/vnd.google-apps.document',
          parents: parentId ? [parentId] : undefined
        }
      });
      
      const documentId = res.data.id;

      if (content && documentId) {
        await docs.documents.batchUpdate({
          documentId,
          requestBody: {
            requests: [
              {
                insertText: {
                  location: { index: 1 },
                  text: content
                }
              }
            ]
          }
        });
      }

      return {
        content: [{ type: 'text', text: `Successfully created Google Doc.\nTitle: ${res.data.name}\nDocument ID: ${documentId}\nFolder ID: ${parentId || 'root'}` }]
      };
    }

    if (name === 'create_google_sheet') {
      const drive = google.drive({ version: 'v3', auth });
      const { title, folderId, folderPath } = args as any;

      let parentId = folderId;
      if (!parentId && folderPath) {
        parentId = await getFolderIdByPath(drive, folderPath);
      }

      const res = await drive.files.create({
        requestBody: {
          name: title,
          mimeType: 'application/vnd.google-apps.spreadsheet',
          parents: parentId ? [parentId] : undefined
        }
      });

      return {
        content: [{ type: 'text', text: `Successfully created Google Sheet.\nTitle: ${res.data.name}\nSpreadsheet ID: ${res.data.id}\nFolder ID: ${parentId || 'root'}` }]
      };
    }

    if (name === 'create_google_slide') {
      const drive = google.drive({ version: 'v3', auth });
      const { title, folderId, folderPath } = args as any;

      let parentId = folderId;
      if (!parentId && folderPath) {
        parentId = await getFolderIdByPath(drive, folderPath);
      }

      const res = await drive.files.create({
        requestBody: {
          name: title,
          mimeType: 'application/vnd.google-apps.presentation',
          parents: parentId ? [parentId] : undefined
        }
      });

      return {
        content: [{ type: 'text', text: `Successfully created Google Slides presentation.\nTitle: ${res.data.name}\nPresentation ID: ${res.data.id}\nFolder ID: ${parentId || 'root'}` }]
      };
    }

    throw new Error(`Unknown tool: ${name}`);
  } catch (error: any) {
    return {
      content: [{ type: 'text', text: `Error: ${error.message}` }],
      isError: true
    };
  }
});

async function getFolderIdByPath(drive: any, folderPath: string): Promise<string> {
  if (!folderPath || folderPath === '/' || folderPath === '') return 'root';
  const parts = folderPath.split('/').filter(p => p.length > 0);
  let currentParent = 'root';
  for (const part of parts) {
    const query = `name = '${part.replace(/'/g, "\\'")}' and '${currentParent}' in parents and mimeType = 'application/vnd.google-apps.folder' and trashed = false`;
    const res = await drive.files.list({ q: query, fields: 'files(id, name)' });
    if (res.data.files && res.data.files.length > 0) {
      currentParent = res.data.files[0].id;
    } else {
      throw new Error(`Folder not found in path: ${part}`);
    }
  }
  return currentParent;
}

async function run() {
  const transport = new StdioServerTransport();
  await server.connect(transport);
  console.error('Google Workspace MCP Server running on stdio');
}

run().catch(error => {
  console.error('Fatal error in server:', error);
  process.exit(1);
});