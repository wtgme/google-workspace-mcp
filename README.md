# Google Workspace MCP Server for Opencode

A Model Context Protocol (MCP) server built specifically for the **Opencode CLI** that connects to your Google Account to manage files and send emails. This server provides tools to create and update Google Drive files, Google Docs, Sheets, Slides, and send emails via Gmail.

## Tools Provided

- **`create_drive_file`**: Create a new text or markdown file on Google Drive.
- **`update_drive_file`**: Update an existing file on Google Drive.
- **`send_gmail`**: Send an email via Gmail.
- **`append_google_doc`**: Append text to the end of a Google Doc.
- **`append_google_sheet`**: Append rows of data to a Google Sheet.
- **`add_google_slide`**: Add a new slide with text to a Google Slides presentation.

## Setup Instructions

### 1. Generate Google Cloud Credentials

To use this server, you need to create an OAuth 2.0 Client ID in the Google Cloud Console.

1. Go to the [Google Cloud Console](https://console.cloud.google.com/).
2. Create a new project (or select an existing one).
3. Navigate to **APIs & Services > Library** and enable the following APIs:
   - **Google Drive API**
   - **Gmail API**
   - **Google Docs API**
   - **Google Sheets API**
   - **Google Slides API**
4. Go to **APIs & Services > OAuth consent screen**:
   - Select **External** (or Internal if using a Google Workspace account).
   - Fill in the required app details.
   - **Important:** Add your own email address as a **Test user** so you can authenticate yourself.
5. Go to **APIs & Services > Credentials**:
   - Click **Create Credentials** and select **OAuth client ID**.
   - Choose **Desktop app** as the application type.
   - Click **Create**, then click **Download JSON**.
6. Rename the downloaded file to `credentials.json` and place it in the root directory of this project.

### 2. Install Dependencies & Build

```bash
npm install
npm run build
```

### 3. Authorize the Server

Before the server can interact with your Google Account, you need to authorize it and generate a token.

Run the following command:

```bash
npm run auth
```

1. This will output a URL in your terminal.
2. Open the URL in your browser and log in with your Google Account.
3. Review and grant the requested permissions.
4. The script will automatically catch the redirect and save a `token.json` file in the root directory.

### 4. Add to Opencode

To make this MCP server globally available across all your Opencode workspaces, you need to add it to your global `opencode.json` configuration.

Open your Opencode configuration file (usually located at `~/.config/opencode/opencode.json`) and add the `mcp` block with the absolute path to your built `index.js` file:

```json
{
  ...
  "mcp": {
    "google-workspace": {
      "type": "local",
      "command": [
        "node",
        "/absolute/path/to/google-workspace-mcp/build/index.js"
      ],
      "enabled": true
    }
  }
}
```

*Note: Replace `/absolute/path/to/google-workspace-mcp/build/index.js` with the actual path on your machine.*

Once added, verify the server is active by running:
```bash
opencode mcp list
```

You should see `google-workspace connected`!

## Security Note
Do **not** commit your `credentials.json` or `token.json` files to version control. They provide access to your Google Account.
