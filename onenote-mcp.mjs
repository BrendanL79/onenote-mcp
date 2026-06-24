#!/usr/bin/env node

import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { Client } from '@microsoft/microsoft-graph-client';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import { z } from 'zod';
import dotenv from 'dotenv';
import fetch from 'node-fetch';

import {
  saveToken,
  loadToken,
  refreshAccessToken,
  isTokenExpired,
  setTokenEndpointConfig,
} from './token-store.mjs';

dotenv.config();

const server = new McpServer(
  { name: "onenote", version: "1.0.0", description: "OneNote MCP Server" },
  { capabilities: { tools: { listChanged: true } } }
);

const clientId = '14d82eec-204b-4c2f-b7e8-296a70dab67e';
const TENANT = process.env.GRAPH_TENANT || 'common';
const SCOPES = 'Notes.ReadWrite Notes.Create User.Read offline_access';

setTokenEndpointConfig({ clientId, tenant: TENANT, scopes: SCOPES });

let storedToken = loadToken();
let accessToken = storedToken ? storedToken.token : null;
if (!accessToken && process.env.GRAPH_ACCESS_TOKEN) {
  accessToken = process.env.GRAPH_ACCESS_TOKEN;
}

let graphClient = null;
let pendingDeviceCode = null;

// Client ID — Microsoft Graph Explorer (public client, pre-consented for all Graph scopes)
// 'common' accepts BOTH personal Microsoft accounts (MSA) and work/school (Azure AD) accounts.
// Override with GRAPH_TENANT for a locked single-tenant flow:
//   'consumers'     -> personal MSA only
//   'organizations' -> work/school only
//   '<tenant-id>'   -> one specific org
// Non-".All" delegated scopes work for both personal and work/school accounts.
// (Personal MSA cannot be granted the ".All" variants at all; ".All" only adds
// access to notebooks the signed-in user does not own.)

function buildGraphClient(token) {
  return Client.initWithMiddleware({
    authProvider: { getAccessToken: async () => token }
  });
}

async function ensureGraphClient() {
  if (isTokenExpired() && accessToken) {
    try {
      const refreshed = await refreshAccessToken();
      if (refreshed) {
        accessToken = refreshed.token;
        graphClient = buildGraphClient(accessToken);
      }
    } catch (err) {
      console.error('Token refresh failed:', err.message);
    }
  }

  if (graphClient) return graphClient;

  if (!accessToken) {
    throw new Error("Not authenticated. Call the authenticate tool first.");
  }
  graphClient = buildGraphClient(accessToken);
  return graphClient;
}

// Start device code flow; returns { user_code, verification_uri, message, device_code, interval }
async function startDeviceCodeFlow() {
  const res = await fetch(
    `https://login.microsoftonline.com/${TENANT}/oauth2/v2.0/devicecode`,
    {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: new URLSearchParams({ client_id: clientId, scope: SCOPES }).toString(),
    }
  );
  const data = await res.json();
  if (data.error) throw new Error(`${data.error}: ${data.error_description}`);
  return data;
}

// Poll for token after user completes device code sign-in
async function pollForToken(deviceCode, intervalSecs) {
  const body = new URLSearchParams({
    grant_type: 'urn:ietf:params:oauth:grant-type:device_code',
    client_id: clientId,
    device_code: deviceCode,
  }).toString();

  while (true) {
    await new Promise(r => setTimeout(r, intervalSecs * 1000));
    const res = await fetch(
      `https://login.microsoftonline.com/${TENANT}/oauth2/v2.0/token`,
      { method: 'POST', headers: { 'Content-Type': 'application/x-www-form-urlencoded' }, body }
    );
    const data = await res.json();
    if (data.access_token) return data.access_token;
    if (data.error === 'authorization_pending') continue;
    if (data.error === 'slow_down') { intervalSecs += 5; continue; }
    throw new Error(`${data.error}: ${data.error_description}`);
  }
}

// Tool for starting authentication flow
server.tool(
  "authenticate",
  "Start the authentication flow with Microsoft Graph",
  async () => {
    try {
      // If already authenticated, confirm
      if (accessToken) {
        return { content: [{ type: "text", text: "Already authenticated with an access token." }] };
      }

      // If a device code flow is in progress, poll once to check if user completed sign-in
      if (pendingDeviceCode) {
        const { device_code, interval } = pendingDeviceCode;
        const res = await fetch(
          `https://login.microsoftonline.com/${TENANT}/oauth2/v2.0/token`,
          {
            method: 'POST',
            headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
            body: new URLSearchParams({
              grant_type: 'urn:ietf:params:oauth:grant-type:device_code',
              client_id: clientId,
              device_code: device_code,
            }).toString(),
          }
        );
        const data = await res.json();

        if (data.access_token) {
          accessToken = data.access_token;
          saveToken({
            token: data.access_token,
            refresh_token: data.refresh_token,
            expires_in: data.expires_in,
          });
          graphClient = buildGraphClient(accessToken);
          pendingDeviceCode = null;
          return { content: [{ type: "text", text: "Authentication complete. You are now signed in." }] };
        }

        if (data.error === 'authorization_pending' || data.error === 'slow_down') {
          return { content: [{ type: "text", text: "Authentication in progress. Complete the sign-in in your browser, then call authenticate again to confirm." }] };
        }

        // Flow expired or errored — reset and start fresh
        pendingDeviceCode = null;
      }

      // Start a new device code flow
      const flowData = await startDeviceCodeFlow();
      pendingDeviceCode = { device_code: flowData.device_code, interval: flowData.interval };

      return {
        content: [{
          type: "text",
          text: flowData.message + "\n\nAfter completing sign-in, call the authenticate tool again to confirm."
        }]
      };
    } catch (error) {
      console.error("Error in authentication:", error);
      throw new Error(`Authentication failed: ${error.message}`);
    }
  }
);

// Tool for saving an access token provided by the user
server.tool(
  "saveAccessToken",
  "Save a Microsoft Graph access token for later use",
  { token: z.string().describe("The access token to save") },
  async ({ token }) => {
    try {
      if (!token || token.length === 0) {
        throw new Error("Token is required");
      }
      accessToken = token;
      saveToken({ token });
      graphClient = buildGraphClient(accessToken);
      return {
        content: [
          { type: "text", text: "Access token saved successfully" }
        ]
      };
    } catch (error) {
      console.error("Error saving access token:", error);
      throw new Error(`Failed to save access token: ${error.message}`);
    }
  }
);

// Tool for listing all notebooks
server.tool(
  "listNotebooks",
  "List all OneNote notebooks",
  {},
  async () => {
    try {
      await ensureGraphClient();
      const response = await graphClient.api("/me/onenote/notebooks").get();
      return {
        content: [
          { type: "text", text: JSON.stringify(response.value) }
        ]
      };
    } catch (error) {
      console.error("Error listing notebooks:", error);
      throw new Error(`Failed to list notebooks: ${error.message}`);
    }
  }
);

// Tool for getting notebook details
server.tool(
  "getNotebook",
  "Get details of a specific notebook",
  { notebookId: z.string().describe("The ID of the notebook to retrieve") },
  async ({ notebookId }) => {
    try {
      await ensureGraphClient();
      const response = await graphClient.api(`/me/onenote/notebooks/${notebookId}`).get();
      return {
        content: [
          { type: "text", text: JSON.stringify(response) }
        ]
      };
    } catch (error) {
      console.error("Error getting notebook:", error);
      throw new Error(`Failed to get notebook: ${error.message}`);
    }
  }
);

// Tool for listing sections in a notebook
server.tool(
  "listSections",
  "List all sections, optionally scoped to a notebook",
  { notebookId: z.string().optional().describe("Optional notebook ID to scope sections to a specific notebook") },
  async ({ notebookId }) => {
    try {
      await ensureGraphClient();
      const api = notebookId
        ? graphClient.api(`/me/onenote/notebooks/${notebookId}/sections`)
        : graphClient.api(`/me/onenote/sections`);
      const response = await api.get();
      return {
        content: [
          { type: "text", text: JSON.stringify(response.value) }
        ]
      };
    } catch (error) {
      console.error("Error listing sections:", error);
      throw new Error(`Failed to list sections: ${error.message}`);
    }
  }
);

// Tool for listing pages in a section
server.tool(
  "listPages",
  "List all pages in a section",
  { sectionId: z.string().describe("The ID of the section to list pages from") },
  async ({ sectionId }) => {
    try {
      await ensureGraphClient();
      const response = await graphClient.api(`/me/onenote/sections/${sectionId}/pages`).get();
      return {
        content: [
          { type: "text", text: JSON.stringify(response.value) }
        ]
      };
    } catch (error) {
      console.error("Error listing pages:", error);
      throw new Error(`Failed to list pages: ${error.message}`);
    }
  }
);

// Tool for getting the content of a page
server.tool(
  "getPage",
  "Get the content of a page by its ID",
  { pageId: z.string().describe("The ID of the page to retrieve") },
  async ({ pageId }) => {
    try {
      await ensureGraphClient();
      const url = `https://graph.microsoft.com/v1.0/me/onenote/pages/${pageId}/content`;
      const response = await fetch(url, {
        headers: { 'Authorization': `Bearer ${accessToken}` }
      });
      if (!response.ok) {
        throw new Error(`HTTP error! Status: ${response.status} ${response.statusText}`);
      }
      const content = await response.text();
      return {
        content: [
          { type: "text", text: content }
        ]
      };
    } catch (error) {
      console.error("Error in getPage:", error);
      return {
        content: [
          { type: "text", text: `Error in getPage: ${error.message}` }
        ]
      };
    }
  }
);

// Tool for creating a new page in a section
server.tool(
  "createPage",
  "Create a new page in a section",
  {
    sectionId: z.string().describe("The ID of the section to create the page in"),
    title: z.string().describe("The title of the new page"),
    body: z.string().optional().describe("Optional HTML body content for the page"),
  },
  async ({ sectionId, title, body }) => {
    try {
      await ensureGraphClient();
      const htmlBody = body || '<p></p>';
      const html = `<!DOCTYPE html><html><head><title>${title}</title></head><body>${htmlBody}</body></html>`;
      const response = await graphClient
        .api(`/me/onenote/sections/${sectionId}/pages`)
        .header("Content-Type", "application/xhtml+xml")
        .post(html);
      return {
        content: [
          { type: "text", text: JSON.stringify(response) }
        ]
      };
    } catch (error) {
      console.error("Error creating page:", error);
      throw new Error(`Failed to create page: ${error.message}`);
    }
  }
);

// Tool for searching pages
server.tool(
  "searchPages",
  "Search for pages across notebooks by title",
  { query: z.string().describe("Search term to filter pages by title") },
  async ({ query }) => {
    try {
      await ensureGraphClient();
      const response = await graphClient.api(`/me/onenote/pages`).get();
      if (!query || query.length === 0) {
        return {
          content: [
            { type: "text", text: JSON.stringify(response.value) }
          ]
        };
      }
      const searchTerm = query.toLowerCase();
      const filteredPages = response.value.filter(page =>
        page.title && page.title.toLowerCase().includes(searchTerm)
      );
      return {
        content: [
          { type: "text", text: JSON.stringify(filteredPages) }
        ]
      };
    } catch (error) {
      console.error("Error searching pages:", error);
      throw new Error(`Failed to search pages: ${error.message}`);
    }
  }
);

// Connect to stdio and start server
async function main() {
  try {
    // Connect to standard I/O
    const transport = new StdioServerTransport();
    await server.connect(transport);
    
    console.error('Server started successfully.');
    console.error('Use the "authenticate" tool to start the authentication flow,');
    console.error('or use "saveAccessToken" if you already have a token.');
    
    // Keep the process alive
    process.on('SIGINT', () => {
      process.exit(0);
    });
  } catch (error) {
    console.error('Error starting server:', error);
    process.exit(1);
  }
}

export { server };

if (import.meta.url === `file://${process.argv[1]}`) {
  main();
} 
