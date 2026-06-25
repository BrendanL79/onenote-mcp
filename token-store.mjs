import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import fetch from 'node-fetch';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

let tokenFilePath = path.join(__dirname, '.access-token.txt');
let endpointConfig = { clientId: null, tenant: 'common', scopes: '' };

export function setTokenFilePath(p) {
  tokenFilePath = p;
}

export function getTokenFilePath() {
  return tokenFilePath;
}

export function saveToken({ token, refresh_token, expires_in }) {
  if (!token || token.length === 0) {
    throw new Error('token is required');
  }
  const data = { token };
  if (refresh_token) data.refresh_token = refresh_token;
  if (expires_in) data.expires_at = Date.now() + expires_in * 1000;
  fs.writeFileSync(tokenFilePath, JSON.stringify(data));
}

export function loadToken() {
  if (!fs.existsSync(tokenFilePath)) return null;
  const raw = fs.readFileSync(tokenFilePath, 'utf8');
  try {
    const parsed = JSON.parse(raw);
    if (parsed && parsed.token) {
      return { token: parsed.token, refresh_token: parsed.refresh_token, expires_at: parsed.expires_at };
    }
    return { token: raw.trim() };
  } catch {
    return { token: raw.trim() };
  }
}

export function clearToken() {
  if (fs.existsSync(tokenFilePath)) {
    fs.unlinkSync(tokenFilePath);
  }
}

export function setTokenEndpointConfig({ clientId, tenant, scopes }) {
  endpointConfig = { clientId, tenant: tenant || 'common', scopes };
}

export function isTokenExpired() {
  const t = loadToken();
  if (!t || !t.expires_at) return true;
  const skewMs = 5 * 60 * 1000;
  return Date.now() + skewMs >= t.expires_at;
}

export async function refreshAccessToken() {
  const t = loadToken();
  if (!t || !t.refresh_token) return null;

  const body = new URLSearchParams({
    grant_type: 'refresh_token',
    client_id: endpointConfig.clientId,
    refresh_token: t.refresh_token,
    scope: endpointConfig.scopes,
  }).toString();

  const res = await (globalThis.fetch || fetch)(
    `https://login.microsoftonline.com/${endpointConfig.tenant}/oauth2/v2.0/token`,
    { method: 'POST', headers: { 'Content-Type': 'application/x-www-form-urlencoded' }, body }
  );
  const data = await res.json();
  if (data.error) throw new Error(`${data.error}: ${data.error_description}`);

  saveToken({
    token: data.access_token,
    expires_in: data.expires_in,
  });
  if (data.refresh_token) {
    saveToken({
      token: data.access_token,
      refresh_token: data.refresh_token,
      expires_in: data.expires_in,
    });
  }

  return loadToken();
}
