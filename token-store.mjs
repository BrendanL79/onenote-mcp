import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

let tokenFilePath = path.join(__dirname, '.access-token.txt');

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
