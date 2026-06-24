import { describe, it, beforeEach, afterEach } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import os from 'node:os';

let tempDir;
let tokenFilePath;

beforeEach(() => {
  tempDir = fs.mkdtempSync(path.join(os.tmpdir(), 'token-test-'));
  tokenFilePath = path.join(tempDir, '.access-token.txt');
});

afterEach(() => {
  fs.rmSync(tempDir, { recursive: true, force: true });
});

async function importTokenStore() {
  const mod = await import(`../token-store.mjs?t=${Date.now()}`);
  mod.setTokenFilePath(tokenFilePath);
  return mod;
}

describe('saveToken', () => {
  it('saves token with refresh_token and expires_at', async () => {
    const { saveToken, loadToken } = await importTokenStore();
    saveToken({ token: 'abc123', refresh_token: 'refresh456', expires_in: 3600 });
    const loaded = loadToken();
    assert.equal(loaded.token, 'abc123');
    assert.equal(loaded.refresh_token, 'refresh456');
    assert.ok(loaded.expires_at > Date.now());
    assert.ok(loaded.expires_at <= Date.now() + 3600 * 1000 + 1000);
  });

  it('saves token without refresh_token or expires_in', async () => {
    const { saveToken, loadToken } = await importTokenStore();
    saveToken({ token: 'abc123' });
    const loaded = loadToken();
    assert.equal(loaded.token, 'abc123');
    assert.equal(loaded.refresh_token, undefined);
    assert.equal(loaded.expires_at, undefined);
  });

  it('throws on empty/undefined token', async () => {
    const { saveToken } = await importTokenStore();
    assert.throws(() => saveToken({ token: undefined }), /token is required/);
    assert.throws(() => saveToken({ token: '' }), /token is required/);
  });
});

describe('loadToken', () => {
  it('returns null when no file exists', async () => {
    const { loadToken } = await importTokenStore();
    assert.equal(loadToken(), null);
  });

  it('reads old {token} JSON format', async () => {
    fs.writeFileSync(tokenFilePath, JSON.stringify({ token: 'legacy' }));
    const { loadToken } = await importTokenStore();
    const loaded = loadToken();
    assert.equal(loaded.token, 'legacy');
  });

  it('reads raw string format (pre-JSON)', async () => {
    fs.writeFileSync(tokenFilePath, 'rawtoken123');
    const { loadToken } = await importTokenStore();
    const loaded = loadToken();
    assert.equal(loaded.token, 'rawtoken123');
  });
});

describe('clearToken', () => {
  it('deletes the token file', async () => {
    fs.writeFileSync(tokenFilePath, JSON.stringify({ token: 'abc' }));
    const { clearToken, loadToken } = await importTokenStore();
    clearToken();
    assert.equal(loadToken(), null);
  });

  it('does not throw when file does not exist', async () => {
    const { clearToken } = await importTokenStore();
    assert.doesNotThrow(() => clearToken());
  });
});
