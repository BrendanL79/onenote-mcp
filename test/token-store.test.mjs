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

describe('isTokenExpired', () => {
  it('returns true when expires_at is missing', async () => {
    const { saveToken, isTokenExpired } = await importTokenStore();
    saveToken({ token: 'abc' });
    assert.equal(isTokenExpired(), true);
  });

  it('returns true when expires_at is in the past', async () => {
    const { saveToken, isTokenExpired } = await importTokenStore();
    saveToken({ token: 'abc', expires_in: -1 });
    assert.equal(isTokenExpired(), true);
  });

  it('returns false when expires_at is in the future (beyond 5-min skew)', async () => {
    const { saveToken, isTokenExpired } = await importTokenStore();
    saveToken({ token: 'abc', expires_in: 600 });
    assert.equal(isTokenExpired(), false);
  });

  it('returns true when expires_at is within 5-min skew window', async () => {
    const { saveToken, isTokenExpired } = await importTokenStore();
    saveToken({ token: 'abc', expires_in: 120 });
    assert.equal(isTokenExpired(), true);
  });
});

describe('refreshAccessToken', () => {
  it('returns null when no refresh_token is stored', async () => {
    const { saveToken, refreshAccessToken } = await importTokenStore();
    saveToken({ token: 'abc' });
    const result = await refreshAccessToken();
    assert.equal(result, null);
  });

  it('returns null when no token file exists', async () => {
    const { refreshAccessToken } = await importTokenStore();
    const result = await refreshAccessToken();
    assert.equal(result, null);
  });

  it('calls the token endpoint and saves the new token', async () => {
    const { saveToken, refreshAccessToken, loadToken, setTokenEndpointConfig } = await importTokenStore();

    let capturedBody;
    const originalFetch = globalThis.fetch;
    globalThis.fetch = async (url, opts) => {
      capturedBody = opts.body;
      return {
        ok: true,
        json: async () => ({
          access_token: 'new-access',
          refresh_token: 'new-refresh',
          expires_in: 3600,
        }),
      };
    };
    try {
      setTokenEndpointConfig({ clientId: 'test-client', tenant: 'common', scopes: 'test-scope' });
      saveToken({ token: 'old-access', refresh_token: 'old-refresh', expires_in: -1 });
      const result = await refreshAccessToken();
      assert.equal(result.token, 'new-access');
      assert.equal(result.refresh_token, 'new-refresh');
      assert.ok(result.expires_at > Date.now());
      const loaded = loadToken();
      assert.equal(loaded.token, 'new-access');
      const params = new URLSearchParams(capturedBody);
      assert.equal(params.get('grant_type'), 'refresh_token');
      assert.equal(params.get('client_id'), 'test-client');
      assert.equal(params.get('refresh_token'), 'old-refresh');
    } finally {
      globalThis.fetch = originalFetch;
    }
  });
});
