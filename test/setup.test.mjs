import { test } from 'node:test';
import assert from 'node:assert/strict';

test('should be importable without side-effects', async () => {
  const module = await import('../onenote-mcp.mjs');
  assert.ok(module.server, 'server should be exported');
});
