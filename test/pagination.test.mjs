import { describe, it } from 'node:test';
import assert from 'node:assert/strict';

describe('fetchAllPages', () => {
  it('concatenates multiple pages of results', async () => {
    const mod = await import(`../onenote-mcp.mjs?t=${Date.now()}`);
    const batches = [Array(100).fill({ id: 'a' }), Array(100).fill({ id: 'b' }), Array(42).fill({ id: 'c' })];
    const graphClient = {
      api: (path) => ({
        query: (q) => ({
          get: async () => {
            const skip = q.$skip || 0;
            const top = q.$top || 100;
            const index = skip / top;
            return { value: batches[index] || [] };
          },
        }),
      }),
    };
    const result = await mod.fetchAllPages(graphClient, '/me/onenote/sections/x/pages', 100);
    assert.equal(result.length, 242);
  });

  it('returns early on an empty batch', async () => {
    const mod = await import(`../onenote-mcp.mjs?t=${Date.now()}`);
    const graphClient = {
      api: () => ({
        query: () => ({
          get: async () => ({ value: [] }),
        }),
      }),
    };
    const result = await mod.fetchAllPages(graphClient, '/me/onenote/sections/x/pages', 100);
    assert.equal(result.length, 0);
  });
});
