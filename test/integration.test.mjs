import { describe, it } from 'node:test';
import assert from 'node:assert/strict';

const EXPECTED_TOOLS = [
  'authenticate',
  'saveAccessToken',
  'listNotebooks',
  'getNotebook',
  'listSections',
  'listPages',
  'getPage',
  'createPage',
  'searchPages',
];

describe('MCP server integration', () => {
  it('registers all expected tools', async () => {
    const mod = await import(`../onenote-mcp.mjs?t=${Date.now()}`);
    const tools = mod.server._registeredTools;
    for (const name of EXPECTED_TOOLS) {
      assert.ok(tools[name], `tool "${name}" should be registered`);
    }
  });

  it('every tool has an inputSchema', async () => {
    const mod = await import(`../onenote-mcp.mjs?t=${Date.now()}`);
    const tools = mod.server._registeredTools;
    for (const name of Object.keys(tools)) {
      assert.ok(tools[name].inputSchema, `tool "${name}" should have an inputSchema`);
    }
  });

  it('no tool handler references params.random_string', async () => {
    const mod = await import(`../onenote-mcp.mjs?t=${Date.now()}`);
    const tools = mod.server._registeredTools;
    for (const name of Object.keys(tools)) {
      const handlerStr = tools[name].handler?.toString() || '';
      assert.ok(
        !handlerStr.includes('random_string'),
        `tool "${name}" handler should not reference params.random_string`
      );
    }
  });
});
