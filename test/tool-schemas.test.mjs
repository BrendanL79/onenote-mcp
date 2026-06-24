import { describe, it } from 'node:test';
import assert from 'node:assert/strict';

describe('saveAccessToken tool schema', () => {
  it('advertises a token parameter in its input schema', async () => {
    const mod = await import(`../onenote-mcp.mjs?t=${Date.now()}`);
    const tools = mod.server._registeredTools;
    const tool = tools.saveAccessToken;
    assert.ok(tool, 'saveAccessToken tool should be registered');
    assert.ok(tool.inputSchema, 'tool should have an inputSchema');
    const props = tool.inputSchema.def.shape;
    assert.ok(props.token, 'schema should include a token property');
  });
});

describe('listNotebooks tool schema', () => {
  it('is registered with an input schema', async () => {
    const mod = await import(`../onenote-mcp.mjs?t=${Date.now()}`);
    const tools = mod.server._registeredTools;
    const tool = tools.listNotebooks;
    assert.ok(tool?.inputSchema, 'listNotebooks should have an inputSchema');
  });
});

describe('getNotebook tool schema', () => {
  it('advertises a notebookId parameter', async () => {
    const mod = await import(`../onenote-mcp.mjs?t=${Date.now()}`);
    const tools = mod.server._registeredTools;
    const tool = tools.getNotebook;
    assert.ok(tool?.inputSchema);
    const props = tool.inputSchema.def.shape;
    assert.ok(props?.notebookId, 'schema should include notebookId');
  });
});

describe('listSections tool schema', () => {
  it('advertises an optional notebookId parameter', async () => {
    const mod = await import(`../onenote-mcp.mjs?t=${Date.now()}`);
    const tools = mod.server._registeredTools;
    const tool = tools.listSections;
    assert.ok(tool?.inputSchema);
    const props = tool.inputSchema.def.shape;
    assert.ok(props?.notebookId, 'schema should include notebookId');
  });
});
