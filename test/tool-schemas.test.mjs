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

describe('listPages tool schema', () => {
  it('advertises a required sectionId parameter', async () => {
    const mod = await import(`../onenote-mcp.mjs?t=${Date.now()}`);
    const tools = mod.server._registeredTools;
    const tool = tools.listPages;
    assert.ok(tool?.inputSchema);
    const props = tool.inputSchema.def.shape;
    assert.ok(props?.sectionId, 'schema should include sectionId');
  });
});

describe('getPage tool schema', () => {
  it('advertises a required pageId parameter', async () => {
    const mod = await import(`../onenote-mcp.mjs?t=${Date.now()}`);
    const tools = mod.server._registeredTools;
    const tool = tools.getPage;
    assert.ok(tool?.inputSchema);
    const props = tool.inputSchema.def.shape;
    assert.ok(props?.pageId, 'schema should include pageId');
  });
});

describe('createPage tool schema', () => {
  it('advertises sectionId and title parameters', async () => {
    const mod = await import(`../onenote-mcp.mjs?t=${Date.now()}`);
    const tools = mod.server._registeredTools;
    const tool = tools.createPage;
    assert.ok(tool?.inputSchema);
    const props = tool.inputSchema.def.shape;
    assert.ok(props?.sectionId, 'schema should include sectionId');
    assert.ok(props?.title, 'schema should include title');
  });
});

describe('searchPages tool schema', () => {
  it('advertises a query parameter', async () => {
    const mod = await import(`../onenote-mcp.mjs?t=${Date.now()}`);
    const tools = mod.server._registeredTools;
    const tool = tools.searchPages;
    assert.ok(tool?.inputSchema);
    const props = tool.inputSchema.def.shape;
    assert.ok(props?.query, 'schema should include query');
  });
});

describe('authenticate tool schema', () => {
  it('is registered with an explicit (empty) input schema', async () => {
    const mod = await import(`../onenote-mcp.mjs?t=${Date.now()}`);
    const tools = mod.server._registeredTools;
    const tool = tools.authenticate;
    assert.ok(tool?.inputSchema, 'authenticate should have an explicit inputSchema');
  });
});

describe('pagination schema params', () => {
  it('listPages advertises optional top and skip params', async () => {
    const mod = await import(`../onenote-mcp.mjs?t=${Date.now()}`);
    const tool = mod.server._registeredTools.listPages;
    const props = tool.inputSchema.def.shape;
    assert.ok(props.top, 'listPages should include top param');
    assert.ok(props.skip, 'listPages should include skip param');
  });

  it('searchPages advertises optional top and skip params', async () => {
    const mod = await import(`../onenote-mcp.mjs?t=${Date.now()}`);
    const tool = mod.server._registeredTools.searchPages;
    const props = tool.inputSchema.def.shape;
    assert.ok(props.top, 'searchPages should include top param');
    assert.ok(props.skip, 'searchPages should include skip param');
  });
});
