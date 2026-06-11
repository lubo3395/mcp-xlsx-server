#!/usr/bin/env node
/**
 * xlsx-mcp-server - MCP server for Excel file read/write operations using SheetJS.
 *
 * Provides tools to list sheets, read data, create xlsx files, add sheets,
 * and update individual cells in Excel files.
 */

import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { registerReadTools } from "./tools/read.js";
import { registerWriteTools } from "./tools/write.js";

const server = new McpServer({
  name: "mcp-xlsx-server",
  version: "1.0.0",
});

registerReadTools(server);
registerWriteTools(server);

async function main() {
  const transport = new StdioServerTransport();
  await server.connect(transport);
  console.error("xlsx-mcp-server running via stdio");
}

main().catch((error) => {
  console.error("Server error:", error);
  process.exit(1);
});
