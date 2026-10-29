import { InMemoryTransport } from "@modelcontextprotocol/sdk/inMemory.js";
import { Server } from "@modelcontextprotocol/sdk/server/index.js";
import { CallToolRequestSchema, ListToolsRequestSchema } from "@modelcontextprotocol/sdk/types.js";
import { fileToolNames, fileTools, runFileTool } from "./tools.js";
import { connectMcpToolSource } from "../mcp/tool-source.js";

export async function connectFileWorkspaceMcp({ fileWorkspace, officeFiles, userId, sendFile, shareLinks, publicOrigin }) {
  if (!fileWorkspace || !officeFiles || !userId) throw new Error("A scoped file workspace is required");
  const server = new Server(
    { name: "whatsapp-user-file-workspace", version: "1.0.0" },
    { capabilities: { tools: {} } },
  );

  server.setRequestHandler(ListToolsRequestSchema, async () => ({
    // Workspace-first policy: hide legacy spreadsheet/word tools from AI.
    // XLSX upload conversion remains available through workspace_files.share_link_from_file.
    tools: fileTools.filter(({ function: tool }) => tool.name === "workspace_files").map(({ function: tool }) => ({
      name: tool.name,
      description: tool.description,
      inputSchema: tool.parameters,
    })),
  }));

  server.setRequestHandler(CallToolRequestSchema, async (request) => {
    const { name, arguments: args = {} } = request.params;
    if (name !== "workspace_files") {
      return { content: [{ type: "text", text: `Unknown file tool: ${name}` }], isError: true };
    }
    try {
      const result = await runFileTool(name, args, { fileWorkspace, officeFiles, userId, sendFile, shareLinks, publicOrigin });
      return { content: [{ type: "text", text: JSON.stringify(result) }] };
    } catch (error) {
      return { content: [{ type: "text", text: error.message }], isError: true };
    }
  });

  const [clientTransport, serverTransport] = InMemoryTransport.createLinkedPair();
  const serverConnection = server.connect(serverTransport);
  try {
    const source = await connectMcpToolSource({
      namespace: "filesystem",
      transport: clientTransport,
      serverInfo: { name: "User-scoped file workspace" },
    });
    return {
      ...source,
      async close() {
        await source.close();
        await server.close().catch(() => {});
        await serverConnection.catch(() => {});
      },
    };
  } catch (error) {
    await server.close().catch(() => {});
    await serverConnection.catch(() => {});
    throw error;
  }
}