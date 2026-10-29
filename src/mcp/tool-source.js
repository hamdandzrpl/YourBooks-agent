import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { CallToolResultSchema } from "@modelcontextprotocol/sdk/types.js";

const maxToolResultCharacters = 30_000;

function resultText(result) {
  const content = (result.content ?? []).map((item) => {
    if (item.type === "text") return item.text;
    if (item.type === "resource" && item.resource?.text) return item.resource.text;
    return `[${item.type} content omitted]`;
  }).join("\n");
  return {
    content: content.slice(0, maxToolResultCharacters),
    truncated: content.length > maxToolResultCharacters,
    isError: Boolean(result.isError),
  };
}

export async function connectMcpToolSource({ namespace, transport, serverInfo = {}, toolFilter = () => true }) {
  const client = new Client({ name: `whatsapp-agent-${namespace}`, version: "1.0.0" });
  try {
    await client.connect(transport);
    const listed = await client.listTools();
    const allowedTools = listed.tools.filter(toolFilter);
    const prefix = `mcp_${namespace}_`;
    const definitions = allowedTools.map((tool) => ({
      type: "function",
      function: {
        name: `${prefix}${tool.name}`,
        description: `${tool.description ?? `MCP tool ${tool.name}`} (MCP server: ${serverInfo.name ?? namespace})`,
        parameters: tool.inputSchema,
      },
    }));

    return {
      namespace,
      definitions,
      async call(functionName, args) {
        if (!functionName.startsWith(prefix)) throw new Error("MCP tool namespace mismatch");
        const name = functionName.slice(prefix.length);
        if (!allowedTools.some((tool) => tool.name === name)) throw new Error("MCP tool is not in the connected server's allowlist");
        const result = await client.callTool({ name, arguments: args }, CallToolResultSchema);
        return resultText(result);
      },
      async close() {
        await client.close().catch(() => {});
      },
    };
  } catch (error) {
    await client.close().catch(() => {});
    throw error;
  }
}