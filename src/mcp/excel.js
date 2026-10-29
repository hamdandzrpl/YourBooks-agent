import { createRequire } from "node:module";
import { readFileSync } from "node:fs";
import { dirname, extname, isAbsolute, resolve } from "node:path";
import { getDefaultEnvironment, StdioClientTransport } from "@modelcontextprotocol/sdk/client/stdio.js";
import { connectMcpToolSource } from "./tool-source.js";

const require = createRequire(import.meta.url);
const packagePath = require.resolve("@negokaz/excel-mcp-server/package.json");
const packageInfo = JSON.parse(readFileSync(packagePath, "utf8"));
const serverEntry = resolve(dirname(packagePath), packageInfo.bin["excel-mcp-server"]);
const supportedExtensions = new Set([".xlsx", ".xlsm", ".xltx", ".xltm"]);
const managedSheetMutations = new Map([
  ["mcp_excel_excel_write_to_sheet", ["sheetName"]],
  ["mcp_excel_excel_format_range", ["sheetName"]],
  ["mcp_excel_excel_create_table", ["sheetName"]],
  ["mcp_excel_excel_copy_sheet", ["dstSheetName"]],
]);

function normalizedAbsolutePath(path) {
  const normalized = resolve(path);
  return process.platform === "win32" ? normalized.toLowerCase() : normalized;
}

function withWorkspaceFileId(definition) {
  const adapted = structuredClone(definition);
  const properties = adapted.function.parameters?.properties;
  if (!properties?.fileAbsolutePath) return adapted;

  properties.file_id = {
    type: "string",
    description: "ID of an Excel file in the authenticated sender's private workspace. Prefer this for uploaded files.",
  };
  properties.fileAbsolutePath.description = "Absolute path to a file accessible to the Excel MCP process inside the server container.";
  adapted.function.parameters.required = (adapted.function.parameters.required ?? [])
    .filter((name) => name !== "fileAbsolutePath");
  return adapted;
}

export class ExcelMcp {
  sourcePromise = null;

  constructor({ financeStore = null } = {}) {
    this.financeStore = financeStore;
  }

  async connect({ fileWorkspace, userId }) {
    const source = await this.getSource();
    const excelMcp = this;
    return {
      namespace: source.namespace,
      definitions: source.definitions.map(withWorkspaceFileId),
      async call(functionName, input) {
        const args = { ...input };
        const fileId = args.file_id;
        const absolutePath = args.fileAbsolutePath;
        delete args.file_id;

        if (Boolean(fileId) === Boolean(absolutePath)) {
          throw new Error("Provide exactly one of file_id or fileAbsolutePath");
        }
        if (fileId) {
          const file = fileWorkspace.getFile(userId, fileId);
          if (!file || !supportedExtensions.has(extname(file.name).toLowerCase())) {
            throw new Error("Excel file was not found in this user's workspace or has an unsupported format");
          }
          args.fileAbsolutePath = await fileWorkspace.getFilePath(userId, fileId);
        } else if (typeof absolutePath !== "string" || !isAbsolute(absolutePath)) {
          throw new Error("fileAbsolutePath must be an absolute path accessible to the server container");
        }

        if (await excelMcp.isManagedSheetMutation({
          fileWorkspace,
          userId,
          fileId,
          absolutePath: args.fileAbsolutePath,
          functionName,
          input,
        })) {
          throw new Error("The Transaksi sheet in a finance book is managed by the ledger and cannot be edited through general Excel tools");
        }

        return source.call(functionName, args);
      },
    };
  }

  async isManagedSheetMutation({ fileWorkspace, userId, fileId, absolutePath, functionName, input }) {
    const sheetFields = managedSheetMutations.get(functionName);
    if (!sheetFields?.some((field) => typeof input[field] === "string"
      && input[field].toLocaleLowerCase("en-US") === "transaksi")) return false;
    const books = this.financeStore?.listBooks(userId) ?? [];
    for (const book of books) {
      if (!book.workbook_file_id) continue;
      if (fileId && fileId.toUpperCase() === book.workbook_file_id) return true;
      const linkedPath = await fileWorkspace.getFilePath(userId, book.workbook_file_id).catch(() => null);
      if (linkedPath && normalizedAbsolutePath(linkedPath) === normalizedAbsolutePath(absolutePath)) return true;
    }
    return false;
  }

  async close() {
    const source = await this.sourcePromise?.catch(() => null);
    await source?.close();
    this.sourcePromise = null;
  }

  getSource() {
    if (!this.sourcePromise) {
      const transport = new StdioClientTransport({
        command: process.execPath,
        args: [serverEntry],
        cwd: process.cwd(),
        env: {
          ...getDefaultEnvironment(),
          EXCEL_MCP_PAGING_CELLS_LIMIT: process.env.EXCEL_MCP_PAGING_CELLS_LIMIT?.trim() || "4000",
        },
        stderr: "inherit",
      });
      this.sourcePromise = connectMcpToolSource({
        namespace: "excel",
        transport,
        serverInfo: { name: "Excel MCP Server" },
        toolFilter: (tool) => tool.name !== "excel_screen_capture" || process.platform === "win32",
      }).catch((error) => {
        this.sourcePromise = null;
        throw error;
      });
    }
    return this.sourcePromise;
  }
}