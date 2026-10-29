import assert from "node:assert/strict";
import { mkdtempSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import test from "node:test";
import { ExcelMcp } from "../src/mcp/excel.js";
import { OfficeFiles } from "../src/files/office.js";
import { FileWorkspace } from "../src/files/workspace.js";

test("Excel MCP reads an uploaded workbook by private workspace file ID", async (t) => {
  const dataDirectory = mkdtempSync(join(tmpdir(), "excel-mcp-"));
  const userId = "6281234567890";
  const fileWorkspace = new FileWorkspace({ dataDirectory });
  const officeFiles = new OfficeFiles(fileWorkspace);
  const excelMcp = new ExcelMcp();
  const workbook = await officeFiles.createSpreadsheet(userId, {
    filename: "ledger.xlsx",
    columns: ["Date", "Amount"],
    rows: [["2026-10-18", 12000]],
  });
  t.after(async () => {
    await excelMcp.close();
    fileWorkspace.close();
    rmSync(dataDirectory, { recursive: true, force: true });
  });

  const source = await excelMcp.connect({ fileWorkspace, userId });
  const describeTool = source.definitions.find((tool) => tool.function.name === "mcp_excel_excel_describe_sheets");
  assert.ok(describeTool);
  assert.ok(describeTool.function.parameters.properties.file_id);
  assert.equal(describeTool.function.parameters.required.includes("fileAbsolutePath"), false);

  const result = await source.call(describeTool.function.name, { file_id: workbook.id });
  assert.equal(result.isError, false);
  assert.match(result.content, /Sheet1/);
  assert.match(result.content, /A1:B2/);

  const absolutePath = await fileWorkspace.getFilePath(userId, workbook.id);
  const directResult = await source.call(describeTool.function.name, { fileAbsolutePath: absolutePath });
  assert.equal(directResult.isError, false);
  assert.match(directResult.content, /A1:B2/);

  await assert.rejects(
    source.call(describeTool.function.name, { file_id: "F-NOT-FOUND" }),
    /not found in this user's workspace/,
  );
});