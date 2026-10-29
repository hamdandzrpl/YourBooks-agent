/** HTML shell for the shared workspace: Excel-like layout, assets served
 *  separately so the browser can cache them. */

function escapeHtml(value) {
  return String(value ?? "").replace(/[&<>"']/g, (character) => ({
    "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;",
  })[character]);
}

/** Serialize JSON so it cannot break out of the inline <script> block. */
function scriptJson(value) {
  return JSON.stringify(value)
    .replace(/</g, "\\u003c")
    .replace(/\u2028/g, "\\u2028")
    .replace(/\u2029/g, "\\u2029");
}

export function renderWorkspacePage(share, workspaceApiBase, initialData) {
  const title = escapeHtml(share.title || share.name);
  const bootstrap = {
    base: workspaceApiBase,
    title: share.title || share.name,
    data: initialData,
  };  return `<!doctype html>
<html lang="id">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<meta name="robots" content="noindex">
<title>${title} - Workspace</title>
<link rel="stylesheet" href="${workspaceApiBase}/app.css">
<link rel="stylesheet" href="${workspaceApiBase}/analysis.css">
</head>
<body>
<div class="app">
  <header class="topbar">
    <div class="doc">
      <span class="doc-icon" aria-hidden="true">
        <svg viewBox="0 0 16 16" width="15" height="15" fill="none" stroke="currentColor" stroke-width="1.4">
          <rect x="2" y="2" width="12" height="12" rx="1.6"/>
          <path d="M2 6h12M2 10h12M6 2v12M10 2v12"/>
        </svg>
      </span>
      <h1>${title}</h1>
    </div>
    <div class="top-actions">
      <div class="mode-switch" role="tablist" aria-label="Tampilan">
        <button type="button" id="mode-analysis" role="tab" aria-selected="true">Workspace</button>
        <button type="button" id="mode-sheet" role="tab" aria-selected="false">Sheet</button>
      </div>
      <button type="button" id="btn-import">Import .xlsx</button>
      <button type="button" id="btn-export">Export .xlsx</button>
      <button type="button" id="btn-save" class="primary">Simpan</button>
    </div>
  </header>

  <div class="toolbar">
    <button type="button" id="btn-add-row">+ Baris</button>
    <button type="button" id="btn-add-col">+ Kolom</button>
    <span class="divider"></span>
    <select id="font-family" class="tb-select" title="Jenis huruf" aria-label="Jenis huruf">
      <option value="">Font</option>
      <option value="Calibri">Calibri</option>
      <option value="Arial">Arial</option>
      <option value="Times New Roman">Times New Roman</option>
      <option value="Verdana">Verdana</option>
      <option value="Georgia">Georgia</option>
      <option value="Courier New">Courier New</option>
      <option value="Tahoma">Tahoma</option>
      <option value="Trebuchet MS">Trebuchet MS</option>
      <option value="Segoe UI">Segoe UI</option>
    </select>
    <select id="font-size" class="tb-select narrow" title="Ukuran huruf" aria-label="Ukuran huruf">
      <option value="">Ukuran</option>
      <option value="8">8</option>
      <option value="9">9</option>
      <option value="10">10</option>
      <option value="11">11</option>
      <option value="12">12</option>
      <option value="14">14</option>
      <option value="16">16</option>
      <option value="18">18</option>
      <option value="20">20</option>
      <option value="24">24</option>
      <option value="28">28</option>
      <option value="36">36</option>
    </select>
    <span class="divider"></span>
    <button type="button" id="cmd-bold" class="tb-toggle" aria-pressed="false" title="Tebal (Ctrl+B)"><b>B</b></button>
    <button type="button" id="cmd-italic" class="tb-toggle" aria-pressed="false" title="Miring (Ctrl+I)"><i>I</i></button>
    <button type="button" id="cmd-underline" class="tb-toggle" aria-pressed="false" title="Garis bawah (Ctrl+U)"><u>U</u></button>
    <button type="button" id="cmd-strike" class="tb-toggle" aria-pressed="false" title="Coret"><s>S</s></button>
    <span class="divider"></span>
    <span class="tb-color" title="Warna isi sel">
      <button type="button" id="fill-button" class="tb-toggle" aria-pressed="false" aria-label="Warna isi sel">
        <svg class="tb-icon" viewBox="0 0 16 16" width="13" height="13" fill="none" stroke="currentColor" stroke-width="1.4">
          <path d="M3 11.5 8.5 6l4 4-4 4H4.5z"/><path d="M8.5 6 11 3.5"/>
        </svg>
        <span class="swatch" id="fill-swatch"></span>
      </button>
      <input type="color" id="fill-color" value="#FFF2CC" aria-label="Pilih warna isi">
    </span>
    <span class="tb-color" title="Warna teks">
      <button type="button" id="color-button" class="tb-toggle" aria-pressed="false" aria-label="Warna teks">
        <span class="swatch text" id="color-swatch">A</span>
      </button>
      <input type="color" id="text-color" value="#C00000" aria-label="Pilih warna teks">
    </span>
    <button type="button" id="cmd-clear-fill" class="tb-plain" title="Hapus warna isi">Tanpa isi</button>
    <span class="divider"></span>
    <button type="button" class="tb-toggle" data-align="left" aria-pressed="false" title="Rata kiri" aria-label="Rata kiri">
      <svg class="tb-icon" viewBox="0 0 16 16" width="14" height="14" stroke="currentColor" stroke-width="1.5" stroke-linecap="round">
        <path d="M2.5 4h11M2.5 8h7M2.5 12h11"/>
      </svg>
    </button>
    <button type="button" class="tb-toggle" data-align="center" aria-pressed="false" title="Rata tengah" aria-label="Rata tengah">
      <svg class="tb-icon" viewBox="0 0 16 16" width="14" height="14" stroke="currentColor" stroke-width="1.5" stroke-linecap="round">
        <path d="M2.5 4h11M4.5 8h7M2.5 12h11"/>
      </svg>
    </button>
    <button type="button" class="tb-toggle" data-align="right" aria-pressed="false" title="Rata kanan" aria-label="Rata kanan">
      <svg class="tb-icon" viewBox="0 0 16 16" width="14" height="14" stroke="currentColor" stroke-width="1.5" stroke-linecap="round">
        <path d="M2.5 4h11M6.5 8h7M2.5 12h11"/>
      </svg>
    </button>
    <span class="divider"></span>
    <button type="button" id="cmd-reset" class="tb-plain" title="Kembalikan format sel">Reset</button>
    <span class="spacer"></span>
    <span class="status" id="status"></span>
  </div>

  <div class="formula-row">
    <div class="name-box" id="name-box" contenteditable="true" spellcheck="false" title="Kotak Nama">A1</div>
    <div class="fx" aria-hidden="true">fx</div>
    <div class="formula-input" id="formula-input" contenteditable="true" spellcheck="false"
         role="textbox" aria-label="Bilah rumus"></div>
  </div>

  <div class="grid-wrap" id="grid-wrap">
    <table class="sheet" id="sheet"><thead id="col-head"></thead><tbody id="body"></tbody></table>
  </div>

  <div class="analysis" id="analysis">
    <div class="aw-status" id="aw-status"></div>
    <div class="aw-meta" id="aw-meta"></div>
    <div class="aw-grid" id="aw-grid"></div>
    <div class="aw-insight" id="aw-insight" hidden></div>
    <details class="aw-tools">
      <summary>Atur analisis &amp; sumber data</summary>
      <div class="aw-tools-body">
        <div class="aw-prompt">
          <input type="text" id="aw-input"
                 placeholder="Contoh: untung bersih per bulan, pengeluaran per kategori"
                 aria-label="Permintaan analisis">
          <button type="button" id="aw-generate" class="primary">Buat widget</button>
          <button type="button" id="aw-auto" title="Kembalikan susunan widget standar">Standar</button>
        </div>
        <div class="aw-source" id="aw-source"></div>
      </div>
    </details>
  </div>

  <div class="tabbar">
    <button type="button" id="btn-add-sheet" title="Tambah sheet">+</button>
    <div class="tabs" id="tabs" role="tablist"></div>
  </div>

  <div class="notice" id="notice" role="status" aria-live="polite"></div>

  <div class="menu" id="menu" role="menu">
    <button type="button" data-action="row-above">Sisipkan baris di atas</button>
    <button type="button" data-action="row-below">Sisipkan baris di bawah</button>
    <button type="button" data-action="row-delete">Hapus baris</button>
    <hr>
    <button type="button" data-action="col-left">Sisipkan kolom di kiri</button>
    <button type="button" data-action="col-right">Sisipkan kolom di kanan</button>
    <button type="button" data-action="col-delete">Hapus kolom</button>
    <hr>
    <button type="button" data-action="clear">Kosongkan sel</button>
    <hr>
    <button type="button" data-action="rename-sheet">Ganti nama sheet</button>
    <button type="button" data-action="delete-sheet">Hapus sheet</button>
  </div>
</div>
<script>window.__WORKSPACE__ = ${scriptJson(bootstrap)};</script>
<script src="${workspaceApiBase}/analysis.js"></script>
<script src="${workspaceApiBase}/app.js"></script>
</body>
</html>`;
}
