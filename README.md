# Docxidian

Docxidian opens `.docx` files directly inside Obsidian and saves edits back to the same vault file.

The editor integration is based on `@eigenpal/docx-editor-react`, adapted from its Vite/React usage pattern to Obsidian's plugin runtime:

- Obsidian loads the DOCX file with `vault.readBinary`.
- React renders `DocxEditor` only after the document buffer is available.
- Editor CSS is bundled as text and injected at runtime.
- The editor saves through its ref API and writes the resulting `ArrayBuffer` with `vault.modifyBinary`.

## Development

Install dependencies:

```bash
npm install
```

Run the esbuild watcher:

```bash
npm run dev
```

Create a production build:

```bash
npm run build
```

## Usage

Enable the plugin in **Settings -> Community plugins**, then open a `.docx` file from the vault. Obsidian will route it to Docxidian's custom file view.

Use the editor toolbar's save action or the command **Save current docx** to write changes back to the open file.
