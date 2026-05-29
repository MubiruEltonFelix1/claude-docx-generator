# DOCX Generator

This repository converts JavaScript code that builds a `docx` Document (using the `docx` npm package) into a `.docx` file.

Getting started

- Install dependencies:

```bash
npm install
```

- Paste your docx-generating code into `paste_here.js` (replace the example area). `paste_here.js` already requires `runner.js` which will redirect any Buffer file writes into `outputs/`.

- Run it:

```bash
node paste_here.js
# or using npm
npm run paste
```

Behavior

- `runner.js` intercepts `fs.writeFileSync` and `fs.writeFile` calls that write a Buffer (typical for `Packer.toBuffer(doc).then(buffer => fs.writeFileSync(path, buffer))`) and redirects them into the `outputs/` folder in this repository. This avoids needing to manually edit absolute paths produced by code from other platforms.

Files

- `paste_here.js` — template where you paste the docx-generating code.
- `runner.js` — path-fixer/interceptor that ensures outputs go to `outputs/`.
- `generate2.js`, `generate2_to_docx.js` — existing examples.
- `outputs/` — output files (ignored by git via `.gitignore`).

Notes

- The interceptor redirects writes that include a Buffer. If your code writes streams or uses other mechanisms, let me know and I can extend the interceptor to cover those cases.
