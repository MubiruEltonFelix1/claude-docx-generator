# DOCX Generator — Paste & Produce

![Platform: Node.js](https://img.shields.io/badge/Platform-Node.js-339933.svg)
![Type: Local Tool](https://img.shields.io/badge/Type-Local%20Tool-blue.svg)
![Outputs: outputs/](https://img.shields.io/badge/Outputs-outputs%2F--ignored-lightgrey.svg)

Turn JavaScript that builds a `docx` Document into a real `.docx` file — without editing file paths.

Why this exists

- People often copy/paste docx-generating scripts from different platforms. Those scripts write absolute paths (e.g. `/mnt/.../file.docx`) which are cumbersome to change every time.
- This repo provides a tiny path-fixer (`runner.js`) so you can paste any `docx`-producing JS into a single file and run it. The resulting `.docx` is automatically saved into `outputs/`.

Badges & quick facts

- **No publish required:** this is a local tool — run with Node.js.
- **Safe outputs:** `outputs/` is in `.gitignore` so generated DOCX files won't be committed accidentally.

Quick start

1. Install dependencies:

```bash
npm install
```

2. Paste your docx-generating code into [paste_here.js](paste_here.js). `paste_here.js` already `require`s [runner.js](runner.js) which redirects Buffer writes into `outputs/`.

3. Run:

```bash
node paste_here.js
# or
npm run paste
```

File references

- [paste_here.js](paste_here.js) — development template where you paste code.
- [runner.js](runner.js) — intercepts `fs.writeFileSync`/`fs.writeFile` Buffer writes and redirects them into `outputs/`.
- [generate2.js](generate2.js) and [generate2_to_docx.js](generate2_to_docx.js) — example generators present in the repo.

How it works (brief)

- Many `docx` scripts follow the pattern:

```js
Packer.toBuffer(doc).then(buffer => fs.writeFileSync('/some/absolute/path/my.docx', buffer));
```

- `runner.js` monkey-patches `fs.writeFileSync` and `fs.writeFile` at runtime. When the patched functions detect a Buffer (the `docx` payload), they replace the destination path with `outputs/<basename>` so you don't have to edit the original code.

Best practices

- Keep `outputs/` small — it contains generated binary files and is git-ignored.
- If your pasted script writes streams, custom promises, or uses other FS APIs, open an issue or ask me to extend `runner.js` to catch that pattern.

Examples

- To test quickly, run the included example:

```bash
node generate2.js
```

- The example will write `MIT_Masterclass_Geometry_of_Meaning.docx` into `outputs/`.

Contributing

- This is a tiny local tool. If you want features such as a CLI argument for destination, stream interception, or previewing generated docx, tell me which feature and I will add it.

License

- No license specified. Add a `LICENSE` file if you want to open-source this project.
