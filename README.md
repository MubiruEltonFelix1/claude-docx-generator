# DOCX Generator — Paste & Produce

![Platform: Node.js](https://img.shields.io/badge/Platform-Node.js-339933.svg)
![Type: Local Tool](https://img.shields.io/badge/Type-Local%20Tool-blue.svg)
![Outputs: outputs/](https://img.shields.io/badge/Outputs-outputs%2F--ignored-lightgrey.svg)

Turn JavaScript that builds a `docx` Document into a real `.docx` file — without editing file paths.

# Why this Exists:
## Scenario: Claude Timed Out, But the DOCX Still Needs to Exist

You're working with an AI assistant to generate a 40-page report, proposal, thesis draft, or technical document.

The AI successfully writes the JavaScript that builds the DOCX using the `docx` package:

```js
const { Document, Packer, Paragraph } = require("docx");
// hundreds of lines of generated content...
```

But then one of these things happens:

* The AI session times out before producing the actual `.docx` file.
* The generated script contains an environment-specific path such as `/mnt/data/report.docx`.
* The code came from Claude, ChatGPT, Gemini, or another platform that saved files to a location that doesn't exist on your machine.
* You just want the document now—not another round of path fixing and debugging.

Instead of editing file paths throughout the script, you paste the generated code into `paste_here.js` and run:

```bash
node paste_here.js
```

This tool automatically intercepts the DOCX file write operation and redirects the output into the local `outputs/` directory.

The result: AI-generated DOCX code becomes a real `.docx` file on your computer in seconds, regardless of where the original script expected to save it.

In short: **when an AI can generate the document code but cannot deliver the document itself, this repository bridges the gap.**


Badges & quick facts

- **No publish required:** this is a local tool — run with Node.js.
- **Safe outputs:** `outputs/` is in `.gitignore` so generated DOCX files won't be committed accidentally.

Quick start

1. Install dependencies:

```bash
npm install claude-docx-generator
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
