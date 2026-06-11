<div align="center">

![DOCX Generator — Paste & Produce](./banner.svg)

<br/>

[![Node.js](https://img.shields.io/badge/Node.js-18%2B-339933?style=for-the-badge&logo=node.js&logoColor=white)](https://nodejs.org)
[![npm](https://img.shields.io/badge/npm-claude--docx--generator-CB3837?style=for-the-badge&logo=npm&logoColor=white)](https://www.npmjs.com/package/claude-docx-generator)
[![License: MIT](https://img.shields.io/badge/License-MIT-blue?style=for-the-badge)](./LICENSE)
[![PRs Welcome](https://img.shields.io/badge/PRs-welcome-brightgreen?style=for-the-badge&logo=github)](./CONTRIBUTING.md)
[![Zero Config](https://img.shields.io/badge/config-zero--required-555?style=for-the-badge)](https://github.com/MubiruEltonFelix1/claude-docx-generator)
[![outputs gitignored](https://img.shields.io/badge/outputs%2F-git--ignored-4CAF50?style=for-the-badge&logo=git&logoColor=white)](https://github.com/MubiruEltonFelix1/claude-docx-generator/blob/main/.gitignore)

<br/>

**Turn AI-generated JavaScript into a real `.docx` file — without touching a single file path.**

[Quick Start](#-quick-start) · [How It Works](#-how-it-works) · [File Reference](#-file-reference) · [Contributing](./CONTRIBUTING.md)

</div>

---

## 🧩 The Problem

You're deep in a session with Claude, ChatGPT, Gemini — or any AI assistant. You've asked it to generate a 40-page report, a proposal, or a thesis draft. It writes the whole thing as a JavaScript script using the [`docx`](https://www.npmjs.com/package/docx) package.

Then one of these things happens:

| Scenario | What goes wrong |
|---|---|
| ⏳ **Session timeout** | The AI wrote the code, but never produced the actual file |
| 🗂️ **Hardcoded paths** | Script writes to `/mnt/data/report.docx` — which doesn't exist on your machine |
| 🤖 **Platform differences** | Code came from ChatGPT or Claude artifacts — paths and environments don't match yours |
| 😩 **Debugging fatigue** | You just want the document now, not another round of path-fixing |

The AI can generate the *code*. This tool generates the *file*.

---

## ✨ The Solution

Paste the AI-generated code into `paste_here.js` and run one command:

```bash
node paste_here.js
```

`runner.js` **monkey-patches `fs.writeFileSync` and `fs.writeFile`** at runtime — silently intercepting any Buffer write (which is exactly how `docx` saves files) and redirecting the output to your local `outputs/` folder.

No path editing. No environment setup. No debugging.

```
AI generates code  →  paste into paste_here.js  →  node paste_here.js  →  ✅ outputs/your-doc.docx
```

---

## 🚀 Quick Start

**1. Install**

```bash
npm install claude-docx-generator
```

**2. Paste your code**

Open [`paste_here.js`](./paste_here.js) and paste in the AI-generated JavaScript. The file already `require`s `runner.js` at the top — that's the magic line, leave it alone.

```js
// paste_here.js — this line is already there, do not remove it
require('./runner.js');

// ↓ paste your AI-generated docx code below ↓
const { Document, Packer, Paragraph } = require("docx");
// ... hundreds of lines of generated content
```

**3. Run**

```bash
node paste_here.js
# or
npm run paste
```

Your `.docx` file appears in `outputs/` — no matter where the original script thought it was going.

---

## ⚙️ How It Works

Most AI-generated `docx` scripts follow this pattern:

```js
Packer.toBuffer(doc).then(buffer =>
  fs.writeFileSync('/some/absolute/path/report.docx', buffer)
);
```

`runner.js` patches `fs.writeFileSync` and `fs.writeFile` before your script runs. When those functions are called with a **Buffer** (the `docx` payload), the destination path is silently replaced:

```
/mnt/data/report.docx              →  outputs/report.docx
/Users/someone/Desktop/thesis.docx →  outputs/thesis.docx
C:\Users\whatever\doc.docx         →  outputs/doc.docx
```

The filename is preserved. The location is always `outputs/`.

---

## 📁 File Reference

| File | Purpose |
|---|---|
| [`paste_here.js`](./paste_here.js) | **Your workspace.** Paste AI-generated code here and run it. |
| [`runner.js`](./runner.js) | **The engine.** Patches `fs` write methods to redirect Buffer outputs into `outputs/`. |
| [`generate2.js`](./generate2.js) | Example generator — run it to test the setup. |
| [`generate2_to_docx.js`](./generate2_to_docx.js) | Another example demonstrating the full pipeline. |
| `outputs/` | **Where your files land.** Git-ignored — generated docs stay off your commit history. |

---

## 🧪 Try It

Not sure it works? Run the included example:

```bash
node generate2.js
```

This writes `MIT_Masterclass_Geometry_of_Meaning.docx` into `outputs/`. Open it and see a fully generated Word document — no cloud dependency, no AI session required.

---

## 💡 Best Practices

- **Keep `outputs/` lean.** It holds binary files. Git ignores it by design — don't fight that.
- **Leave the `require('./runner.js')` line at the top of `paste_here.js`.** That's the whole trick.
- **Using streams or custom Promise chains?** The current patch covers the most common `docx` patterns. If your script uses something unusual, [open an issue](https://github.com/MubiruEltonFelix1/claude-docx-generator/issues) and `runner.js` will be extended to handle it.

---

## 🤝 Contributing

Contributions are welcome — see [CONTRIBUTING.md](./CONTRIBUTING.md) for the full guide.

The short version: fork, branch, change, and open a PR. If you'd like a feature — CLI argument for destination, stream interception, or DOCX preview — file an issue describing it first.

---

## 📜 License

[MIT](./LICENSE) © 2024 Mubiru Elton Felix

---

<div align="center">

Built for the moment when the AI wrote the whole document — but the session ended before it saved the file.

</div>
