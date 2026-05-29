const fs = require("fs");
const path = require("path");
const { Document, Packer, Paragraph, TextRun } = require("docx");

const src = path.join(__dirname, "generate2.js");
const outDir = path.join(__dirname, "outputs");
if (!fs.existsSync(outDir)) fs.mkdirSync(outDir, { recursive: true });
const dest = path.join(outDir, "generate2.docx");

if (!fs.existsSync(src)) {
  console.error("Source file not found:", src);
  process.exit(1);
}

const code = fs.readFileSync(src, "utf8");
const lines = code.replace(/\r\n/g, "\n").split("\n");

const children = [];

children.push(new Paragraph({ children: [new TextRun({ text: "generate2.js", bold: true, size: 28 })] }));
children.push(new Paragraph({ children: [new TextRun({ text: "", size: 8 })] }));

for (const line of lines) {
  if (line === "") {
    children.push(new Paragraph({ children: [new TextRun({ text: "" })] }));
    continue;
  }
  children.push(new Paragraph({ children: [new TextRun({ text: line, font: "Courier New", size: 20 })] }));
}

const doc = new Document({ sections: [{ properties: {}, children }] });

Packer.toBuffer(doc).then((buffer) => {
  fs.writeFileSync(dest, buffer);
  console.log("Wrote", dest);
}).catch((err) => {
  console.error(err);
  process.exit(1);
});
