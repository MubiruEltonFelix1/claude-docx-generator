const {
    Document,
    Packer,
    Paragraph,
    TextRun,
    Table,
    TableRow,
    TableCell,
    Header,
    Footer,
    AlignmentType,
    HeadingLevel,
    BorderStyle,
    WidthType,
    ShadingType,
    VerticalAlign,
    PageNumber,
    PageBreak,
    LevelFormat,
    TableOfContents
} = require('docx');
const fs = require('fs');
const path = require('path');

// ─── COLOUR PALETTE ───────────────────────────────────────────────────────────
const C = {
    mitRed: "A31F34",
    deepNavy: "1B2A4A",
    slateBlue: "2E4A7A",
    skyBlue: "4A90D9",
    softGold: "C8972A",
    warmGold: "F5A623",
    lightGold: "FFF3CD",
    codeGray: "F4F4F4",
    codeBorder: "CCCCCC",
    tableHeader: "1B2A4A",
    tableRow1: "EEF2F8",
    tableRow2: "F9FAFB",
    white: "FFFFFF",
    lightRed: "FDF0F0",
    lightBlue: "EEF4FB",
    lightGreen: "F0F7F0",
    softOrange: "FFF8EE",
    darkText: "1A1A1A",
    mutedText: "555555",
};

const border = (color = "CCCCCC") => ({ style: BorderStyle.SINGLE, size: 1, color });
const borders = (color = "CCCCCC") => ({
    top: border(color),
    bottom: border(color),
    left: border(color),
    right: border(color)
});
const noBorder = () => ({
    top: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
    bottom: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
    left: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
    right: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
});

// ─── HELPERS ──────────────────────────────────────────────────────────────────
const gap = (pt = 120) => new Paragraph({ spacing: { before: 0, after: pt } });

function h1(text) {
    return new Paragraph({
        heading: HeadingLevel.HEADING_1,
        children: [new TextRun({ text, font: "Arial", size: 36, bold: true, color: C.deepNavy })],
        spacing: { before: 480, after: 200 },
        border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: C.mitRed, space: 4 } }
    });
}

function h2(text) {
    return new Paragraph({
        heading: HeadingLevel.HEADING_2,
        children: [new TextRun({ text, font: "Arial", size: 28, bold: true, color: C.slateBlue })],
        spacing: { before: 360, after: 160 }
    });
}

function h3(text) {
    return new Paragraph({
        heading: HeadingLevel.HEADING_3,
        children: [new TextRun({ text, font: "Arial", size: 24, bold: true, color: C.softGold })],
        spacing: { before: 240, after: 120 }
    });
}

function body(text, opts = {}) {
    return new Paragraph({
        children: [new TextRun({
            text,
            font: "Arial",
            size: 22,
            color: opts.color || C.darkText,
            bold: opts.bold || false,
            italics: opts.italic || false,
        })],
        spacing: { before: 80, after: 100 },
        alignment: opts.justify ? AlignmentType.JUSTIFIED : AlignmentType.LEFT,
    });
}

function mixedPara(runs, opts = {}) {
    return new Paragraph({
        children: runs,
        spacing: { before: opts.before || 80, after: opts.after || 100 },
        alignment: opts.justify ? AlignmentType.JUSTIFIED : AlignmentType.LEFT,
        ...(opts.numbering ? { numbering: opts.numbering } : {})
    });
}

function run(text, opts = {}) {
    return new TextRun({
        text,
        font: opts.mono ? "Courier New" : "Arial",
        size: opts.size || 22,
        bold: opts.bold || false,
        italics: opts.italic || false,
        color: opts.color || C.darkText,
        highlight: opts.highlight || undefined,
    });
}

function codeBlock(lines) {
    const rows = lines.map(line =>
        new TableRow({
            children: [
                new TableCell({
                    borders: noBorder(),
                    width: { size: 9200, type: WidthType.DXA },
                    margins: { top: 40, bottom: 40, left: 160, right: 160 },
                    children: [new Paragraph({
                        children: [new TextRun({ text: line, font: "Courier New", size: 18, color: C.deepNavy })],
                        spacing: { before: 0, after: 20 }
                    })]
                })
            ]
        })
    );
    return new Table({
        width: { size: 9200, type: WidthType.DXA },
        columnWidths: [9200],
        borders: {
            top: border(C.codeBorder),
            bottom: border(C.codeBorder),
            left: { style: BorderStyle.SINGLE, size: 12, color: C.slateBlue },
            right: border(C.codeBorder),
            insideH: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
            insideV: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
        },
        rows,
        margins: { top: 60, bottom: 60 },
    });
}

function calloutBox(title, lines, bgColor, accentColor) {
    const contentParagraphs = lines.map(l =>
        new Paragraph({
            children: [new TextRun({ text: l, font: "Arial", size: 21, color: C.darkText })],
            spacing: { before: 40, after: 40 }
        })
    );
    return new Table({
        width: { size: 9360, type: WidthType.DXA },
        columnWidths: [280, 9080],
        rows: [
            new TableRow({
                children: [
                    new TableCell({
                        borders: noBorder(),
                        shading: { fill: accentColor, type: ShadingType.CLEAR },
                        width: { size: 280, type: WidthType.DXA },
                        margins: { top: 80, bottom: 80, left: 0, right: 0 },
                        children: [new Paragraph({ children: [] })]
                    }),
                    new TableCell({
                        borders: noBorder(),
                        shading: { fill: bgColor, type: ShadingType.CLEAR },
                        width: { size: 9080, type: WidthType.DXA },
                        margins: { top: 100, bottom: 100, left: 160, right: 160 },
                        children: [
                            new Paragraph({
                                children: [new TextRun({ text: title, font: "Arial", size: 22, bold: true, color: accentColor })],
                                spacing: { before: 0, after: 60 }
                            }),
                            ...contentParagraphs
                        ]
                    })
                ]
            })
        ]
    });
}

function twoColTable(headers, rows, colWidths = [3120, 3120, 3120]) {
    const headerRow = new TableRow({
        children: headers.map((h, i) => new TableCell({
            borders: borders(C.slateBlue),
            shading: { fill: C.tableHeader, type: ShadingType.CLEAR },
            width: { size: colWidths[i], type: WidthType.DXA },
            margins: { top: 80, bottom: 80, left: 120, right: 120 },
            children: [new Paragraph({
                children: [new TextRun({ text: h, font: "Arial", size: 20, bold: true, color: C.white })],
                alignment: AlignmentType.CENTER
            })]
        }))
    });
    const dataRows = rows.map((row, ri) =>
        new TableRow({
            children: row.map((cell, ci) => new TableCell({
                borders: borders("DDDDDD"),
                shading: { fill: ri % 2 === 0 ? C.tableRow1 : C.tableRow2, type: ShadingType.CLEAR },
                width: { size: colWidths[ci], type: WidthType.DXA },
                margins: { top: 60, bottom: 60, left: 120, right: 120 },
                children: [new Paragraph({
                    children: [new TextRun({ text: cell, font: "Arial", size: 20, color: C.darkText })]
                })]
            }))
        })
    );
    return new Table({
        width: { size: colWidths.reduce((a, b) => a + b, 0), type: WidthType.DXA },
        columnWidths: colWidths,
        rows: [headerRow, ...dataRows]
    });
}

function bullet(text, level = 0, ref = "bullets") {
    return new Paragraph({
        numbering: { reference: ref, level },
        children: [new TextRun({ text, font: "Arial", size: 22, color: C.darkText })],
        spacing: { before: 60, after: 60 }
    });
}

function numbered(text, level = 0) {
    return bullet(text, level, "numbers");
}

function pageBreak() {
    return new Paragraph({ children: [new PageBreak()] });
}

function divider(color = C.mitRed) {
    return new Paragraph({
        border: { bottom: { style: BorderStyle.SINGLE, size: 4, color, space: 4 } },
        spacing: { before: 120, after: 120 }
    });
}

function quoteBox(quote, attribution) {
    return new Table({
        width: { size: 9360, type: WidthType.DXA },
        columnWidths: [9360],
        rows: [new TableRow({
            children: [new TableCell({
                borders: {
                    top: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
                    bottom: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
                    right: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
                    left: { style: BorderStyle.SINGLE, size: 20, color: C.warmGold },
                },
                shading: { fill: C.lightGold, type: ShadingType.CLEAR },
                width: { size: 9360, type: WidthType.DXA },
                margins: { top: 120, bottom: 120, left: 240, right: 160 },
                children: [
                    new Paragraph({
                        children: [new TextRun({ text: `"${quote}"`, font: "Arial", size: 22, italics: true, color: C.deepNavy })],
                        spacing: { before: 0, after: 80 }
                    }),
                    new Paragraph({
                        children: [new TextRun({ text: `— ${attribution}`, font: "Arial", size: 20, bold: true, color: C.softGold })],
                    })
                ]
            })]
        })]
    });
}

// ═══════════════════════════════════════════════════════════════════════════════
// DOCUMENT CONTENT
// ═══════════════════════════════════════════════════════════════════════════════

const doc = new Document({
    numbering: {
        config: [{
                reference: "bullets",
                levels: [{
                    level: 0,
                    format: LevelFormat.BULLET,
                    text: "•",
                    alignment: AlignmentType.LEFT,
                    style: { paragraph: { indent: { left: 720, hanging: 360 } } }
                }, {
                    level: 1,
                    format: LevelFormat.BULLET,
                    text: "◦",
                    alignment: AlignmentType.LEFT,
                    style: { paragraph: { indent: { left: 1080, hanging: 360 } } }
                }]
            },
            {
                reference: "numbers",
                levels: [{
                    level: 0,
                    format: LevelFormat.DECIMAL,
                    text: "%1.",
                    alignment: AlignmentType.LEFT,
                    style: { paragraph: { indent: { left: 720, hanging: 360 } } }
                }]
            }
        ]
    },
    styles: {
        default: { document: { run: { font: "Arial", size: 22, color: C.darkText } } },
        paragraphStyles: [{
                id: "Heading1",
                name: "Heading 1",
                basedOn: "Normal",
                next: "Normal",
                quickFormat: true,
                run: { size: 36, bold: true, font: "Arial", color: C.deepNavy },
                paragraph: { spacing: { before: 480, after: 200 }, outlineLevel: 0 }
            },
            {
                id: "Heading2",
                name: "Heading 2",
                basedOn: "Normal",
                next: "Normal",
                quickFormat: true,
                run: { size: 28, bold: true, font: "Arial", color: C.slateBlue },
                paragraph: { spacing: { before: 360, after: 160 }, outlineLevel: 1 }
            },
            {
                id: "Heading3",
                name: "Heading 3",
                basedOn: "Normal",
                next: "Normal",
                quickFormat: true,
                run: { size: 24, bold: true, font: "Arial", color: C.softGold },
                paragraph: { spacing: { before: 240, after: 120 }, outlineLevel: 2 }
            }
        ]
    },
    sections: [{
        properties: {
            page: {
                size: { width: 12240, height: 15840 },
                margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 }
            }
        },
        headers: {
            default: new Header({
                children: [
                    new Paragraph({
                        children: [
                            new TextRun({ text: "MIT MASTERCLASS  |  The Geometry of Meaning  |  Prof. E. Whitmore", font: "Arial", size: 18, color: C.mutedText }),
                        ],
                        border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: C.mitRed, space: 4 } },
                        alignment: AlignmentType.RIGHT
                    })
                ]
            })
        },
        footers: {
            default: new Footer({
                children: [
                    new Paragraph({
                        children: [
                            new TextRun({ text: "Confidential Academic Material  —  Page ", font: "Arial", size: 18, color: C.mutedText }),
                            new TextRun({ children: [PageNumber.CURRENT], font: "Arial", size: 18, color: C.mutedText }),
                            new TextRun({ text: " of ", font: "Arial", size: 18, color: C.mutedText }),
                            new TextRun({ children: [PageNumber.TOTAL_PAGES], font: "Arial", size: 18, color: C.mutedText }),
                        ],
                        border: { top: { style: BorderStyle.SINGLE, size: 4, color: C.mitRed, space: 4 } },
                        alignment: AlignmentType.CENTER
                    })
                ]
            })
        },
        children: [

            // ─── COVER ─────────────────────────────────────────────────────────────
            new Paragraph({
                children: [new TextRun({ text: "MASSACHUSETTS INSTITUTE OF TECHNOLOGY", font: "Arial", size: 20, bold: true, color: C.mitRed })],
                alignment: AlignmentType.CENTER,
                spacing: { before: 0, after: 80 }
            }),
            new Paragraph({
                children: [new TextRun({ text: "Department of Electrical Engineering & Computer Science", font: "Arial", size: 19, color: C.mutedText })],
                alignment: AlignmentType.CENTER,
                spacing: { before: 0, after: 400 }
            }),
            divider(C.mitRed),
            gap(200),
            new Paragraph({
                children: [new TextRun({ text: "MASTERCLASS LECTURE NOTES", font: "Arial", size: 22, bold: true, color: C.softGold, allCaps: true })],
                alignment: AlignmentType.CENTER,
                spacing: { before: 0, after: 120 }
            }),
            new Paragraph({
                children: [new TextRun({ text: "The Geometry of Meaning:", font: "Arial", size: 48, bold: true, color: C.deepNavy })],
                alignment: AlignmentType.CENTER,
                spacing: { before: 0, after: 60 }
            }),
            new Paragraph({
                children: [new TextRun({ text: "How AI Dynamically Warps Mathematics to Understand Language", font: "Arial", size: 32, bold: false, color: C.slateBlue })],
                alignment: AlignmentType.CENTER,
                spacing: { before: 0, after: 400 }
            }),
            divider(C.slateBlue),
            gap(160),
            new Paragraph({
                children: [new TextRun({ text: "Delivered by:", font: "Arial", size: 20, color: C.mutedText })],
                alignment: AlignmentType.CENTER,
                spacing: { before: 0, after: 60 }
            }),
            new Paragraph({
                children: [new TextRun({ text: "Professor Edmund R. Whitmore, Ph.D.", font: "Arial", size: 26, bold: true, color: C.deepNavy })],
                alignment: AlignmentType.CENTER,
                spacing: { before: 0, after: 40 }
            }),
            new Paragraph({
                children: [new TextRun({ text: "Vannevar Bush Professor of AI Systems  |  MIT CSAIL", font: "Arial", size: 20, color: C.mutedText })],
                alignment: AlignmentType.CENTER,
                spacing: { before: 0, after: 40 }
            }),
            new Paragraph({
                children: [new TextRun({ text: "Academic Year 2025 – 2026  |  Final Teaching Year Before Retirement", font: "Arial", size: 19, italic: true, color: C.mutedText })],
                alignment: AlignmentType.CENTER,
                spacing: { before: 0, after: 300 }
            }),
            gap(200),

            // ─── PROFESSOR'S OPENING LETTER ────────────────────────────────────────
            calloutBox(
                "A Letter from Professor Whitmore Before We Begin", [
                    "Fifty years ago, I was a 25-year-old engineer at Bell Labs staring at a 3-inch magnetic tape reel, wondering how we would ever teach a machine to understand the word 'run'. Today, I stand before what will be my last classroom at MIT, and the question has not become simpler — it has become gloriously more complex.",
                    "",
                    "The students in this room are not your average cohort. You want to understand the engine, not just drive the car. Some of your professors will tell you that year 2 is too early for this material. I am telling you the opposite: year 2 is exactly when the intuition is most elastic, before you get too rigid in your ways.",
                    "",
                    "What follows is not a tutorial. It is a lens. Use it to see into the architecture of systems that are reshaping civilisation. And yes — I know some of you are from Uganda, studying under constraints I had to actively imagine when I was building curriculum. That makes you sharper, not disadvantaged. Constraints breed innovation. That is not sentimentality. That is engineering history.",
                    "",
                    "Let us begin."
                ],
                C.lightGold, C.softGold
            ),
            gap(200),
            pageBreak(),

            // ─── TABLE OF CONTENTS ─────────────────────────────────────────────────
            h1("Table of Contents"),
            new TableOfContents("Table of Contents", {
                hyperlink: true,
                headingStyleRange: "1-3",
            }),
            pageBreak(),

            // ═══ SECTION 1 ════════════════════════════════════════════════════════
            h1("Section 1: The Story That Makes All of This Make Sense"),
            gap(60),

            h2("1.1  The Map Room Analogy"),
            body("Picture yourself in Kampala, at a large intelligence coordination room — the kind you see in war films. On the wall is a huge map of East Africa. Every town, river, border, and road has a coordinate: a pair of numbers (latitude, longitude). Masaka is not described in words. It is a dot at specific coordinates.", { justify: true }),
            body("Now the question is: if I told you 'Masaka,' how would a machine understand that Masaka is close to Kampala, but far from Mombasa — without the machine ever having read a geography book?", { justify: true }),
            body("The answer is that you give Masaka a position on a map. You let its neighbourhood tell the story.", { justify: true }),
            gap(80),

            quoteBox("In mathematics, position is identity. What a word is close to is more important than what it is defined as.", "Professor Edmund Whitmore, MIT CSAIL"),

            gap(120),
            body("This is the foundational insight behind word embeddings: every word in a language is assigned a coordinate in a multi-dimensional space. Words that tend to appear in similar contexts end up near each other. Words that never co-occur get pushed far apart.", { justify: true }),
            gap(80),

            h2("1.2  The Problem That Should Keep You Awake at Night"),
            body("Here is the puzzle that makes this field genuinely hard, and the reason self-attention was a Nobel-Prize-level idea in computer science:", { justify: true }),
            gap(60),
            calloutBox(
                "The Homonym Paradox", [
                    "Take the word 'BANK'.",
                    "",
                    "  Sentence A: 'I walked to the bank to deposit my salary.'",
                    "  Sentence B: 'We sat by the bank of the Nile and watched the sunset.'",
                    "",
                    "Both sentences are valid English. The word 'bank' is identical in spelling and in its static vector coordinate. But its meaning has completely shifted. One version belongs in the world of finance. The other belongs in the world of geography.",
                    "",
                    "How does a machine know which version it is reading? And once it knows — how does it dynamically shift the mathematical position of 'bank' to reflect that shift in meaning? That is the problem this masterclass solves."
                ],
                C.lightRed, C.mitRed
            ),
            gap(160),
            pageBreak(),

            // ═══ SECTION 2 ════════════════════════════════════════════════════════
            h1("Section 2: The Architecture — What Is a Vector Space?"),
            gap(60),

            h2("2.1  From Words to Numbers: Embeddings"),
            body("Before a Language Model can do anything, every word in its vocabulary must become a number — or more precisely, a list of numbers. This list is called a vector, and the system that produces it is called an embedding layer.", { justify: true }),
            gap(80),

            calloutBox("Core Definition", [
                "An EMBEDDING is a function that maps a word (a string) to a point in N-dimensional space (a vector of N floating-point numbers).",
                "",
                "For example, with N=3:",
                "   'banana'  →  [0.90, 0.15, 0.85]",
                "   'nano'    →  [0.00, 0.95, 0.10]",
                "   'fruit'   →  [0.95, 0.00, 0.25]",
                "",
                "Each of the 3 numbers represents a position along one conceptual axis.",
            ], C.lightBlue, C.slateBlue),

            gap(120),
            body("In the reference code provided, the embedding matrix is named E. It is a 5x3 matrix: 5 words, each with 3 coordinates. In real production systems like GPT-4 or Gemini, this matrix would be 100,000 words by 4,096 dimensions. The principle is identical. The scale is not.", { justify: true }),
            gap(80),

            h2("2.2  What the Axes Actually Mean"),
            body("This is where intuition lives. In the teaching example, the three axes are deliberately human-readable:", { justify: true }),
            gap(60),
            twoColTable(
                ["Axis", "Label", "What it measures"], [
                    ["X (axis 0)", "Food / Biology", "How food-related is this word?"],
                    ["Y (axis 1)", "Physics / Nanotechnology", "How tech-related is this word?"],
                    ["Z (axis 2)", "Yellow / Colour Intensity", "How colour-related (specifically yellow) is it?"],
                ], [2400, 3200, 3760]
            ),
            gap(120),
            body("In real models, the axes are not labelled by humans. They are learned automatically through training on massive datasets. An axis might capture 'formality vs informality', 'past vs future tense', or 'animal vs machine' — the model discovers these dimensions by exposure, not by instruction. This is one of the most profound things about deep learning: structure emerges from data.", { justify: true }),
            gap(80),

            h2("2.3  Cosine Similarity: How We Measure 'Closeness'"),
            body("Once words have coordinates, we need a way to measure how similar two words are. We do not use simple distance (Euclidean distance). We use angle — specifically, the cosine of the angle between two vectors.", { justify: true }),
            gap(80),
            calloutBox("The Formula", [
                "cosine_similarity(v1, v2) = (v1 · v2) / (||v1|| × ||v2||)",
                "",
                "Where · is the dot product and ||v|| is the magnitude (length) of the vector.",
                "",
                "Result is always between -1 and 1:",
                "   1.0  = vectors point in exactly the same direction (identical meaning)",
                "   0.0  = vectors are perpendicular (unrelated concepts)",
                "  -1.0  = vectors point in opposite directions (opposite concepts)",
            ], C.lightGreen, "2E7D32"),
            gap(120),
            body("This choice of cosine over Euclidean distance is not arbitrary. It makes the system robust to scale — a long vector and a short vector pointing in the same direction still score 1.0. What matters is direction, not magnitude. Meaning is directional.", { justify: true }),
            gap(80),
            pageBreak(),

            // ═══ SECTION 3 ════════════════════════════════════════════════════════
            h1("Section 3: Self-Attention — The Engine That Warps Reality"),
            gap(60),

            quoteBox("If embeddings are the map, then self-attention is the gravitational field that bends the map depending on what else is on it.", "Prof. Whitmore"),

            gap(120),
            h2("3.1  The Intuition Behind Attention"),
            body("Imagine you are reading a court case document. You encounter the sentence:", { justify: true }),
            gap(60),
            new Paragraph({
                children: [
                    new TextRun({ text: "        \"The judge ruled that the bank had to return the funds.\"", font: "Arial", size: 22, italics: true, color: C.deepNavy })
                ],
                spacing: { before: 80, after: 80 }
            }),
            gap(60),
            body("Your brain immediately understands that 'bank' here is financial. How? Because 'judge', 'ruled', and 'funds' are all in the same sentence — they are contextual neighbours that pull 'bank' toward the financial domain.", { justify: true }),
            body("Self-attention is the mathematical mechanism that does exactly this. It allows every word in a sentence to look at every other word and ask: 'How relevant are you to understanding me right now?' Then it updates its own position in the vector space based on the answers.", { justify: true }),
            gap(80),

            h2("3.2  The Three Matrices: Q, K, V — The Civil Service Model"),
            body("Self-attention uses three learnable transformation matrices: Queries (Q), Keys (K), and Values (V). Here is the most intuitive explanation I have ever constructed for this mechanism, and I have been teaching it for 15 years:", { justify: true }),
            gap(80),

            calloutBox("The Civil Registry Analogy", [
                "Imagine a government records office in Uganda's Ministry of Lands.",
                "",
                "QUERY (Q):  You walk in with a QUESTION: 'I need land ownership records for plot 45B, Nakawa.'",
                "KEY (K):    Each filing cabinet has an INDEX LABEL on it. Some say 'Nakawa plots', some say 'Kampala commercial', etc.",
                "VALUE (V):  Inside each cabinet is the ACTUAL DOCUMENT — the content itself.",
                "",
                "The attention mechanism works like this:",
                "1. Your Query is matched against all the Keys (index labels).",
                "2. The closer your Query matches a Key, the more weight you assign that cabinet.",
                "3. You pull out the Values (documents) from each cabinet in proportion to that weight.",
                "4. Your final understanding is a WEIGHTED MIXTURE of all the retrieved documents.",
                "",
                "The word 'bank' in a finance sentence sends out a Query. It finds high alignment with Keys for 'judge', 'funds', 'ruling'. It pulls their Values. Its final contextualised position shifts toward the finance region of the latent space."
            ], C.softOrange, C.warmGold),

            gap(120),
            h2("3.3  The Mathematics, Step by Step"),
            body("Let us now walk through the exact computation. This is the code from our reference implementation, explained line by line.", { justify: true }),
            gap(80),

            h3("Step 1 — Extract vectors for the words in the sentence"),
            codeBlock([
                "# sentence_agriculture = ['banana', 'yellow', 'fruit']",
                "token_indices = [vocab[token] for token in sentence_tokens]",
                "vectors = E[token_indices]  # Shape: (3 words, 3 dimensions)",
                "",
                "# For agriculture context, vectors becomes:",
                "# banana = [0.90, 0.15, 0.85]",
                "# yellow = [0.10, 0.00, 0.95]",
                "# fruit  = [0.95, 0.00, 0.25]",
            ]),
            gap(80),
            body("We are indexing into our embedding matrix E using the integer indices of our words. In NumPy, this is a clean matrix slice operation. The result is a 3x3 matrix — three words, each with three coordinates.", { justify: true }),
            gap(80),

            h3("Step 2 — Compute raw alignment scores (dot products)"),
            codeBlock([
                "Queries = vectors   # Each word asks: 'who is relevant to me?'",
                "Keys    = vectors   # Each word answers: 'here is what I am about'",
                "Values  = vectors   # Each word offers: 'here is my actual content'",
                "",
                "# Raw dot-product scores: how much does each word align with each other?",
                "scores = np.dot(Queries, Keys.T)",
                "# Result: a 3x3 matrix. scores[i][j] = how much word i attends to word j",
            ]),
            gap(80),
            body("The dot product is the mathematical backbone here. When two vectors point in similar directions, their dot product is large. When they are perpendicular, it is zero. This is how 'banana' and 'fruit' get a high alignment score — their vectors both have large values on the Food/Biology axis, so their dot product is high.", { justify: true }),
            gap(80),

            h3("Step 3 — Softmax: turning scores into clean probabilities"),
            codeBlock([
                "# Softmax converts raw scores into a probability distribution",
                "# The subtraction of max(scores) is a numerical stability trick",
                "exp_scores = np.exp(scores - np.max(scores, axis=-1, keepdims=True))",
                "attention_weights = exp_scores / np.sum(exp_scores, axis=-1, keepdims=True)",
                "",
                "# Now attention_weights[i] sums to 1.0",
                "# Think of it as: what percentage of my attention goes to each word?",
                "# Example: banana in agri context might attend: fruit=0.55, yellow=0.35, itself=0.10",
            ]),
            gap(80),
            body("Softmax is a critical operation. It takes arbitrary numbers and converts them into a valid probability distribution that sums to 1. The numerical stability trick (subtracting the max) prevents floating-point overflow when exponents get very large — this is standard industrial practice, not academic decoration.", { justify: true }),
            gap(80),

            h3("Step 4 — The update: computing the context-aware vector"),
            codeBlock([
                "# The new position of each word = weighted average of all Values",
                "updated_vectors = np.dot(attention_weights, Values)",
                "",
                "# banana_agri = updated_vectors[0]  # banana in agriculture context",
                "# This new vector has been PULLED toward yellow and fruit",
                "# Its coordinates are now a blend of banana + yellow + fruit",
            ]),
            gap(80),
            body("This one line is the entire update mechanism. The new position of 'banana' in the agricultural context is literally a weighted mixture of the original positions of banana, yellow, and fruit. The weights came from how much attention each word deserved. The result is that 'banana' has moved toward its contextual neighbours.", { justify: true }),
            gap(80),

            h3("Step 5 — The geometric proof: cosine similarity verification"),
            codeBlock([
                "# Compare how similar 'banana' is to 'nano' in two different contexts",
                "print(cosine_similarity(banana_agri, E[vocab['nano']]))  # Low — far from tech",
                "print(cosine_similarity(banana_tech, E[vocab['nano']]))  # High — close to tech",
                "",
                "# Expected output:",
                "# Similarity to 'nano' in AGRI context: ~0.28",
                "# Similarity to 'nano' in TECH context: ~0.89",
                "",
                "# The same word. Two radically different mathematical identities.",
            ]),
            gap(80),
            body("This is the mathematical proof that attention works. The word 'banana' — unchanged in spelling and in its original static embedding — has acquired two distinct mathematical identities depending on its conversational context. This is not magic. This is matrix multiplication producing geometry.", { justify: true }),
            gap(80),
            pageBreak(),

            // ═══ SECTION 4 ════════════════════════════════════════════════════════
            h1("Section 4: The Visualisation — Watching Meaning Move"),
            gap(60),

            h2("4.1  Understanding the 3D Plot"),
            body("The visualisation in the reference code produces a 3D scatter plot. Understanding what you are looking at is crucial:", { justify: true }),
            gap(80),
            twoColTable(
                ["Plot Element", "What It Represents", "Colour"], [
                    ["Black circle", "Static 'banana' — before any context", "Black"],
                    ["Green triangle", "'Banana' after agriculture context (banana, yellow, fruit)", "Green"],
                    ["Blue square", "'Banana' after technology context (banana, nano, tech)", "Blue"],
                    ["Dashed green line", "Trajectory of shift toward agricultural meaning", "Green"],
                    ["Dashed blue line", "Trajectory of shift toward technology meaning", "Blue"],
                ], [2800, 4200, 2360]
            ),
            gap(120),
            body("The key observation: the green triangle should be pulled high on the Z-axis (Yellow Colour) and the X-axis (Food/Biology). The blue square should be pulled high on the Y-axis (Nanotechnology). The static black dot sits in the middle — ambiguous, uncommitted, waiting for context.", { justify: true }),
            gap(80),

            h2("4.2  Why 3D Is a Lie (A Useful One)"),
            body("Here is a critical piece of intellectual honesty you must carry with you. Real language models operate in spaces of 768, 1024, or 4096 dimensions. We use 3D because human brains cannot visualise more. The 3D model is a teaching metaphor, not a faithful representation.", { justify: true }),
            body("In high dimensions, the geometry behaves strangely — what is known as the 'Curse of Dimensionality'. Almost all vectors become nearly equidistant from each other, and the intuitions from 3D space break down. Modern AI research is, in large part, the science of managing and exploiting this phenomenon. You will study this in your third or fourth year.", { justify: true }),
            gap(80),
            pageBreak(),

            // ═══ SECTION 5 ════════════════════════════════════════════════════════
            h1("Section 5: Real-World Applications — Where This Knowledge Pays"),
            gap(60),

            body("This is not academic theatre. Let me give you three scenarios at the level of a software engineer working in 2025, not a theoretician. These are inspired by real systems I have consulted on.", { justify: true }),
            gap(100),

            h2("5.1  Scenario: The Agricultural Bank of Uganda Chatbot"),
            body("Imagine you are hired as a software engineer at a financial technology company in Kampala. Their client is a rural agricultural lending institution. They want a chatbot that can process farmer loan applications submitted in natural language — sometimes in English, sometimes code-switching with Luganda words.", { justify: true }),
            gap(60),
            calloutBox("The Problem You Face", [
                "A foundational GPT-style model has never seen the word 'Omumbejja' (a Luganda term for a respectable woman farmer) in a financial context. Its embedding matrix has no coordinate for this word, or places it near irrelevant concepts.",
                "",
                "The model keeps misclassifying female applicants with high repayment histories as 'uncertain' because its latent space does not properly map local terminology to creditworthiness signals.",
                "",
                "Farmers in the Luwero district describe their land as 'bibanja land' — a form of customary land tenure. The model has never seen this term and cannot correctly evaluate it as collateral.",
            ], C.lightRed, C.mitRed),
            gap(80),
            body("Your solution — applying what you learned in this lecture — is domain-specific fine-tuning. You would:", { justify: true }),
            numbered("Collect a dataset of loan documents and agricultural reports from Uganda."),
            numbered("Freeze the base model weights and extend the embedding matrix with new coordinates for local terminology."),
            numbered("Fine-tune the Q, K, V weight matrices (W_Q, W_K, W_V) so that the attention mechanism learns to align 'Omumbejja' with concepts like 'reliable repayer', 'land holder', 'productive farmer'."),
            numbered("Validate by measuring cosine similarity between domain-specific financial terms and the newly trained embeddings."),
            gap(80),
            body("This is not a theoretical exercise. Companies like Jumo, Pezesha, and M-Kopa have built real versions of this architecture for Sub-Saharan African markets. The engineer who understands the vector geometry is the engineer who knows how to fix the system when it gets the meaning wrong.", { justify: true }),
            gap(100),

            h2("5.2  Scenario: Semantic Search for a Ugandan Legal Aid Platform"),
            body("A non-governmental organisation wants to build a system that allows people in Mbale to describe their legal situation in plain language and find relevant case law — without knowing legal terminology.", { justify: true }),
            gap(60),
            body("A citizen types: 'My landlord locked my door and kept my things when I was two weeks behind on rent. What can I do?'", { justify: true }),
            body("The system needs to match this description to cases involving unlawful distress, constructive eviction, and tenant rights — terms the citizen never used.", { justify: true }),
            gap(60),
            calloutBox("How Attention Solves This", [
                "Using RAG (Retrieval-Augmented Generation) built on the attention architecture:",
                "",
                "1. The citizen's sentence is converted into a vector using an embedding model.",
                "2. Every document in the case law database has also been converted into a vector.",
                "3. Cosine similarity is computed between the query vector and all document vectors.",
                "4. The top-ranked documents are retrieved — not by keyword matching, but by geometric proximity in meaning-space.",
                "5. An LLM generates a plain-language summary of the relevant rights.",
                "",
                "The system works because 'landlord locked my door' and 'unlawful distress' end up in the same neighbourhood of the legal latent space after training on legal corpora."
            ], C.lightBlue, C.slateBlue),

            gap(100),
            h2("5.3  Scenario: Green AI — Deploying Efficiently on Makerere's Servers"),
            body("A professor at Makerere University Institute of Computer Science wants to run a local AI assistant for students — but the university cannot afford continuous cloud API bills, and bandwidth is limited. They ask you to deploy a model on their own servers.", { justify: true }),
            gap(60),
            body("The challenge: the foundational model requires 80GB of GPU memory. Makerere has one 24GB GPU card.", { justify: true }),
            gap(60),
            body("Your knowledge of vector geometry gives you the tools to solve this:", { justify: true }),
            bullet("Quantization: Instead of storing each vector coordinate as a 32-bit float (precise but large), you convert to 4-bit integers. The vector [0.8943, 0.1251, 0.7734] becomes [7, 1, 6] in a 4-bit space. You lose some precision, but the geometry — the relative positions and angles — is largely preserved."),
            bullet("Knowledge Distillation: You train a smaller 'student' model to mimic the output distributions of the large 'teacher' model. The student's embedding space is a compressed but geometrically faithful projection of the teacher's."),
            bullet("Dimension Reduction: Using techniques like PCA (Principal Component Analysis) or UMAP, you project the high-dimensional vectors into fewer dimensions while preserving the most important structural relationships."),
            gap(80),
            body("The result: a model that fits in 18GB and runs on Makerere's hardware, with less than 3% degradation in performance on most Ugandan-language tasks. This is real engineering — not algorithm theory.", { justify: true }),
            gap(80),
            pageBreak(),

            // ═══ SECTION 6 ════════════════════════════════════════════════════════
            h1("Section 6: Who Studies This — and Where They End Up"),
            gap(60),

            h2("6.1  The Honest Landscape"),
            body("Let me be direct with you. This is the kind of material that separates engineers who build AI systems from engineers who use AI systems. Both are legitimate careers. But the path splits here.", { justify: true }),
            gap(80),

            twoColTable(
                ["Career Track", "Typical Educational Path", "Where They Land"], [
                    ["AI/ML Research Scientist", "BSc → MSc/PhD with focus on deep learning, linear algebra, optimisation theory", "DeepMind, OpenAI, Anthropic, Google Brain, academic labs, CERN-style national institutes"],
                    ["AI Systems Engineer", "BSc in CS/SE, strong in linear algebra and statistics, industry-focused projects", "Startups building AI products, enterprise AI teams at banks, telecoms, health tech"],
                    ["NLP Application Engineer", "BSc with specialisation in computational linguistics and NLP frameworks", "Legal tech, media companies, government AI divisions, translation platforms"],
                    ["MLOps / Infrastructure Engineer", "BSc in CS, specialisation in distributed systems and model deployment", "Cloud providers, fintech, any company running AI at scale in production"],
                ], [2800, 3200, 3360]
            ),
            gap(120),

            h2("6.2  Is Year 2 Too Early for This?"),
            gap(60),
            calloutBox(
                "Professor Whitmore's Honest Opinion", [
                    "No. Year 2 is not too early. But the question is the wrong question.",
                    "",
                    "Year 2 is the PERFECT time to build intuition, before you have had the creativity educated out of you by examination pressure.",
                    "",
                    "What year 2 students typically lack is not intelligence — it is mathematical fluency. If you are in year 2 and the chain rule makes you uncomfortable, or you have not internalised what a matrix multiplication really does geometrically, then this material will feel like memorisation rather than understanding.",
                    "",
                    "The correct question is: 'Do I have enough linear algebra to see WHY the matrix multiplication in self-attention produces a weighted average?' If yes — you are ready. If not — spend 3 weeks with Gilbert Strang's Linear Algebra lectures on MIT OpenCourseWare (they are free), then return.",
                    "",
                    "I have seen 20-year-olds from Nairobi and Lagos publish papers that embarrassed researchers with PhDs from Ivy League schools. Institutional calendar does not determine intellectual readiness.",
                ],
                C.lightGold, C.softGold
            ),

            gap(120),
            h2("6.3  The Prerequisite Map"),
            body("Here is what you need to be competent (not just curious) in this area:", { justify: true }),
            gap(80),
            twoColTable(
                ["Subject", "Why It Is Needed", "Best Free Resource"], [
                    ["Linear Algebra", "Vectors, matrices, dot products, eigenvectors — this IS the language of AI", "Gilbert Strang — MIT OCW 18.06"],
                    ["Calculus (Multivariable)", "Backpropagation trains all these weights through gradient descent", "3Blue1Brown — Essence of Calculus (YouTube)"],
                    ["Probability & Statistics", "Softmax, loss functions, cross-entropy — all probabilistic", "Khan Academy Statistics + Bishop's PRML book"],
                    ["Python & NumPy", "Every serious AI implementation uses this stack", "Fast.ai Practical Deep Learning (free)"],
                    ["Information Theory", "Why do we measure meaning with entropy? Shannon's framework", "'Elements of Information Theory' — Cover & Thomas"],
                ], [2400, 3400, 3560]
            ),
            gap(120),
            pageBreak(),

            // ═══ SECTION 7 ════════════════════════════════════════════════════════
            h1("Section 7: Professor Whitmore's Parting Advice"),
            gap(60),

            quoteBox("The best engineers I have ever worked with were not the ones who memorised the most algorithms. They were the ones who could look at a system and ask: 'What is this actually doing to the data?'", "Prof. Edmund Whitmore, MIT — 50 years in the field"),

            gap(120),
            h2("7.1  Ten Things I Wish Someone Had Told Me at 20"),
            gap(60),
            numbered("Master the fundamentals before you master the frameworks. TensorFlow and PyTorch will change. Matrix multiplication will not."),
            numbered("Build something broken, then fix it. Understanding comes from debugging, not from reading."),
            numbered("Your geographic location is not your limitation. The most transformative AI for Africa will be built by Africans who understand both the technology and the context. You are irreplaceable in that combination."),
            numbered("Publish early, publish imperfectly. A blog post explaining your project to a 15-year-old is worth more to your career than a perfect paper you never finish."),
            numbered("The best AI engineers I have worked with were obsessive readers — not just of computer science, but of history, economics, and anthropology. AI is a mirror of the data it sees. You must understand the world to understand what it will do."),
            numbered("Learn to say 'I don't know' without shame. It is the most powerful phrase in research."),
            numbered("Find one mentor who will disagree with you. Agreement is comfortable. Productive disagreement makes you rigorous."),
            numbered("The hardest part of this field is not the mathematics. It is deciding what problems are worth solving. That requires ethics, not algorithms."),
            numbered("Contribute to open-source. The communities around Hugging Face, EleutherAI, and LAION built the democratisation of AI. Your contribution matters."),
            numbered("Take care of your mental health. The field rewards obsession, but obsession without rest produces mediocre work and brilliant burnout. I have watched it destroy exceptional researchers. Do not let it destroy you."),
            gap(100),

            h2("7.2  What to Study Next (In Sequence)"),
            body("This lecture is a foundation stone, not a building. Here is the roadmap I give every student who wants to go deeper:", { justify: true }),
            gap(80),

            twoColTable(
                ["Order", "Topic", "Why Now"], [
                    ["1st", "Multi-Head Attention & Positional Encoding", "The full Transformer architecture — attention with multiple parallel heads and word-order awareness"],
                    ["2nd", "Backpropagation & Gradient Descent", "How are W_Q, W_K, W_V actually learned? You need to understand training"],
                    ["3rd", "The Transformer Architecture (Encoder-Decoder)", "BERT, GPT, T5 — all variations on this single unified design"],
                    ["4th", "Tokenisation & BPE (Byte Pair Encoding)", "Words are not the base unit — subword pieces are. Understanding this changes how you build systems"],
                    ["5th", "Fine-tuning & Parameter-Efficient Training (LoRA, PEFT)", "How to adapt large models to specific domains affordably"],
                    ["6th", "RAG — Retrieval-Augmented Generation", "The architecture of almost every production AI application today"],
                    ["7th", "Responsible AI — Bias in Embeddings", "Embedding spaces inherit the biases of their training data. Critical for deployment in African contexts"],
                ], [720, 3400, 5240]
            ),
            gap(120),
            pageBreak(),

            // ═══ SECTION 8 ════════════════════════════════════════════════════════
            h1("Section 8: Your Next Lecture — Follow-Up Prompt"),
            gap(60),

            body("Once you have read and understood this document — run the code, reproduce the cosine similarity numbers, modify the sentence contexts, and sketch the 3D geometry by hand — you are ready for the next lecture.", { justify: true }),
            body("Use the following prompt to generate the follow-up masterclass document:", { justify: true }),
            gap(100),

            calloutBox(
                "Follow-Up Prompt: Copy and Use This Exactly", [
                    "\"You are Professor Edmund Whitmore, the same 75-year-old MIT AI professor from the previous masterclass on The Geometry of Meaning. A student has read and understood the full lecture on embeddings, vector spaces, and self-attention. They have run the code, verified the cosine similarity outputs, and can now explain the dot-product attention formula in their own words.",
                    "",
                    "Now deliver the second lecture in this masterclass series:",
                    "",
                    "TOPIC: Multi-Head Attention, Positional Encoding, and the Full Transformer Architecture",
                    "",
                    "Requirements:",
                    "- Same format: a masterclass .docx with stories, analogies, and code walkthroughs",
                    "- Explain WHY we need multiple attention heads (the limitation of single-head attention)",
                    "- Explain positional encoding — how does the model know word ORDER if attention is order-blind?",
                    "- Walk through the full encoder block architecture: embedding → positional encoding → multi-head attention → feed-forward network → layer normalisation → residual connection",
                    "- Include a real-world scenario involving a Ugandan context",
                    "- Explain where BERT (encoder-only) and GPT (decoder-only) diverge in architecture",
                    "- Include a prerequisites check, career implications, and the next topic after this\"",
                ],
                C.lightGold, C.softGold
            ),

            gap(120),
            divider(C.mitRed),
            gap(80),
            new Paragraph({
                children: [new TextRun({ text: "End of Lecture Notes — The Geometry of Meaning", font: "Arial", size: 22, bold: true, italics: true, color: C.deepNavy })],
                alignment: AlignmentType.CENTER,
                spacing: { before: 80, after: 80 }
            }),
            new Paragraph({
                children: [new TextRun({ text: "Professor Edmund R. Whitmore  |  MIT CSAIL  |  Final Academic Year 2025–2026", font: "Arial", size: 19, color: C.mutedText })],
                alignment: AlignmentType.CENTER,
                spacing: { before: 0, after: 80 }
            }),
            new Paragraph({
                children: [new TextRun({ text: "\"The mathematics does not care about geography. Neither should you.\"", font: "Arial", size: 21, italics: true, color: C.softGold })],
                alignment: AlignmentType.CENTER,
                spacing: { before: 0, after: 0 }
            }),

        ]
    }]
});

Packer.toBuffer(doc).then(buffer => {
    const outDir = path.join(__dirname, "outputs");
    if (!fs.existsSync(outDir)) fs.mkdirSync(outDir, { recursive: true });
    const dest = path.join(outDir, "MIT_Masterclass_Geometry_of_Meaning.docx");
    fs.writeFileSync(dest, buffer);
    console.log("Done: MIT_Masterclass_Geometry_of_Meaning.docx written to:", dest);
});