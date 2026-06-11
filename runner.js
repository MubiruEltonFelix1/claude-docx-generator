const fs = require('fs');
const path = require('path');

const originalWriteFileSync = fs.writeFileSync.bind(fs);
const originalWriteFile = fs.writeFile.bind(fs);

const outDir = path.join(__dirname, 'outputs');
if (!fs.existsSync(outDir)) fs.mkdirSync(outDir, { recursive: true });

function makeDest(originalPath) {
    try {
        // If originalPath is a Buffer or not a string, fall back
        const p = String(originalPath || '');
        const base = path.basename(p) || `output_${Date.now()}.docx`;
        return path.join(outDir, base);
    } catch (e) {
        return path.join(outDir, `output_${Date.now()}.docx`);
    }
}

fs.writeFileSync = function(dest, data, ...args) {
    if (Buffer.isBuffer(data) || data instanceof Uint8Array) {
        const newDest = makeDest(dest);
        console.log('runner: Redirecting write to', newDest);
        return originalWriteFileSync(newDest, data, ...args);
    }
    return originalWriteFileSync(dest, data, ...args);
};

fs.writeFile = function(dest, data, ...args) {
    if (Buffer.isBuffer(data) || data instanceof Uint8Array) {
        const newDest = makeDest(dest);
        console.log('runner: Redirecting async write to', newDest);
        return originalWriteFile(newDest, data, ...args);
    }
    return originalWriteFile(dest, data, ...args);
};

// ── Auto-reset paste_here.js after each run ──────────────────
(function () {
    try {
        const mainFile = require.main && require.main.filename;
        const pasteHerePath = path.join(__dirname, 'paste_here.js');

        if (mainFile && path.resolve(mainFile) === pasteHerePath) {
            // We were loaded from paste_here.js — stash the clean template
            const pasteTemplate = [
                '// Paste your docx-generating JavaScript code into this file.',
                '// This file loads `runner.js` which redirects any Buffer writes',
                '// (such as the docx Buffer) into the local `outputs/` folder.',
                '',
                "require('./runner');",
                '',
                '// ----------------------- PASTE BELOW -----------------------',
                '// Example: paste the entire contents of a script like `generate2.js` here.',
                '// When you run `node paste_here.js`, any calls to `fs.writeFileSync(..., buffer)`',
                '// that write a Buffer will be redirected into the project\'s `outputs/` folder.',
                '',
                '// Paste your code starting on the next line:',
                '',
                '// ------------------------ END PASTE ------------------------',
                '',
            ].join('\n');

            process.on('exit', function () {
                try {
                    originalWriteFileSync(pasteHerePath, pasteTemplate, 'utf8');
                } catch (_) {
                    // If we can't restore, don't crash — it's just cleanup
                }
            });
        }
    } catch (_) {
        // Swallow any detection errors; they don't affect core functionality
    }
})();

module.exports = { outDir };