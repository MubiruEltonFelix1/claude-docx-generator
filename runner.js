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

module.exports = { outDir };