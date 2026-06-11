#!/usr/bin/env node
const path = require('path');
const fs = require('fs');

const targetFile = process.argv[2];

if (!targetFile) {
    console.error(
        'Usage: claude-docx-gen <path-to-script.js>\n' +
        '\n' +
        '  Takes a JavaScript file that builds a docx Document and writes it via\n' +
        '  fs.writeFileSync(buffer), intercepts the write, and saves the .docx into\n' +
        '  an outputs/ directory in the current working directory.\n'
    );
    process.exit(1);
}

const scriptPath = path.resolve(targetFile);

if (!fs.existsSync(scriptPath)) {
    console.error('Error: file not found — ' + scriptPath);
    process.exit(1);
}

// Set up the write-intercepting runner first
require('./runner');

// Then execute the target script
require(scriptPath);
