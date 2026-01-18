import { VisioPackage } from '../src/VisioPackage';
import * as path from 'path';
import * as fs from 'fs';

async function run() {
    const vsdxPath = path.resolve(__dirname, 'simple-schema.vsdx');
    if (!fs.existsSync(vsdxPath)) {
        console.error("File not found:", vsdxPath);
        return;
    }

    console.log("Inspecting:", vsdxPath);
    // VisioPackage needs to handle distinct reading.
    // Actually VisioPackage ctor creates new empty or expects template.
    // We need to load from existing buffer.

    // Let's rely on standard zip reading if we can, or hack VisioPackage to open existing.
    // VisioPackage uses adm-zip internally.

    // Let's try to just use adm-zip directly if installed, or just use my logic.
    // Looking at VisioPackage.ts... I don't see a static 'load' or 'open' method that takes a path to *edit*.
    // But I can create a new VisioPackage and manually load the data?
    // No, VisioPackage constructor initializes a NEW zip or loads a template.

    // Actually, looking at `VisioDocument.create()` it calls `new VisioPackage()`.
    // I need a way to READ the generated VSDX.

    // Let's just use `unzip` command line tool to cat the file?
    // Using `run_command` with `unzip -p examples/simple-schema.vsdx visio/pages/page1.xml` is easiest.
    console.log("Use the command line to inspect.");
}

run();
