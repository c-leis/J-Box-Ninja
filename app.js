function normalize(text) {
    let t = text.toUpperCase();
    t = t.replace(/&AMP;/g, "&");

    t = t.replace(/CB\s*(\d+)([-A-Z]*)/g, "CIRCUITBRK$1$2");
    t = t.replace(/TRANSF(\d+)/g, "TRANSFO$1");

    return t.replace(/\s+/g, " ");
}

function findParts(text) {
    let parts = {};
    let qtyParts = new Set();

    text = text.replace(/\(\d{4}\)/g, "");
    text = text.replace(/\b\d+\s*[xX]\s*\d+\/0/g, "");
    text = text.replace(/\b\d+\s*[xX]\s*#?\d+\b/g, "");
    text = text.replace(/\b\d+\s*A\b/g, "");

    let match;

    // qty front
    let front = text.matchAll(/\((\d+)\)([A-Z0-9_\-/]+)/g);
    for (let m of front) {
        parts[m[2]] = (parts[m[2]] || 0) + parseInt(m[1]);
        qtyParts.add(m[2]);
    }

    // qty back
    let back = text.matchAll(/([A-Z0-9_\-/]+)\s*\((\d+)\)/g);
    for (let m of back) {
        parts[m[1]] = (parts[m[1]] || 0) + parseInt(m[2]);
        qtyParts.add(m[1]);
    }

    let tokens = text.match(/[A-Z0-9_\-/]+/g) || [];

    tokens.forEach(tok => {

        if (qtyParts.has(tok)) return;

        if (tok.startsWith("#")) tok = tok.replace("#","");

        if (/^\d+A$/.test(tok)) return;

        if (/^\d+\/0$/.test(tok)) {
            parts["WIRE_" + tok] = (parts["WIRE_" + tok] || 0) + 1;
            return;
        }

        if (/^\d+MCM$/.test(tok)) {
            parts["WIRE_" + tok] = (parts["WIRE_" + tok] || 0) + 1;
            return;
        }

        let wireMap = {
            "2":["WIRE100"],
            "4":["WIRE102"],
            "6":["WIRE101"],
            "8":["WIRE30"],
            "10":["WIRE35"],
            "12":["WIRE47"],
            "14":["WIRE29","WIRE44","WIRE45"]
        };

        if (wireMap[tok]) {
            wireMap[tok].forEach(w => {
                parts[w] = (parts[w] || 0) + 1;
            });
            return;
        }

        if (tok === "FUSE24") {
            parts[tok] = (parts[tok] || 0) + 2;
            return;
        }

        if (/^(J-BOX|PLATE|CIRCUITBRK|LUG|LOCK|ROTARY|HANDLE|COVER|TRANSFO|MECH|MOTOR|BLOCK|FUSE|SHAFT|BOX)/.test(tok)) {
            parts[tok] = (parts[tok] || 0) + 1;
        }
    });

    return parts;
}

function buildTable(parts) {
    let rows = [];

    for (let part in parts) {
        if (part.startsWith("WIRE")) {
            rows.push([part, 120, "", "FT", 5040]);
        } else {
            rows.push([part, 120, parts[part], "EA", 5040]);
        }
    }

    return rows;
}

function process() {

    const boxes = [
        box1.value,
        box2.value,
        box3.value,
        box4.value
    ];

    let wb = XLSX.utils.book_new();

    boxes.forEach((txt, i) => {

        if (!txt.trim()) return;

        let norm = normalize(txt);
        let parts = findParts(norm);
        let rows = buildTable(parts);

        let ws = {};

        // --------------------
        // HEADER ROW
        // --------------------
        const headers = ["PART","OPR","QTY","UOM","FIND"];

        headers.forEach((h, col) => {
            let cellRef = XLSX.utils.encode_cell({ r: 4, c: col + 1 });

            ws[cellRef] = {
                v: h,
                t: "s",
                s: {
                    fill: { fgColor: { rgb: "800080" } },
                    font: { color: { rgb: "FFFFFF" }, bold: true },
                    alignment: { horizontal: "center" }
                }
            };
        });

        // --------------------
        // DATA ROWS
        // --------------------
        rows.forEach((row, r) => {
            row.forEach((val, c) => {

                let cellRef = XLSX.utils.encode_cell({
                    r: r + 5,
                    c: c + 1
                });

                ws[cellRef] = {
                    v: val,
                    t: typeof val === "number" ? "n" : "s",
                    s: {
                        alignment: (c >= 1 && c <= 4)
                            ? { horizontal: "center" }
                            : {}
                    }
                };
            });
        });

        // --------------------
        // COLUMN WIDTHS
        // --------------------
        ws["!cols"] = [
            {}, // A (empty spacer)
            { wch: 23 },
            { wch: 9 },
            { wch: 9 },
            { wch: 9 },
            { wch: 9 },
            { wch: 9 },
            { wch: 9 }
        ];

        // --------------------
        // MERGES (WARNING BOX)
        // --------------------
        ws["!merges"] = [
            // Warning block
            {
                s: { r: 4, c: 15 },
                e: { r: 9, c: 18 }
            },
            // Info block
            {
                s: { r: 13, c: 15 },
                e: { r: 16, c: 18 }
            }
        ];

        // --------------------
        // WARNING CELL
        // --------------------
        ws["P5"] = {
            v: "VERIFY WIRE QTY & UOM",
            t: "s",
            s: {
                fill: { fgColor: { rgb: "FF0000" } },
                font: { bold: true },
                alignment: {
                    horizontal: "center",
                    vertical: "center",
                    wrapText: true
                }
            }
        };

        // --------------------
        // INFO CELL
        // --------------------
        ws["P14"] = {
            v: "Called out Wires are highlighted in the Wire Options table. Verify quantities and UOM.",
            t: "s",
            s: {
                fill: { fgColor: { rgb: "FFFF00" } },
                alignment: {
                    wrapText: true,
                    horizontal: "center",
                    vertical: "center"
                }
            }
        };

        // --------------------
        // RANGE
        // --------------------
        ws["!ref"] = XLSX.utils.encode_range({
            s: { r: 0, c: 0 },
            e: { r: rows.length + 20, c: 20 }
        });

        XLSX.utils.book_append_sheet(wb, ws, "JBOX" + (i+1));
    });

    XLSX.writeFile(wb, "jbox_parts.xlsx");
}
