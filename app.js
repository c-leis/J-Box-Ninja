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

async function process() {
    alert("Button clicked");
    const boxes = [
        box1.value,
        box2.value,
        box3.value,
        box4.value
    ];

    const wb = new ExcelJS.Workbook();

    for (let i = 0; i < boxes.length; i++) {

        let txt = boxes[i];
        if (!txt.trim()) continue;

        let norm = normalize(txt);
        let parts = findParts(norm);
        let rows = buildTable(parts);

        let ws = wb.addWorksheet("JBOX" + (i+1));

        // Column widths (match your Python)
        ws.columns = [
            {}, // spacer
            { width: 23 },
            { width: 9 },
            { width: 9 },
            { width: 9 },
            { width: 9 }
        ];

        // Header row (Row 5)
        const headers = ["PART","OPR","QTY","UOM","FIND"];
        let headerRow = ws.getRow(5);

        headers.forEach((h, i) => {
            let cell = headerRow.getCell(i+2);
            cell.value = h;
            cell.fill = {
                type: "pattern",
                pattern: "solid",
                fgColor: { argb: "800080" }
            };
            cell.font = { color: { argb: "FFFFFF" }, bold: true };
            cell.alignment = { horizontal: "center" };
        });

        // Data rows
        rows.forEach((row, r) => {
            let excelRow = ws.getRow(r + 6);

            row.forEach((val, c) => {
                let cell = excelRow.getCell(c + 2);
                cell.value = val;

                if (c >= 1 && c <= 4) {
                    cell.alignment = { horizontal: "center" };
                }
            });
        });

        // WARNING BLOCK (same as Python)
        ws.mergeCells("P5:S10");
        let warningCell = ws.getCell("P5");

        warningCell.value = "⚠ VERIFY WIRE QTY & UOM ⚠";
        warningCell.fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "FF0000" }
        };
        warningCell.alignment = {
            horizontal: "center",
            vertical: "middle",
            wrapText: true
        };
        warningCell.font = { bold: true };

        // INFO BLOCK
        ws.mergeCells("P14:S17");
        let infoCell = ws.getCell("P14");

        infoCell.value =
            "Called out wires are highlighted in the wire table.\nVerify quantities and UOM.";
        infoCell.fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "FFFF00" }
        };
        infoCell.alignment = {
            horizontal: "center",
            vertical: "middle",
            wrapText: true
        };
    }

    const buffer = await wb.xlsx.writeBuffer();

    const blob = new Blob([buffer], {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    });

    const link = document.createElement("a");
    link.href = URL.createObjectURL(blob);
    link.download = "jbox_parts.xlsx";
    link.click();
}
}
