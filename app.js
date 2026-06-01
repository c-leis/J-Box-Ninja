function normalize(text) {
    let t = text.toUpperCase();
    t = t.replace(/&AMP;/g, "&");

    t = t.replace(/CB\s*(\d+)([-A-Z]*)/g, "CIRCUITBRK$1$2");
    t = t.replace(/TRANSF(\d+)/g, "TRANSFO$1");

    return t.replace(/\s+/g, " ");
}

function breakerSize(part) {
    let match = part.match(/CIRCUITBRK(\d+)/);
    return match ? parseInt(match[1]) : 0;
}

function partPriority(p) {
    if (p.startsWith("J-BOX")) return 0;
    if (p.startsWith("PLATE")) return 1;

    if (/^(CIRCUITBRK|LUG|LOCK|ROTARY|HANDLE|COVER)/.test(p)) return 2;
    if (/^(TRANSFO|MECH|MOTOR)/.test(p)) return 3;
    if (p.startsWith("BLOCK")) return 4;
    if (/^(SHAFT|FUSE)/.test(p)) return 5;
    if (/^(GFCI-OUTLET|BOX|PLATE25)/.test(p)) return 6;
    if (/^(GROUNDBAR|BRACKET|TERMINA)/.test(p)) return 7;
    if (p.startsWith("WIRE")) return 8;

    return 9;
}

function findParts(text) {
    let parts = {};
    let qtyParts = new Set();

    text = text.replace(/\(\d{4}\)/g, "");
    text = text.replace(/\b\d+\s*[xX]\s*\d+\/0/g, "");
    text = text.replace(/\b\d+\s*[xX]\s*#?\d+\b/g, "");
    text = text.replace(/\b\d+\s*A\b/g, "");

    for (let m of text.matchAll(/\((\d+)\)([A-Z0-9_\-/]+)/g)) {
        parts[m[2]] = (parts[m[2]] || 0) + parseInt(m[1]);
        qtyParts.add(m[2]);
    }

    for (let m of text.matchAll(/([A-Z0-9_\-/]+)\s*\((\d+)\)/g)) {
        parts[m[1]] = (parts[m[1]] || 0) + parseInt(m[2]);
        qtyParts.add(m[1]);
    }

    let tokens = text.match(/[A-Z0-9_\-/]+/g) || [];

    tokens.forEach(tok => {

        if (qtyParts.has(tok)) return;

        if (tok.startsWith("#")) tok = tok.replace("#", "");
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

        if (/^(J-BOX|PLATE|CIRCUITBRK|LUG|LOCK|ROTARY|HANDLE|COVER|TRANSFO|MECH|MOTOR|BLOCK|FUSE|SHAFT|BOX|GFCI-OUTLET|PLATE25|GROUNDBAR|BRACKET|TERMINA)/.test(tok)) {
            parts[tok] = (parts[tok] || 0) + 1;
        }
    });

    return parts;
}

function buildTable(parts) {
    let rows = [];
    let usedWires = new Set();

    for (let part in parts) {

        let row;

        if (part.startsWith("WIRE")) {
            usedWires.add(part);
            row = { part, opr:120, qty:"", uom:"FT", find:5040 };
        } else {
            row = { part, opr:120, qty:parts[part], uom:"EA", find:5040 };
        }

        row.priority = partPriority(part);
        row.breaker = breakerSize(part);

        rows.push(row);
    }

    // ✅ SORT (matches Python)
    rows.sort((a, b) => {
        if (a.priority !== b.priority) return a.priority - b.priority;
        if (a.breaker !== b.breaker) return b.breaker - a.breaker;
        return a.part.localeCompare(b.part);
    });

    return { rows, usedWires };
}

function addWireTable(ws, usedWires) {

    const wires = [
        ["WIRE_500MCM",""],["WIRE_400MCM",""],["WIRE_350MCM",""],["WIRE_300MCM",""],["WIRE_250MCM",""],
        ["WIRE_4/0","4/0"],["WIRE_3/0","3/0"],["WIRE_2/0","2/0"],["WIRE_1/0","1/0"],
        ["WIRE103","#1"],["WIRE116","#1 Green"],["WIRE100","#2"],["WIRE107","#2 Green"],
        ["WIRE102","#4"],["WIRE109","#4 Green"],["WIRE101","#6"],["WIRE110","#6 Green"],
        ["WIRE30","#8"],["WIRE39","#8 Green"],["WIRE35","#10"],["WIRE36","#10 Green"],
        ["WIRE47","#12"],["WIRE50","#12 Green"],["WIRE106","#12 White"],
        ["WIRE29","#14 Black"],["WIRE44","#14 White"],["WIRE45","#14 Green"]
    ];

    ws.mergeCells("I5:N5");

    let header = ws.getCell("I5");
    header.value = "WIRE OPTIONS";
    header.fill = { type:"pattern", pattern:"solid", fgColor:{ argb:"FFFF00" } };
    header.alignment = { horizontal:"center" };
    header.font = { bold:true };

    const headers = ["PART","OPR","QTY","UOM","FIND","SIZE"];

    headers.forEach((h,i) => {
        ws.getRow(6).getCell(i+9).value = h;
    });

    wires.forEach((w,i) => {
        let row = ws.getRow(7+i);

        row.getCell(9).value = w[0];
        row.getCell(10).value = 120;
        row.getCell(12).value = "FT";
        row.getCell(13).value = 5040;
        row.getCell(14).value = w[1];

        if (usedWires.has(w[0])) {
            for (let c=9;c<=14;c++) {
                row.getCell(c).fill = {
                    type:"pattern",
                    pattern:"solid",
                    fgColor:{ argb:"00FF00" }
                };
            }
        }
    });
}

async function process() {

    try {

        const boxes = [
            document.getElementById("box1").value,
            document.getElementById("box2").value,
            document.getElementById("box3").value,
            document.getElementById("box4").value
        ];

        const wb = new ExcelJS.Workbook();

        for (let i=0;i<boxes.length;i++) {

            let txt = boxes[i];
            if (!txt.trim()) continue;

            let { rows, usedWires } = buildTable(findParts(normalize(txt)));

            let ws = wb.addWorksheet("JBOX"+(i+1));

            ws.columns = [
                {}, { width:23 }, { width:9 }, { width:9 }, { width:9 }, { width:9 }
            ];

            // HEADER
            let headers = ["PART","OPR","QTY","UOM","FIND"];
            let headerRow = ws.getRow(5);

            headers.forEach((h,i) => {
                let cell = headerRow.getCell(i+2);
                cell.value = h;
                cell.fill = { type:"pattern", pattern:"solid", fgColor:{argb:"800080"} };
                cell.font = { color:{argb:"FFFFFF"}, bold:true };
                cell.alignment = { horizontal:"center" };
            });

            // DATA
            rows.forEach((r,idx) => {
                let row = ws.getRow(6+idx);

                row.getCell(2).value = r.part;
                row.getCell(3).value = r.opr;
                row.getCell(4).value = r.qty;
                row.getCell(5).value = r.uom;
                row.getCell(6).value = r.find;

                [3,4,5,6].forEach(c=>{
                    row.getCell(c).alignment = { horizontal:"center" };
                });
            });

            // WARNING
            ws.mergeCells("P5:S10");
            ws.getCell("P5").value = "VERIFY WIRE QTY & UOM";
            ws.getCell("P5").fill = {
                type: "pattern", 
                pattern: "solid", 
                fgColor: { argb: "FF0000" }
            };

            ws.getCell("P5").alignment = {
                horizontal: "center",
                vertical: "middle",
                wrapText: true
            };

            // INFO
            ws.mergeCells("P14:S17");
            ws.getCell("P14").value =
                "Highlighted wires shown in table.\nVerify qty + UOM.";
            ws.getCell("P14").fill = {
                type: "pattern", 
                pattern: "solid", 
                fgColor: { argb: "FFFF00" }
            };

            ws.getCell("P14").alignment = {
                horizontal: "center",
                vertical: "middle",
                wrapText: true
            };

            addWireTable(ws, usedWires);
        }

        const buffer = await wb.xlsx.writeBuffer();

        const blob = new Blob([buffer], {
            type:"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        });

        const link = document.createElement("a");
        link.href = URL.createObjectURL(blob);
        link.download = "jbox_parts.xlsx";
        link.click();

    } catch (err) {
        console.error(err);
        alert("Error: " + err.message);
    }
}
