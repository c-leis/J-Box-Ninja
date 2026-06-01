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

    let boxes = [
        document.getElementById("box1").value,
        document.getElementById("box2").value,
        document.getElementById("box3").value,
        document.getElementById("box4").value
    ];

    let wb = XLSX.utils.book_new();
    let outputText = "";

    boxes.forEach((txt, i) => {

        if (txt.trim()) {

            let norm = normalize(txt);
            let parts = findParts(norm);
            let rows = buildTable(parts);

            let ws = XLSX.utils.aoa_to_sheet([
                ["PART","OPR","QTY","UOM","FIND"],
                ...rows
            ]);

            XLSX.utils.book_append_sheet(wb, ws, "JBOX" + (i+1));

            outputText += `JBOX ${i+1}:\n` + JSON.stringify(parts, null, 2) + "\n\n";
        }
    });

    document.getElementById("output").innerText = outputText;

    XLSX.writeFile(wb, "jbox_parts.xlsx");
}
``