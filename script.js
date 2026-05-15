// ✅ THEME
function toggleTheme(){
    document.body.classList.toggle("light");

    const text = document.getElementById("themeText");

    text.innerText = document.body.classList.contains("light")
        ? "☀ Light"
        : "🌙 Dark";
}

document.addEventListener("DOMContentLoaded", () => {

    const drop = document.getElementById("drop");
    const input = document.getElementById("fileInput");
    const status = document.getElementById("status");
    const downloadBtn = document.getElementById("downloadBtn");
    const themeToggle = document.getElementById("themeToggle");

    let parsedData = null;
    let originalFileName = "output";

    let headers = [];
    let weightedIndex = -1;
    let scoreIndexes = [];
    let identIndexes = [];

    // ✅ THEME EVENT
    themeToggle.addEventListener("change", toggleTheme);

    // ✅ DEFAULT DARK
    document.body.classList.remove("light");

    // ✅ CLICK
    drop.onclick = () => input.click();

    input.onchange = e => {
        if (e.target.files.length) loadFile(e.target.files[0]);
    };

    // ✅ DRAG
    drop.ondragover = e => {
        e.preventDefault();
        drop.classList.add("drag");
    };

    drop.ondragleave = () => drop.classList.remove("drag");

    drop.ondrop = e => {
        e.preventDefault();
        drop.classList.remove("drag");

        if (e.dataTransfer.files.length) {
            loadFile(e.dataTransfer.files[0]);
        }
    };

    // ✅ CSV PARSER
    function parseCSVLine(line) {
        let result = [];
        let current = "";
        let inQuotes = false;

        for (let i = 0; i < line.length; i++) {
            let c = line[i];

            if (c === '"') inQuotes = !inQuotes;
            else if (c === ',' && !inQuotes) {
                result.push(current);
                current = "";
            } else current += c;
        }

        result.push(current);
        return result;
    }

    function loadFile(file){

        originalFileName = file.name.replace(/\.[^/.]+$/, "");
        status.innerText = "Dosya okunuyor...";

        const reader = new FileReader();

        reader.onload = e => {
            let content = e.target.result;

            if (file.name.endsWith(".xlsx") || file.name.endsWith(".xls")) {
                const wb = XLSX.read(content, { type: "array" });
                const ws = wb.Sheets[wb.SheetNames[0]];
                content = XLSX.utils.sheet_to_csv(ws);
            }

            const rows = content.split(/\r?\n/).filter(r => r.trim());

            parsedData = rows.map(r => parseCSVLine(r));

            // ✅ HEADER ANALIZ
            headers = parsedData[0] || [];

            weightedIndex = headers.findIndex(h =>
                h.toLowerCase().includes("weighted")
            );

            scoreIndexes = headers
                .map((h,i)=>h.toLowerCase().includes("score") ? i : -1)
                .filter(i=>i!==-1);

            identIndexes = headers
                .map((h,i)=>h.toLowerCase().includes("ident") ? i : -1)
                .filter(i=>i!==-1);

            status.innerText = "✅ Yüklendi (" + parsedData.length + " satır)";
        };

        if (file.name.endsWith(".csv"))
            reader.readAsText(file,"UTF-8");
        else
            reader.readAsArrayBuffer(file);
    }

    // ✅ SMART FORMAT (EN KRİTİK)
    function formatCell(value, colIndex){

        if (!value) return "";

        let v = String(value).replace(/"/g, "").trim();

        // ✅ IDENT → dokunma
        if (identIndexes.includes(colIndex)) {
            return v;
        }

        // ✅ WEIGHTED AVG → TR string
        if (colIndex === weightedIndex) {
            if (/^\d+\.\d+$/.test(v)) {
                return v.replace(".", ",");
            }
        }

        // ✅ SCORE → TR string
        if (scoreIndexes.includes(colIndex)) {
            if (/^\d+\.\d+$/.test(v)) {
                return v.replace(".", ",");
            }
        }

        // ✅ DATE
        if (/^\d{4}-\d{2}-\d{2}/.test(v)) {
            let d = new Date(v);
            if (!isNaN(d)) {
                return { v: d, t: 'd' };
            }
        }

        // ✅ NORMAL NUMBER
        if (/^\d+\.\d+$/.test(v)) {
            return { v: parseFloat(v), t: 'n' };
        }

        if (/^\d+$/.test(v)) {
            return { v: parseInt(v), t: 'n' };
        }

        return v;
    }

    // ✅ EXPORT
    downloadBtn.onclick = () => {

        if (!parsedData) return alert("Önce dosya yükle");

        const ws = XLSX.utils.aoa_to_sheet(
            parsedData.map((row, rIndex) =>
                row.map((cell, cIndex) =>
                    rIndex === 0 ? cell : formatCell(cell, cIndex)
                )
            )
        );

        // ✅ FORMAT UYGULA
        Object.keys(ws).forEach(cell => {
            if (cell[0] === '!') return;

            let c = ws[cell];

            if (c.t === 'n') c.z = '#,##0.00';
            if (c.t === 'd') c.z = 'dd.mm.yyyy';
        });

        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, "Data");

        XLSX.writeFile(wb, originalFileName + "_clean.xlsx");

        status.innerText = "✅ Excel indirildi";
    };

    // ✅ BACK BUTTON
    document.getElementById("backBtnPlain").addEventListener("click", () => {
        window.location.href = "index.html";
    });
});
