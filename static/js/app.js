/**
 * Frontend Application Logic untuk Traffic Counting ITDP v2.3
 * Pemetaan Lengkap 15 Kolom Kendaraan Sesuai Excel:
 * - 9 Kolom Nonlistrik (C s/d K): MCY, CAR, LGV, MGV, HGV, MIKROBUS, LBUS, MBUS, HBUS
 * - 6 Kolom Listrik EV (L s/d Q): MCY, CAR, MIKROBUS, LBUS, MBUS, HBUS (Aktif 06.00–07.00)
 * Fitur v2.3:
 * - Tombol -5 dan -1 untuk koreksi cepat
 * - Input langsung angka (bisa diketik manual)
 * - Hotkey Shift + 1..6 untuk EV (hardware keycode mapping)
 */

const AppState = {
    hasFile: false,
    currentFilepath: null,
    templateData: null,
    activeRow: 11,
    currentViewMode: "counter", // 'counter' | 'excel'
    packageFilter: "ALL",       // ALL, PG, SR
    kelompokFilter: "ALL",      // ALL, K1, K2
    countsData: {},
    noteText: "",
    soundEnabled: true
};

document.addEventListener("DOMContentLoaded", () => {
    initApp();
});

async function initApp() {
    setupDragAndDrop();
    setupEventListeners();
    setupKeyboardShortcuts();
    await loadQuickPickFiles();
}

// 1. Setup Drag & Drop Slot
function setupDragAndDrop() {
    const dropArea = document.getElementById("dragDropArea");
    const fileInput = document.getElementById("directFileInput");

    if (!dropArea || !fileInput) return;

    window.addEventListener("dragover", (e) => e.preventDefault(), false);
    window.addEventListener("drop", (e) => e.preventDefault(), false);

    let dragCounter = 0;

    dropArea.addEventListener("dragenter", (e) => {
        e.preventDefault();
        e.stopPropagation();
        dragCounter++;
        dropArea.classList.add("border-sky-400", "bg-sky-950/60", "ring-4", "ring-sky-500/30", "scale-[1.02]");
        const titleEl = dropArea.querySelector("h2");
        if (titleEl) titleEl.textContent = "Lepaskan Berkas Excel di Sini Sekarang!";
    }, false);

    dropArea.addEventListener("dragover", (e) => {
        e.preventDefault();
        e.stopPropagation();
    }, false);

    dropArea.addEventListener("dragleave", (e) => {
        e.preventDefault();
        e.stopPropagation();
        dragCounter--;
        if (dragCounter <= 0) {
            dragCounter = 0;
            dropArea.classList.remove("border-sky-400", "bg-sky-950/60", "ring-4", "ring-sky-500/30", "scale-[1.02]");
            const titleEl = dropArea.querySelector("h2");
            if (titleEl) titleEl.textContent = "Tarik & Lepaskan File Excel (.xlsx) di Sini";
        }
    }, false);

    dropArea.addEventListener("drop", async (e) => {
        e.preventDefault();
        e.stopPropagation();
        dragCounter = 0;
        dropArea.classList.remove("border-sky-400", "bg-sky-950/60", "ring-4", "ring-sky-500/30", "scale-[1.02]");
        const titleEl = dropArea.querySelector("h2");
        if (titleEl) titleEl.textContent = "Tarik & Lepaskan File Excel (.xlsx) di Sini";

        const dt = e.dataTransfer;
        if (dt && dt.files && dt.files.length > 0) {
            await handleFileUpload(dt.files[0]);
        }
    }, false);

    dropArea.addEventListener("click", (e) => {
        if (e.target.tagName !== "INPUT") {
            fileInput.click();
        }
    });

    fileInput.addEventListener("change", async (e) => {
        if (e.target.files && e.target.files.length > 0) {
            await handleFileUpload(e.target.files[0]);
        }
    });
}

// 2. Setup Event Listeners
function setupEventListeners() {
    const soundToggle = document.getElementById("soundToggle");
    if (soundToggle) {
        soundToggle.addEventListener("click", () => {
            AppState.soundEnabled = !AppState.soundEnabled;
            window.soundFx.enabled = AppState.soundEnabled;
            soundToggle.innerHTML = AppState.soundEnabled 
                ? '<i class="fas fa-volume-up text-emerald-400"></i> Suara Aktif' 
                : '<i class="fas fa-volume-mute text-slate-400"></i> Suara Nonaktif';
        });
    }

    const noteInput = document.getElementById("catatanInput");
    if (noteInput) {
        noteInput.addEventListener("input", (e) => {
            AppState.noteText = e.target.value;
        });
    }
}

// 3. Quick Pick List Files
async function loadQuickPickFiles() {
    try {
        const res = await fetch("/api/files");
        const json = await res.json();
        const listContainer = document.getElementById("quickPickFilesList");
        if (listContainer && json.files) {
            listContainer.innerHTML = "";
            if (json.files.length === 0) {
                document.getElementById("quickPickContainer").classList.add("hidden");
                return;
            }
            document.getElementById("quickPickContainer").classList.remove("hidden");
            json.files.forEach(f => {
                const btn = document.createElement("button");
                btn.type = "button";
                btn.className = "flex items-center gap-1.5 px-3 py-1.5 rounded-lg bg-slate-800 hover:bg-sky-900/50 border border-slate-700 hover:border-sky-500 text-slate-300 hover:text-sky-300 transition-all font-mono text-xs cursor-pointer";
                btn.innerHTML = `<i class="fas fa-file-excel text-emerald-400"></i> <span class="truncate max-w-[240px]">${f.filename}</span>`;
                btn.addEventListener("click", (e) => {
                    e.stopPropagation();
                    loadExistingFile(f.filepath);
                });
                listContainer.appendChild(btn);
            });
        }
    } catch (e) {
        console.error("Gagal load quick pick files:", e);
    }
}

// 4. File Upload Handler
async function handleFileUpload(file) {
    if (!file || !file.name) return;

    if (!file.name.toLowerCase().endsWith(".xlsx")) {
        showToast("File harus berekstensi .xlsx", "error");
        return;
    }

    const formData = new FormData();
    formData.append("file", file);

    try {
        showToast("Mengunggah dan membaca template Excel...", "info");
        const res = await fetch("/api/upload", {
            method: "POST",
            body: formData
        });
        const json = await res.json();
        if (json.success) {
            onFileReady(json.data);
            showToast(json.message, "success");
        } else {
            showToast(json.message || "Gagal memproses berkas.", "error");
        }
    } catch (err) {
        showToast("Gagal mengunggah berkas: " + err, "error");
    }
}

// 5. Load Existing File Handler
async function loadExistingFile(filepath) {
    try {
        showToast("Membuka template Excel...", "info");
        const res = await fetch("/api/load-file", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ filepath })
        });
        const json = await res.json();
        if (json.success) {
            onFileReady(json.data);
            showToast("Template berhasil dibuka!", "success");
        } else {
            showToast(json.message, "error");
        }
    } catch (e) {
        showToast("Gagal memuat template: " + e, "error");
    }
}

// 6. On File Ready
function onFileReady(data) {
    AppState.hasFile = true;
    AppState.templateData = data;
    AppState.currentFilepath = data.filepath;
    AppState.countsData = data.counts_data || {};
    AppState.noteText = data.header_info.catatan || "";

    const h = data.header_info;

    // A. Aturan Paket: Otomatis terkunci dari judul file
    AppState.packageFilter = h.package || "ALL";

    // B. Default Kelompok: TAMPILKAN SEMUA KATEGORI (ALL) kecuali bila nama file terkunci K1 atau K2
    if (h.kelompok && (h.kelompok === "K1" || h.kelompok === "K2") && h.kelompok_locked) {
        AppState.kelompokFilter = h.kelompok;
    } else {
        AppState.kelompokFilter = "ALL"; // SEMUA KATEGORI C s/d Q DITAMPILKAN LENGKAP
    }

    if (AppState.packageFilter === "SR") {
        AppState.activeRow = 51;
    } else {
        AppState.activeRow = 11;
    }

    document.getElementById("dropZoneSlotSection").classList.add("hidden");
    document.getElementById("activeWorkspace").classList.remove("hidden");
    document.getElementById("activeWorkspace").classList.add("flex");
    
    document.getElementById("viewModeSwitcher").classList.remove("hidden");
    document.getElementById("viewModeSwitcher").classList.add("flex");
    document.getElementById("activeFilePill").classList.remove("hidden");
    document.getElementById("activeFilePill").classList.add("flex");
    document.getElementById("fileActionButtons").classList.remove("hidden");
    document.getElementById("fileActionButtons").classList.add("flex");

    document.getElementById("activeFileName").textContent = data.filename;
    document.getElementById("activeFileName").title = data.filename;

    if (data.existing_stats && data.existing_stats.filled_sessions > 0) {
        const banner = document.getElementById("existingDataBanner");
        const statsEl = document.getElementById("existingDataStats");
        banner.classList.remove("hidden");
        banner.classList.add("flex");
        statsEl.textContent = `File ini sudah memuat ${data.existing_stats.total_vehicles} kendaraan (${data.existing_stats.filled_sessions} sesi terisi). Data langsung terbaca dan Anda dapat melanjutkan tanpa mulai dari 0.`;
    } else {
        document.getElementById("existingDataBanner").classList.add("hidden");
    }

    updateHeaderUI();
    renderSessionTimeline();
    renderCounterDeck();
    renderExcelSpreadsheetTable();
}

// 7. Close / Ganti File
async function closeActiveFile() {
    if (!confirm("Apakah Anda ingin mengganti file? Pastikan data terakhir sudah disimpan ke Excel.")) return;
    try {
        await fetch("/api/close-file", { method: "POST" });
        AppState.hasFile = false;
        AppState.currentFilepath = null;
        AppState.templateData = null;
        AppState.countsData = {};

        document.getElementById("dropZoneSlotSection").classList.remove("hidden");
        document.getElementById("activeWorkspace").classList.add("hidden");
        document.getElementById("activeWorkspace").classList.remove("flex");
        document.getElementById("viewModeSwitcher").classList.add("hidden");
        document.getElementById("viewModeSwitcher").classList.remove("flex");
        document.getElementById("activeFilePill").classList.add("hidden");
        document.getElementById("activeFilePill").classList.remove("flex");
        document.getElementById("fileActionButtons").classList.add("hidden");
        document.getElementById("fileActionButtons").classList.remove("flex");
        document.getElementById("existingDataBanner").classList.add("hidden");

        await loadQuickPickFiles();
        showToast("File ditutup. Silakan drag & drop file lain.", "info");
    } catch (e) {
        console.error(e);
    }
}

// 8. Update Header UI
function updateHeaderUI() {
    const data = AppState.templateData;
    if (!data) return;
    const h = data.header_info;

    document.getElementById("headerTitik").textContent = h.titik || "-";
    document.getElementById("headerRuas").textContent = h.ruas || "-";
    document.getElementById("headerRuas").title = h.ruas || "";
    document.getElementById("headerArah").textContent = h.arah || "-";
    document.getElementById("headerHari").textContent = h.hari_tgl || "-";
    document.getElementById("headerHari").title = h.hari_tgl || "";
    document.getElementById("headerPencacah").textContent = h.pencacah || "-";
    document.getElementById("headerDariKe").textContent = `${h.dari} ➔ ${h.ke}`;
    document.getElementById("headerTipe").textContent = h.tipe || "-";
    document.getElementById("catatanInput").value = AppState.noteText || "";

    document.getElementById("packageBadgeText").textContent = h.package_label || h.package;

    const kelText = document.getElementById("kelompokBadgeText");
    const kelSelect = document.getElementById("kelompokSelect");

    if (h.is_8_lanes) {
        if (h.kelompok_locked) {
            kelText.textContent = `${h.kelompok} (${h.kelompok === 'K1' ? 'Motor & Barang' : 'Mobil & Bus'}) · Ruas 8/2 D`;
            kelText.classList.remove("hidden");
            kelSelect.classList.add("hidden");
        } else {
            // Ruas 8/2 D tapi belum terkunci di nama file: tampilkan dropdown dengan default ALL
            kelText.classList.add("hidden");
            kelSelect.classList.remove("hidden");
            kelSelect.value = AppState.kelompokFilter;
        }
    } else {
        kelText.textContent = `Semua Kategori (Ruas ${h.tipe})`;
        kelText.classList.remove("hidden");
        kelSelect.classList.add("hidden");
    }
}

function changeKelompok(val) {
    AppState.kelompokFilter = val;
    renderCounterDeck();
}

// 9. Switch View Mode ('counter' | 'excel')
function switchViewMode(mode) {
    AppState.currentViewMode = mode;
    const btnCounter = document.getElementById("btnModeCounter");
    const btnExcel = document.getElementById("btnModeExcel");
    const counterContainer = document.getElementById("counterModeContainer");
    const excelContainer = document.getElementById("excelViewModeContainer");

    if (mode === "counter") {
        btnCounter.className = "px-3 py-1.5 rounded-lg text-xs font-bold transition-all bg-sky-500 text-white shadow-md";
        btnExcel.className = "px-3 py-1.5 rounded-lg text-xs font-bold transition-all text-slate-400 hover:text-white";
        counterContainer.classList.remove("hidden");
        excelContainer.classList.add("hidden");
    } else {
        btnExcel.className = "px-3 py-1.5 rounded-lg text-xs font-bold transition-all bg-emerald-500 text-white shadow-md";
        btnCounter.className = "px-3 py-1.5 rounded-lg text-xs font-bold transition-all text-slate-400 hover:text-white";
        excelContainer.classList.remove("hidden");
        counterContainer.classList.add("hidden");
        renderExcelSpreadsheetTable();
    }
}

// 10. Filter Intervals Sesuai Paket
function getFilteredIntervals() {
    if (!AppState.templateData) return [];
    const all = AppState.templateData.intervals;
    if (AppState.packageFilter === "PG") {
        return all.filter(i => 11 <= i.row && i.row <= 26);
    } else if (AppState.packageFilter === "SR") {
        return all.filter(i => 51 <= i.row && i.row <= 66);
    } else {
        return all;
    }
}

// 11. Render 15-Minute Session Timeline ("Anti-Numpuk")
function renderSessionTimeline() {
    const container = document.getElementById("sessionTimelineContainer");
    if (!container) return;
    container.innerHTML = "";

    const intervals = getFilteredIntervals();
    intervals.forEach((item) => {
        const sess = AppState.countsData[String(item.row)] || { status: "empty" };
        const isActive = item.row === AppState.activeRow;

        const chip = document.createElement("button");
        chip.type = "button";
        chip.id = `session-chip-${item.row}`;
        chip.className = `session-chip flex-shrink-0 flex items-center gap-2 px-3.5 py-2 rounded-xl text-xs font-semibold border transition-all ${
            isActive 
                ? "active bg-sky-500 text-white border-sky-400 shadow-lg shadow-sky-500/25 ring-2 ring-sky-300"
                : item.color === "blue" 
                    ? "bg-slate-800/90 text-slate-200 border-slate-700/80 hover:bg-slate-750" 
                    : "bg-amber-950/40 text-amber-200 border-amber-800/40 hover:bg-amber-900/40"
        }`;

        let statusBadge = '<span class="w-2 h-2 rounded-full bg-slate-500"></span>';
        if (sess.status === "verified") {
            statusBadge = '<i class="fas fa-check-circle text-emerald-400 text-xs"></i>';
        } else if (sess.status === "draft") {
            statusBadge = '<span class="w-2 h-2 rounded-full bg-amber-400 animate-pulse"></span>';
        }

        chip.innerHTML = `
            ${statusBadge}
            <span class="font-mono-numbers font-bold">${item.time_label}</span>
            <span class="text-[10px] px-1.5 py-0.5 rounded ${
                item.color === "blue" ? "bg-sky-950/60 text-sky-300" : "bg-yellow-950/60 text-yellow-300"
            }">B.${item.row}</span>
        `;

        chip.addEventListener("click", () => {
            setActiveSession(item.row);
        });

        container.appendChild(chip);
    });

    setTimeout(() => {
        const activeChip = document.getElementById(`session-chip-${AppState.activeRow}`);
        if (activeChip) {
            activeChip.scrollIntoView({ behavior: "smooth", inline: "center", block: "nearest" });
        }
    }, 50);

    updateSessionCounterSummary();
}

// 12. Set Active Session
function setActiveSession(row) {
    AppState.activeRow = row;
    renderSessionTimeline();
    renderCounterDeck();
}

// 13. Render Counter Deck (9 Non-EV & 6 EV Lengkap)
function renderCounterDeck() {
    const data = AppState.templateData;
    if (!data) return;

    const row = AppState.activeRow;
    const intervalInfo = data.intervals.find(i => i.row === row) || {
        time_label: "06.00-06.15",
        period: "Peak Pagi",
        color: "blue",
        is_ev_active: (11 <= row && row <= 14)
    };

    const isEvActive = (11 <= row && row <= 14);

    document.getElementById("activeSessionTime").textContent = intervalInfo.time_label;
    document.getElementById("activeSessionRow").textContent = `Baris ${row}`;
    document.getElementById("activeSessionPeriod").textContent = intervalInfo.period;
    
    const evTag = document.getElementById("activeSessionEvTag");
    if (isEvActive) {
        evTag.className = "px-2 py-0.5 rounded-md text-[11px] font-bold bg-emerald-950 text-emerald-400 border border-emerald-800";
        evTag.textContent = "⚡ EV Aktif (Kolom L–Q)";
    } else {
        evTag.className = "px-2 py-0.5 rounded-md text-[11px] font-bold bg-slate-800 text-slate-400 border border-slate-700";
        evTag.textContent = "EV Digabung ke Konvensional";
    }

    const bannerEl = document.getElementById("activeSessionBanner");
    if (intervalInfo.color === "blue") {
        bannerEl.className = "flex flex-wrap items-center justify-between p-4 rounded-2xl bg-gradient-to-r from-sky-900/60 via-slate-850 to-slate-900 border border-sky-700/40 shadow-xl";
    } else {
        bannerEl.className = "flex flex-wrap items-center justify-between p-4 rounded-2xl bg-gradient-to-r from-amber-950/60 via-slate-850 to-slate-900 border border-amber-700/40 shadow-xl";
    }

    if (!AppState.countsData[String(row)]) {
        AppState.countsData[String(row)] = { counts: {}, status: "empty" };
    }
    const sessCounts = AppState.countsData[String(row)].counts;

    // A. Render Non-EV Deck (Kolom C s/d K · 9 Kolom Lengkap)
    const nonEvContainer = document.getElementById("nonEvCardsContainer");
    nonEvContainer.innerHTML = "";

    data.categories_non_ev.forEach(cat => {
        // Tampilkan semua kecuali jika user secara spesifik memilih K1 atau K2
        if (AppState.kelompokFilter === "K1" && cat.k_type !== "K1") return;
        if (AppState.kelompokFilter === "K2" && cat.k_type !== "K2") return;

        const val = sessCounts[cat.code] !== undefined && sessCounts[cat.code] !== null 
            ? sessCounts[cat.code] 
            : 0;

        const card = createCounterCard({
            cat: cat,
            val: val,
            isEv: false,
            badgeColor: "bg-sky-500/20 text-sky-300 border-sky-500/40"
        });
        nonEvContainer.appendChild(card);
    });

    // B. Render EV Deck (Kolom L s/d Q · 6 Kolom Lengkap)
    const evContainer = document.getElementById("evCardsContainer");
    const evNotice = document.getElementById("evNoticeDisabled");

    if (isEvActive) {
        evNotice.classList.add("hidden");
        evContainer.classList.remove("hidden");
        evContainer.innerHTML = "";

        data.categories_ev.forEach(cat => {
            // Tampilkan semua kecuali jika user secara spesifik memilih K1 atau K2
            if (AppState.kelompokFilter === "K1" && cat.k_type !== "K1") return;
            if (AppState.kelompokFilter === "K2" && cat.k_type !== "K2") return;

            const val = sessCounts[cat.code] !== undefined && sessCounts[cat.code] !== null 
                ? sessCounts[cat.code] 
                : 0;

            const card = createCounterCard({
                cat: cat,
                val: val,
                isEv: true,
                badgeColor: "bg-emerald-500/20 text-emerald-300 border-emerald-500/40"
            });
            evContainer.appendChild(card);
        });
    } else {
        // Luar jam 06.00-07.00 dinonaktifkan
        evNotice.classList.remove("hidden");
        evContainer.classList.add("hidden");
    }

    updateSessionCounterSummary();
}

// 14. Create Counter Card Component (Dengan Label Kolom Jelas & Input Manual)
function createCounterCard({ cat, val, isEv, badgeColor }) {
    const card = document.createElement("div");
    card.id = `card-${cat.code}`;
    card.className = "counter-card bg-slate-850/90 rounded-2xl p-4 border border-slate-750 flex flex-col justify-between shadow-lg hover:shadow-xl transition-all relative overflow-hidden";

    let iconClass = "fa-car";
    if (cat.code.includes("MCY")) iconClass = "fa-motorcycle";
    else if (cat.code.includes("LGV")) iconClass = "fa-truck-pickup";
    else if (cat.code.includes("MGV") || cat.code.includes("HGV")) iconClass = "fa-truck";
    else if (cat.code.includes("BUS")) iconClass = "fa-bus";

    const displayTitle = cat.display_code || cat.code;

    card.innerHTML = `
        <div class="flex items-start justify-between gap-2 mb-3">
            <div>
                <div class="flex items-center gap-2">
                    <span class="inline-flex items-center gap-1.5 px-2.5 py-0.5 rounded-full text-xs font-bold border ${badgeColor}">
                        <i class="fas ${iconClass} text-[11px]"></i>
                        ${displayTitle}
                    </span>
                    <span class="text-[11px] font-mono font-bold text-sky-400 bg-sky-950/80 px-2 py-0.5 rounded border border-sky-800">Kolom ${cat.col_letter}</span>
                    <span class="text-[10px] text-slate-400 font-mono bg-slate-800 px-1.5 py-0.5 rounded border border-slate-700">Gol. ${cat.group}</span>
                </div>
                <h4 class="text-xs font-medium text-slate-300 mt-1 line-clamp-1" title="${cat.name}">${cat.name}</h4>
            </div>
            <span class="px-2 py-0.5 rounded text-[11px] font-mono font-bold bg-slate-800 text-sky-400 border border-slate-700 shadow-inner">
                ${cat.hotkey}
            </span>
        </div>

        <!-- Kontrol Input & Tombol Cepat: -5, -1, Input Langsung, +1, +5 -->
        <div class="my-2 flex items-center justify-between gap-1.5 bg-slate-900/90 rounded-xl p-2.5 border border-slate-800">
            <!-- Tombol Pengurang: -5 dan -1 -->
            <div class="flex items-center gap-1">
                <button type="button" onclick="decrementTally('${cat.code}', 5)" class="h-9 px-2 rounded-lg bg-slate-800 hover:bg-rose-900/40 text-slate-300 hover:text-rose-300 active:scale-90 transition-all border border-slate-700 text-xs font-mono font-bold cursor-pointer" title="Kurangi 5 (-5)">
                    -5
                </button>
                <button type="button" onclick="decrementTally('${cat.code}', 1)" class="w-9 h-9 flex items-center justify-center rounded-lg bg-slate-800 hover:bg-rose-900/40 text-slate-300 hover:text-rose-300 active:scale-90 transition-all border border-slate-700 text-xs font-bold cursor-pointer" title="Kurangi 1 (-1)">
                    <i class="fas fa-minus"></i>
                </button>
            </div>

            <!-- Input Nilai Langsung (Bisa Diketik Manual & Live Sync) -->
            <div class="flex flex-col items-center flex-1 mx-1">
                <input type="number" 
                       id="counter-val-${cat.code}" 
                       value="${val}" 
                       min="0"
                       step="1"
                       oninput="setManualTally('${cat.code}', this.value)"
                       onchange="if(this.value === '' || parseInt(this.value, 10) < 0) { this.value = 0; setManualTally('${cat.code}', 0); }"
                       class="w-full max-w-[84px] text-center text-2xl sm:text-3xl font-extrabold font-mono text-white bg-slate-950/80 rounded-lg border border-slate-700 focus:border-sky-400 focus:ring-2 focus:ring-sky-500/30 focus:outline-none py-0.5 px-1 tracking-tight"
                       title="Ketik manual angka langsung di sini bila ada revisi" />
                <span class="text-[10px] text-slate-400 font-medium">kendaraan</span>
            </div>

            <!-- Tombol Penambah: +1 dan +5 -->
            <div class="flex items-center gap-1">
                <button type="button" onclick="incrementTally('${cat.code}', 1)" class="w-9 h-9 flex items-center justify-center rounded-lg bg-slate-800 hover:bg-sky-900/40 text-slate-300 hover:text-sky-300 active:scale-90 transition-all border border-slate-700 text-xs font-mono font-bold cursor-pointer" title="Tambah 1 (+1)">
                    +1
                </button>
                <button type="button" onclick="incrementTally('${cat.code}', 5)" class="h-9 px-2 rounded-lg bg-slate-800 hover:bg-sky-900/40 text-slate-300 hover:text-sky-300 active:scale-90 transition-all border border-slate-700 text-xs font-mono font-bold cursor-pointer" title="Tambah 5 (+5)">
                    +5
                </button>
            </div>
        </div>

        <!-- Tombol Utama Pencacahan (+1) -->
        <button type="button" onclick="incrementTally('${cat.code}', 1)" class="tally-btn-active w-full py-3.5 mt-2 rounded-xl bg-gradient-to-r from-sky-600 to-blue-600 hover:from-sky-500 hover:to-blue-500 text-white font-bold text-sm sm:text-base shadow-lg shadow-sky-600/30 flex items-center justify-center gap-2 border border-sky-400/30 transition-all cursor-pointer">
            <i class="fas fa-plus text-sm"></i>
            HITUNG +1
        </button>
    `;

    return card;
}

// 15. Tally Actions & Manual Edit
function setManualTally(code, rawVal) {
    const row = AppState.activeRow;
    const isEv = code.endsWith("_EV");
    if (isEv && !(11 <= row && row <= 14)) {
        showToast("Kendaraan listrik hanya aktif di jam 06.00-07.00!", "error");
        return;
    }

    if (!AppState.countsData[String(row)]) {
        AppState.countsData[String(row)] = { counts: {}, status: "draft" };
    }
    const sess = AppState.countsData[String(row)];
    
    let parsed = parseInt(rawVal, 10);
    if (isNaN(parsed) || parsed < 0) {
        parsed = 0;
    }

    sess.counts[code] = parsed;
    sess.status = "draft";

    animateCounterBump(code);
    updateSessionCounterSummary();
}

function incrementTally(code, delta = 1) {
    const row = AppState.activeRow;
    const isEv = code.endsWith("_EV");
    if (isEv && !(11 <= row && row <= 14)) {
        showToast("Kendaraan listrik hanya aktif di jam 06.00-07.00!", "error");
        return;
    }

    if (!AppState.countsData[String(row)]) {
        AppState.countsData[String(row)] = { counts: {}, status: "draft" };
    }
    const sess = AppState.countsData[String(row)];
    const curVal = sess.counts[code] || 0;
    const newVal = Math.max(0, curVal + delta);
    sess.counts[code] = newVal;
    sess.status = "draft";

    window.soundFx.playClick();
    animateCounterBump(code);

    const el = document.getElementById(`counter-val-${code}`);
    if (el) {
        if (el.tagName === "INPUT") el.value = newVal;
        else el.textContent = newVal;
    }

    updateSessionCounterSummary();
}

function decrementTally(code, delta = 1) {
    const row = AppState.activeRow;
    if (!AppState.countsData[String(row)]) return;
    const sess = AppState.countsData[String(row)];
    const curVal = sess.counts[code] || 0;
    if (curVal <= 0) return;

    const newVal = Math.max(0, curVal - delta);
    sess.counts[code] = newVal;
    sess.status = "draft";

    window.soundFx.playDecrement();
    animateCounterBump(code);

    const el = document.getElementById(`counter-val-${code}`);
    if (el) {
        if (el.tagName === "INPUT") el.value = newVal;
        else el.textContent = newVal;
    }

    updateSessionCounterSummary();
}

function animateCounterBump(code) {
    const el = document.getElementById(`counter-val-${code}`);
    if (el) {
        el.classList.remove("bump-animate");
        void el.offsetWidth;
        el.classList.add("bump-animate");
    }
}

// 16. Summary Calculator
function updateSessionCounterSummary() {
    const row = AppState.activeRow;
    const sess = AppState.countsData[String(row)] || { counts: {} };
    let nonEvTotal = 0;
    let evTotal = 0;

    if (AppState.templateData) {
        AppState.templateData.categories_non_ev.forEach(cat => {
            nonEvTotal += (sess.counts[cat.code] || 0);
        });
        if (11 <= row && row <= 14) {
            AppState.templateData.categories_ev.forEach(cat => {
                evTotal += (sess.counts[cat.code] || 0);
            });
        }
    }

    const totalEl = document.getElementById("sessionTotalVehicles");
    if (totalEl) totalEl.textContent = nonEvTotal + evTotal;

    const nonEvEl = document.getElementById("sessionNonEvTotal");
    if (nonEvEl) nonEvEl.textContent = nonEvTotal;

    const evEl = document.getElementById("sessionEvTotal");
    if (evEl) evEl.textContent = evTotal;

    let totalFilled = 0;
    const intervals = getFilteredIntervals();
    intervals.forEach(i => {
        const s = AppState.countsData[String(i.row)];
        if (s && (s.status === "verified" || s.total > 0)) totalFilled++;
    });

    const progressEl = document.getElementById("overallProgressText");
    if (progressEl) progressEl.textContent = `${totalFilled} / ${intervals.length} Sesi Terverifikasi`;
    
    const progressBar = document.getElementById("overallProgressBar");
    if (progressBar && intervals.length > 0) {
        const pct = Math.round((totalFilled / intervals.length) * 100);
        progressBar.style.width = `${pct}%`;
    }
}

// 17. Sesi Verifikasi & Navigasi
function verifyCurrentSession() {
    const row = AppState.activeRow;
    if (!AppState.countsData[String(row)]) {
        AppState.countsData[String(row)] = { counts: {}, status: "verified" };
    }
    const sess = AppState.countsData[String(row)];

    AppState.templateData.categories_non_ev.forEach(cat => {
        if (sess.counts[cat.code] === undefined || sess.counts[cat.code] === null) {
            sess.counts[cat.code] = 0;
        }
    });

    if (11 <= row && row <= 14) {
        AppState.templateData.categories_ev.forEach(cat => {
            if (sess.counts[cat.code] === undefined || sess.counts[cat.code] === null) {
                sess.counts[cat.code] = 0;
            }
        });
    }

    sess.status = "verified";
    window.soundFx.playSuccess();

    showToast(`Sesi Baris ${row} terverifikasi!`, "success");
    goToNextSession();
}

function fillZerosCurrentSession() {
    const row = AppState.activeRow;
    if (!AppState.countsData[String(row)]) {
        AppState.countsData[String(row)] = { counts: {}, status: "draft" };
    }
    const sess = AppState.countsData[String(row)];

    AppState.templateData.categories_non_ev.forEach(cat => {
        if (!sess.counts[cat.code]) sess.counts[cat.code] = 0;
    });

    if (11 <= row && row <= 14) {
        AppState.templateData.categories_ev.forEach(cat => {
            if (!sess.counts[cat.code]) sess.counts[cat.code] = 0;
        });
    }

    renderCounterDeck();
    showToast("Semua kategori kosong diisi 0.", "info");
}

function resetCurrentSession() {
    if (!confirm(`Yakin ingin mereset semua hitungan di sesi Baris ${AppState.activeRow}?`)) return;
    const row = AppState.activeRow;
    AppState.countsData[String(row)] = { counts: {}, status: "empty" };
    renderCounterDeck();
    renderSessionTimeline();
    showToast("Sesi direset ke 0.", "info");
}

function goToPrevSession() {
    const intervals = getFilteredIntervals();
    const curIdx = intervals.findIndex(i => i.row === AppState.activeRow);
    if (curIdx > 0) {
        setActiveSession(intervals[curIdx - 1].row);
    } else {
        showToast("Ini adalah sesi pertama.", "info");
    }
}

function goToNextSession() {
    const intervals = getFilteredIntervals();
    const curIdx = intervals.findIndex(i => i.row === AppState.activeRow);
    if (curIdx >= 0 && curIdx < intervals.length - 1) {
        setActiveSession(intervals[curIdx + 1].row);
    } else {
        showToast("Sudah mencapai sesi terakhir paket!", "success");
    }
}

// 18. Keyboard Shortcuts (Mendukung Shift + 1..6 untuk EV & 1..9 untuk Non-EV)
function getDigitFromEvent(e) {
    // A. Deteksi fisik tombol (hardware independent - bekerja di semua OS & layout)
    if (e.code) {
        const m = e.code.match(/^(?:Digit|Numpad)([1-9])$/);
        if (m) return parseInt(m[1], 10);
    }
    // B. Deteksi e.key numerik biasa
    if (/^[1-9]$/.test(e.key)) {
        return parseInt(e.key, 10);
    }
    // C. Pemetaan simbol Shift pada keyboard QWERTY standar (!, @, #, $, %, ^, &, *, ()
    const shiftSymbols = {
        '!': 1,
        '@': 2,
        '#': 3,
        '$': 4,
        '%': 5,
        '^': 6,
        '&': 7,
        '*': 8,
        '(': 9
    };
    if (shiftSymbols[e.key]) {
        return shiftSymbols[e.key];
    }
    return null;
}

function setupKeyboardShortcuts() {
    window.addEventListener("keydown", (e) => {
        // Jangan picu hotkey bila sedang mengetik angka langsung di input atau text field
        if (e.target.tagName === "INPUT" || e.target.tagName === "TEXTAREA") {
            if (e.key === "Enter") {
                e.target.blur(); // Enter untuk selesai input
            }
            return;
        }
        if (!AppState.hasFile) return;

        // Navigasi & Verifikasi Sesi
        if ((e.ctrlKey && e.key === "Enter") || e.code === "Space") {
            e.preventDefault();
            verifyCurrentSession();
            return;
        }

        if (e.key === "ArrowLeft") {
            e.preventDefault();
            goToPrevSession();
            return;
        }
        if (e.key === "ArrowRight") {
            e.preventDefault();
            goToNextSession();
            return;
        }

        const isShift = e.shiftKey;
        const digit = getDigitFromEvent(e);
        if (!digit) return;

        if (isShift) {
            // Hotkey Kendaraan Listrik / EV: Shift + 1 s/d 6
            if (11 <= AppState.activeRow && AppState.activeRow <= 14) {
                const evMap = {
                    1: "MCY_EV",      // Kolom L
                    2: "CAR_EV",      // Kolom M
                    3: "MIKROBUS_EV", // Kolom N
                    4: "LBUS_EV",     // Kolom O
                    5: "MBUS_EV",     // Kolom P
                    6: "HBUS_EV"      // Kolom Q
                };
                const targetCode = evMap[digit];
                if (targetCode) {
                    // Hormati filter kelompok K1 / K2 bila aktif
                    if (AppState.templateData && AppState.templateData.categories_ev) {
                        const catObj = AppState.templateData.categories_ev.find(c => c.code === targetCode);
                        if (catObj) {
                            if (AppState.kelompokFilter === "K1" && catObj.k_type !== "K1") return;
                            if (AppState.kelompokFilter === "K2" && catObj.k_type !== "K2") return;
                        }
                    }
                    e.preventDefault();
                    incrementTally(targetCode, 1);
                    return;
                }
            } else {
                if (digit >= 1 && digit <= 6) {
                    e.preventDefault();
                    showToast("Kendaraan listrik dinonaktifkan di luar jam 06.00–07.00!", "error");
                    return;
                }
            }
        } else {
            // Hotkey Kendaraan Nonlistrik: 1 s/d 9
            const nonEvMap = {
                1: "MCY",      // Kolom C
                2: "CAR",      // Kolom D
                3: "LGV",      // Kolom E
                4: "MGV",      // Kolom F
                5: "HGV",      // Kolom G
                6: "MIKROBUS", // Kolom H
                7: "LBUS",     // Kolom I
                8: "MBUS",     // Kolom J
                9: "HBUS"      // Kolom K
            };
            const targetCode = nonEvMap[digit];
            if (targetCode) {
                // Hormati filter kelompok K1 / K2 bila aktif
                if (AppState.templateData && AppState.templateData.categories_non_ev) {
                    const catObj = AppState.templateData.categories_non_ev.find(c => c.code === targetCode);
                    if (catObj) {
                        if (AppState.kelompokFilter === "K1" && catObj.k_type !== "K1") return;
                        if (AppState.kelompokFilter === "K2" && catObj.k_type !== "K2") return;
                    }
                }
                e.preventDefault();
                incrementTally(targetCode, 1);
                return;
            }
        }
    });
}

// 19. TAMPILAN VIEW EXCEL (SPREADSHEET TABLE RENDERER)
function renderExcelSpreadsheetTable() {
    const tableBody = document.getElementById("excelSpreadsheetTableBody");
    if (!tableBody || !AppState.templateData) return;
    tableBody.innerHTML = "";

    const intervals = AppState.templateData.intervals;
    intervals.forEach(item => {
        const row = item.row;
        const sess = AppState.countsData[String(row)] || { counts: {} };
        const counts = sess.counts || {};
        const isEv = item.is_ev_active;
        const isActiveSession = (row === AppState.activeRow);

        let rowBgClass = "";
        if (isActiveSession) {
            rowBgClass = "bg-sky-900/50 ring-2 ring-sky-400 font-bold";
        } else if (item.color === "blue") {
            rowBgClass = "bg-sky-950/20 hover:bg-sky-900/30";
        } else {
            rowBgClass = "bg-amber-950/20 hover:bg-amber-900/30";
        }

        const tr = document.createElement("tr");
        tr.className = `border-b border-slate-800 transition-colors cursor-pointer ${rowBgClass}`;
        tr.title = `Klik untuk beralih ke sesi B.${row} (${item.time_label})`;
        tr.addEventListener("click", () => {
            setActiveSession(row);
            switchViewMode("counter");
        });

        const mcyTot = (counts["MCY"] || 0) + (isEv ? (counts["MCY_EV"] || 0) : 0);
        const carTot = (counts["CAR"] || 0) + (isEv ? (counts["CAR_EV"] || 0) : 0);
        const lgvTot = counts["LGV"] || 0;
        const mgvTot = counts["MGV"] || 0;
        const hgvTot = counts["HGV"] || 0;
        const mikroTot = (counts["MIKROBUS"] || 0) + (isEv ? (counts["MIKROBUS_EV"] || 0) : 0);
        const busTot = (
            (counts["LBUS"] || 0) + (counts["MBUS"] || 0) + (counts["HBUS"] || 0) +
            (isEv ? ((counts["LBUS_EV"] || 0) + (counts["MBUS_EV"] || 0) + (counts["HBUS_EV"] || 0)) : 0)
        );
        const nonMcyTot = carTot + lgvTot + mgvTot + hgvTot + mikroTot + busTot;
        const grandTot = nonMcyTot + mcyTot;

        let html = `
            <td class="p-2 border-r border-slate-750 text-center font-mono ${isActiveSession ? 'text-sky-300 font-black' : 'text-slate-400'}">${row}</td>
            <td class="p-2 border-r border-slate-750 text-center font-mono font-bold text-white">${item.time_label}</td>
            
            <td class="p-2 border-r border-slate-750 text-center font-mono text-slate-200">${counts["MCY"] !== undefined ? counts["MCY"] : 0}</td>
            <td class="p-2 border-r border-slate-750 text-center font-mono text-slate-200">${counts["CAR"] !== undefined ? counts["CAR"] : 0}</td>
            <td class="p-2 border-r border-slate-750 text-center font-mono text-slate-200">${counts["LGV"] !== undefined ? counts["LGV"] : 0}</td>
            <td class="p-2 border-r border-slate-750 text-center font-mono text-slate-200">${counts["MGV"] !== undefined ? counts["MGV"] : 0}</td>
            <td class="p-2 border-r border-slate-750 text-center font-mono text-slate-200">${counts["HGV"] !== undefined ? counts["HGV"] : 0}</td>
            <td class="p-2 border-r border-slate-750 text-center font-mono text-slate-200">${counts["MIKROBUS"] !== undefined ? counts["MIKROBUS"] : 0}</td>
            <td class="p-2 border-r border-slate-750 text-center font-mono text-slate-200">${counts["LBUS"] !== undefined ? counts["LBUS"] : 0}</td>
            <td class="p-2 border-r border-slate-750 text-center font-mono text-slate-200">${counts["MBUS"] !== undefined ? counts["MBUS"] : 0}</td>
            <td class="p-2 border-r border-slate-750 text-center font-mono text-slate-200">${counts["HBUS"] !== undefined ? counts["HBUS"] : 0}</td>
            
            <td class="p-2 border-r border-slate-750 text-center font-mono ${isEv ? 'text-emerald-300' : 'text-slate-600 bg-slate-900/80'}">${isEv ? (counts["MCY_EV"] !== undefined ? counts["MCY_EV"] : 0) : '-'}</td>
            <td class="p-2 border-r border-slate-750 text-center font-mono ${isEv ? 'text-emerald-300' : 'text-slate-600 bg-slate-900/80'}">${isEv ? (counts["CAR_EV"] !== undefined ? counts["CAR_EV"] : 0) : '-'}</td>
            <td class="p-2 border-r border-slate-750 text-center font-mono ${isEv ? 'text-emerald-300' : 'text-slate-600 bg-slate-900/80'}">${isEv ? (counts["MIKROBUS_EV"] !== undefined ? counts["MIKROBUS_EV"] : 0) : '-'}</td>
            <td class="p-2 border-r border-slate-750 text-center font-mono ${isEv ? 'text-emerald-300' : 'text-slate-600 bg-slate-900/80'}">${isEv ? (counts["LBUS_EV"] !== undefined ? counts["LBUS_EV"] : 0) : '-'}</td>
            <td class="p-2 border-r border-slate-750 text-center font-mono ${isEv ? 'text-emerald-300' : 'text-slate-600 bg-slate-900/80'}">${isEv ? (counts["MBUS_EV"] !== undefined ? counts["MBUS_EV"] : 0) : '-'}</td>
            <td class="p-2 border-r border-slate-750 text-center font-mono ${isEv ? 'text-emerald-300' : 'text-slate-600 bg-slate-900/80'}">${isEv ? (counts["HBUS_EV"] !== undefined ? counts["HBUS_EV"] : 0) : '-'}</td>
            
            <td class="p-2 border-r border-slate-750 text-center font-mono font-bold text-sky-400 bg-slate-850/80">${mcyTot}</td>
            <td class="p-2 border-r border-slate-750 text-center font-mono font-bold text-sky-400 bg-slate-850/80">${carTot}</td>
            <td class="p-2 border-r border-slate-750 text-center font-mono font-bold text-sky-400 bg-slate-850/80">${busTot}</td>
            <td class="p-2 border-r border-slate-750 text-center font-mono font-bold text-amber-300 bg-slate-850/80">${nonMcyTot}</td>
            <td class="p-2 text-center font-mono font-extrabold text-emerald-400 bg-slate-850/80">${grandTot}</td>
        `;

        tr.innerHTML = html;
        tableBody.appendChild(tr);
    });
}

// 20. Simpan & Tulis ke Excel
async function syncToExcel() {
    if (!AppState.currentFilepath) {
        showToast("Belum ada file Excel yang dipilih!", "error");
        return;
    }

    try {
        showToast("Menulis hasil ke sel biru & kuning file Excel...", "info");
        const res = await fetch("/api/sync-to-excel", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({
                filepath: AppState.currentFilepath,
                counts_data: AppState.countsData,
                catatan: AppState.noteText
            })
        });
        const json = await res.json();
        if (json.success) {
            window.soundFx.playSuccess();
            showToast(json.message, "success");
            getFilteredIntervals().forEach(i => {
                if (AppState.countsData[String(i.row)]) {
                    AppState.countsData[String(i.row)].status = "verified";
                }
            });
            renderSessionTimeline();
            updateSessionCounterSummary();
            renderExcelSpreadsheetTable();
        } else {
            showToast(json.message, "error");
        }
    } catch (err) {
        showToast("Gagal menyimpan ke Excel: " + err, "error");
    }
}

// 21. Unduh File Excel
function downloadCurrentExcel() {
    if (!AppState.currentFilepath) return;
    window.location.href = `/api/download?filepath=${encodeURIComponent(AppState.currentFilepath)}`;
}

// 22. Modal Review & QC Checklist
async function openReviewModal() {
    const modal = document.getElementById("reviewModal");
    modal.classList.remove("hidden");

    try {
        const res = await fetch("/api/qc-check", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({
                counts_data: AppState.countsData,
                package: AppState.packageFilter,
                kelompok: AppState.kelompokFilter
            })
        });
        const json = await res.json();
        renderQCResults(json.qc);
    } catch (e) {
        console.error(e);
    }

    renderReviewTable();
}

function closeReviewModal() {
    document.getElementById("reviewModal").classList.add("hidden");
}

function renderQCResults(qc) {
    const qcContainer = document.getElementById("qcChecklistContainer");
    if (!qc) return;

    let html = `
        <div class="p-4 rounded-xl border ${qc.status === 'PASS' ? 'bg-emerald-950/40 border-emerald-700/60' : 'bg-amber-950/40 border-amber-700/60'} mb-4">
            <div class="flex items-center gap-2 font-bold ${qc.status === 'PASS' ? 'text-emerald-400' : 'text-amber-400'}">
                <i class="fas ${qc.status === 'PASS' ? 'fa-check-circle' : 'fa-exclamation-triangle'}"></i>
                Status QC: ${qc.status === 'PASS' ? 'Lolos Pemeriksaan Mandiri (Siap Submit)' : 'Perlu Ditinjau'}
            </div>
            <div class="text-xs text-slate-300 mt-1">
                Kelengkapan Paket: <span class="font-bold font-mono text-white">${qc.progress_percent}%</span> · Total Kendaraan: <span class="font-bold font-mono text-white">${qc.total_vehicles}</span>
            </div>
        </div>
    `;

    if (qc.issues && qc.issues.length > 0) {
        html += '<div class="space-y-1.5 mb-4">';
        qc.issues.forEach(iss => {
            html += `
                <div class="text-xs p-2.5 rounded-lg bg-rose-950/40 border border-rose-800/60 text-rose-200 flex items-start gap-2">
                    <i class="fas fa-times-circle text-rose-400 mt-0.5"></i>
                    <div>${iss.message}</div>
                </div>
            `;
        });
        html += '</div>';
    }

    qcContainer.innerHTML = html;
}

function renderReviewTable() {
    const tableBody = document.getElementById("reviewTableBody");
    if (!tableBody || !AppState.templateData) return;
    tableBody.innerHTML = "";

    const intervals = getFilteredIntervals();
    intervals.forEach(i => {
        const sess = AppState.countsData[String(i.row)] || { counts: {} };
        const tr = document.createElement("tr");
        tr.className = "border-b border-slate-800 hover:bg-slate-800/50 text-xs";

        let totalRow = 0;
        let cellsHtml = `
            <td class="p-2 font-mono font-bold text-sky-400">B.${i.row}</td>
            <td class="p-2 font-mono text-white">${i.time_label}</td>
            <td class="p-2">
                <span class="px-2 py-0.5 rounded text-[10px] ${
                    i.color === 'blue' ? 'bg-sky-950 text-sky-300 border border-sky-800' : 'bg-yellow-950 text-yellow-300 border border-yellow-800'
                }">${i.color.toUpperCase()}</span>
            </td>
        `;

        AppState.templateData.categories_non_ev.forEach(cat => {
            const val = sess.counts[cat.code] !== undefined ? sess.counts[cat.code] : 0;
            if (typeof val === "number") totalRow += val;
            cellsHtml += `<td class="p-2 text-center font-mono">${val}</td>`;
        });

        cellsHtml += `<td class="p-2 text-right font-mono font-bold text-emerald-400">${totalRow}</td>`;

        tr.innerHTML = cellsHtml;
        tableBody.appendChild(tr);
    });
}

// 23. Toast Notification
function showToast(message, type = "info") {
    const toast = document.getElementById("toast");
    if (!toast) return;

    let icon = '<i class="fas fa-info-circle text-sky-400"></i>';
    let border = 'border-sky-500/50 bg-slate-900';

    if (type === "success") {
        icon = '<i class="fas fa-check-circle text-emerald-400"></i>';
        border = 'border-emerald-500/50 bg-slate-900';
    } else if (type === "error") {
        icon = '<i class="fas fa-exclamation-circle text-rose-400"></i>';
        border = 'border-rose-500/50 bg-slate-900';
    }

    toast.className = `fixed bottom-6 right-6 z-50 flex items-center gap-3 px-4 py-3 rounded-xl border ${border} shadow-2xl text-sm font-medium text-white transition-all transform duration-300 opacity-100 translate-y-0`;
    toast.innerHTML = `${icon} <span>${message}</span>`;

    clearTimeout(window.toastTimer);
    window.toastTimer = setTimeout(() => {
        toast.className = "fixed bottom-6 right-6 z-50 opacity-0 translate-y-4 pointer-events-none transition-all duration-300";
    }, 3200);
}
