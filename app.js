
// ==========================================
// 1. ΚΕΝΤΡΙΚΗ ΚΑΤΑΣΤΑΣΗ (STATE) & ΡΥΘΜΙΣΕΙΣ
// ==========================================
const state = {
    datasets: { d1: [], d2: [], d3: [], d4: [] },
    allSections: [],
    currentFiltered: [],
    teachersBySection: new Map(),
    studentsBySection: new Map(),
    teachersInfoByAm: new Map(),
    epalContactMap: new Map(),
    selectedTmimaId: ''
};

const FILE_CONFIGS = [
    { index: 1, key: 'd1', url: 'Τμήματα.xlsx', label: '1. Τμήματα', match: (name, row) => name.includes('τμημα') || 'ΠΛΗΘΟΣ ΕΓΚΡΙΣΗΣ ΤΜΗΜΑΤΟΣ' in row },
    { index: 2, key: 'd2', url: 'Αναθέσεις.xlsx', label: '2. Αναθέσεις', match: (name, row) => name.includes('αναθεσ') || 'Αρ. Τμήματος' in row },
    { index: 3, key: 'd3', url: 'Μαθητευόμενοι.xlsx', label: '3. Μαθητευόμενοι', match: (name, row) => name.includes('μαθητ') || ('Α/Α ΤΜΗΜΑΤΟΣ' in row && 'ΕΠΩΝΥΜΟ' in row) },
    { index: 4, key: 'd4', url: 'Εκπαιδευτικοί.xlsx', label: '4. Εκπαιδευτικοί', match: (name, row) => name.includes('εκπαιδ') || ('ΑΜ' in row && !('Αρ. Τμήματος' in row)) }
];

// ==========================================
// 2. ΒΟΗΘΗΤΙΚΕΣ ΣΥΝΑΡΤΗΣΕΙΣ (UTILS)
// ==========================================
const cleanKey = (val) => String(val ?? '').trim();

const normalizeText = (str) =>
    cleanKey(str)
        .toLowerCase()
        .normalize('NFD')
        .replace(/[\u0300-\u036f]/g, '');

const normalizeAm = (val) => cleanKey(val).replace(/^0+/, '');

function escapeHTML(str) {
    return String(str ?? '')
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;');
}

function debounce(fn, delay = 150) {
    let timer;
    return (...args) => {
        clearTimeout(timer);
        timer = setTimeout(() => fn(...args), delay);
    };
}

function formatDate(val) {
    const s = cleanKey(val);
    if (!s) return '';
    if (/^\d{5}$/.test(s)) {
        const d = new Date(Math.round((Number(s) - 25569) * 86400 * 1000));
        return d.toLocaleDateString('el-GR', { day: '2-digit', month: '2-digit', year: 'numeric' });
    }
    const isoMatch = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
    if (isoMatch) {
        const yyyy = isoMatch[1];
        const mm = isoMatch[2].padStart(2, '0');
        const dd = isoMatch[3].padStart(2, '0');
        return `${dd}/${mm}/${yyyy}`;
    }
    return s;
}

function getColByKeywords(row, keywords) {
    for (const [key, val] of Object.entries(row)) {
        const normKey = normalizeText(key);
        if (keywords.some(kw => normKey.includes(kw))) {
            const v = cleanKey(val);
            if (v) return v;
        }
    }
    return '';
}

function isStudentStopped(s) {
    const rawStop = cleanKey(
        s['ΗΜΕΡΟΜΗΝΙΑ ΔΙΑΚΟΠΗΣ'] ||
        s['ΗΜ/ΝΙΑ ΔΙΑΚΟΠΗΣ'] ||
        s['ΗΜ. ΔΙΑΚΟΠΗΣ'] ||
        s['ΔΙΑΚΟΠΗ']
    ) || getColByKeywords(s, ['διακοπ']);

    const formatted = formatDate(rawStop);
    return Boolean(formatted && formatted !== '-' && formatted !== '—');
}


function isStudentEneegyl(s) {
    // 1. Άμεσος έλεγχος στη στήλη ΕΝΕΕΓΥΛ (στήλη Q)
    const directVal = normalizeText(s['ΕΝΕΕΓΥΛ'] || s['ΕΝΕΕΓΥΛ '] || s[' ΕΝΕΕΓΥΛ']);
    if (directVal && directVal !== 'οχι' && directVal !== '0' && directVal !== '-' && directVal !== '—') {
        return true;
    }

    // 2. Εφεδρικός έλεγχος σε οποιαδήποτε στήλη έχει όνομα ή τιμή σχετική με ΕΝΕΕΓΥΛ
    for (const [key, rawVal] of Object.entries(s)) {
        const k = normalizeText(key).replace(/[.\-\/\s]/g, '');
        const v = normalizeText(rawVal).replace(/[.\-\/\s]/g, '');
        if (k.includes('ενεεγυλ') && v && v !== 'οχι' && v !== '0' && v !== '—') {
            return true;
        }
        if (v.includes('ενεεγυλ')) {
            return true;
        }
    }
    return false;
}



function updateFileStatus(index, label, count, isOk = true) {
    const el = document.getElementById(`status${index}`);
    if (!el) return;
    el.textContent = isOk ? `${label}: ✔ ${count}` : `${label}: ⚠️ Επιλέξτε το χειροκίνητα`;
    el.className = isOk ? 'status-pill loaded' : 'status-pill warn';
}

function downloadCsvFile(filename, headers, rows) {
    const csvRows = rows.map(row =>
        row.map(v => `"${String(v ?? '').replace(/"/g, '""')}"`).join(';')
    );
    const csvContent = '\uFEFF' + [headers.join(';'), ...csvRows].join('\r\n');
    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    a.click();
    URL.revokeObjectURL(url);
}







// ==========================================
// 3. ΑΝΑΓΝΩΣΗ XLSX (PARSER)
// ==========================================
async function parseXlsxBuffer(buf) {
    const view = new DataView(buf);
    const bytes = new Uint8Array(buf);
    const dec = new TextDecoder('utf-8');

    let eocd = -1;
    for (let i = bytes.length - 22; i >= Math.max(0, bytes.length - 65557); i--) {
        if (view.getUint32(i, true) === 0x06054b50) { eocd = i; break; }
    }
    if (eocd === -1) throw new Error('Μη έγκυρο αρχείο .xlsx');

    const totalEntries = view.getUint16(eocd + 10, true);
    let cdOffset = view.getUint32(eocd + 16, true);
    const entries = {};

    for (let i = 0; i < totalEntries; i++) {
        if (view.getUint32(cdOffset, true) !== 0x02014b50) break;
        const method = view.getUint16(cdOffset + 10, true);
        const compSize = view.getUint32(cdOffset + 20, true);
        const nameLen = view.getUint16(cdOffset + 28, true);
        const extraLen = view.getUint16(cdOffset + 30, true);
        const commentLen = view.getUint16(cdOffset + 32, true);
        const localOffset = view.getUint32(cdOffset + 42, true);
        const name = dec.decode(bytes.subarray(cdOffset + 46, cdOffset + 46 + nameLen));
        entries[name] = { method, compSize, localOffset };
        cdOffset += 46 + nameLen + extraLen + commentLen;
    }

    async function getXmlText(path) {
        const ent = entries[path];
        if (!ent) return null;
        const lNameLen = view.getUint16(ent.localOffset + 26, true);
        const lExtraLen = view.getUint16(ent.localOffset + 28, true);
        const dataStart = ent.localOffset + 30 + lNameLen + lExtraLen;
        const compData = bytes.subarray(dataStart, dataStart + ent.compSize);
        if (ent.method === 0) return dec.decode(compData);
        const ds = new DecompressionStream('deflate-raw');
        const stream = new Blob([compData]).stream().pipeThrough(ds);
        return await new Response(stream).text();
    }

    const parser = new DOMParser();
    const sharedStrings = [];
    const ssXml = await getXmlText('xl/sharedStrings.xml');
    if (ssXml) {
        const doc = parser.parseFromString(ssXml, 'application/xml');
        doc.querySelectorAll('si').forEach(si => {
            let txt = '';
            si.querySelectorAll('t').forEach(t => txt += t.textContent);
            sharedStrings.push(txt);
        });
    }

    const sheetPath = Object.keys(entries)
        .filter(k => /^xl\/worksheets\/sheet\d+\.xml$/i.test(k))
        .sort()
        .shift();

    if (!sheetPath) throw new Error('Δεν βρέθηκε φύλλο εργασίας στο αρχείο.');

    const sheetXml = await getXmlText(sheetPath);
    const sheetDoc = parser.parseFromString(sheetXml, 'application/xml');

    const colToIdx = (ref) => {
        const letters = ref.replace(/\d+/g, '');
        let n = 0;
        for (let i = 0; i < letters.length; i++) n = n * 26 + (letters.charCodeAt(i) - 64);
        return n - 1;
    };

    const matrix = [];
    sheetDoc.querySelectorAll('sheetData > row').forEach(rowEl => {
        const rowArr = [];
        rowEl.querySelectorAll('c').forEach(c => {
            const idx = colToIdx(c.getAttribute('r') || 'A1');
            const type = c.getAttribute('t');
            let val = '';
            if (type === 's') {
                const v = c.querySelector('v');
                val = v ? (sharedStrings[Number(v.textContent)] ?? '') : '';
            } else if (type === 'inlineStr') {
                c.querySelectorAll('t').forEach(t => val += t.textContent);
            } else {
                const v = c.querySelector('v');
                val = v ? v.textContent : '';
            }
            rowArr[idx] = val.trim();
        });
        if (rowArr.some(x => x)) matrix.push(rowArr);
    });

    const headers = matrix.shift() || [];
    return matrix.map(r => {
        const obj = {};
        headers.forEach((h, i) => { if (h) obj[h] = r[i] ?? ''; });
        return obj;
    });
}







// ==========================================
// 4. ΦΟΡΤΩΣΗ & ΕΠΕΞΕΡΓΑΣΙΑ ΔΕΔΟΜΕΝΩΝ
// ==========================================
async function autoLoadExcelFiles() {
    for (const cfg of FILE_CONFIGS) {
        try {
            const res = await fetch(`${encodeURI(cfg.url)}?t=${Date.now()}`);
            if (!res.ok) throw new Error('Δεν βρέθηκε');
            const buf = await res.arrayBuffer();
            const parsed = await parseXlsxBuffer(buf);

            state.datasets[cfg.key] = parsed;
            updateFileStatus(cfg.index, cfg.label, parsed.length, true);
        } catch {
            updateFileStatus(cfg.index, cfg.label, 0, false);
        }
    }
    rebuildData();
}

function rebuildData() {
    const map = new Map();
    state.teachersBySection = new Map();
    state.studentsBySection = new Map();
    state.teachersInfoByAm = new Map();
    state.epalContactMap = new Map();

    // 4. Στοιχεία επικοινωνίας εκπαιδευτικών
    (state.datasets.d4 || []).forEach(r => {
        const rawAm = cleanKey(r['ΑΜ'] || r['Α.Μ.'] || r['Αριθμός Μητρώου']) || getColByKeywords(r, ['αμ', 'μητρω']);
        const am = normalizeAm(rawAm);
        if (!am) return;

        const phone = cleanKey(r['ΚΙΝΗΤΟ'] || r['Κινητό'] || r['ΤΗΛΕΦΩΝΟ'] || r['Τηλέφωνο']) || getColByKeywords(r, ['κινητ', 'τηλεφ']);
        const email = cleanKey(r['EMAIL'] || r['Email'] || r['E-MAIL'] || r['E-mail']) || getColByKeywords(r, ['mail', 'ταχυδρομ', 'ηλεκτρον']);
        state.teachersInfoByAm.set(am, { phone, email });
    });

    // 1. Τμήματα & Επαφές ΕΠΑΛ
    (state.datasets.d1 || []).forEach(r => {
        const id = cleanKey(r['Α/Α ΤΜΗΜΑΤΟΣ']);
        const pde = cleanKey(r['ΠΔΕ'] || r['Π.Δ.Ε.']) || getColByKeywords(r, ['πδε', 'περιφερειακ']) || 'ΧΩΡΙΣ ΠΔΕ';
        const dde = cleanKey(r['ΔΔΕ']) || 'ΧΩΡΙΣ ΔΔΕ';
        const epal = cleanKey(r['ΕΠΑΛ']) || 'ΧΩΡΙΣ ΕΠΑΛ';

        const epalKey = `${dde}||${epal}`;
        const phone = cleanKey(r['ΤΗΛΕΦΩΝΟ'] || r['Τηλέφωνο'] || r['ΤΗΛ'] || r['ΤΗΛ. ΕΠΑΛ'] || r['ΤΗΛΕΦΩΝΟ ΕΠΑΛ']) || getColByKeywords(r, ['τηλεφ', 'τηλ', 'κινητ']);
        const email = cleanKey(r['EMAIL'] || r['Email'] || r['E-MAIL'] || r['EMAIL ΕΠΑΛ'] || r['E-MAIL ΕΠΑΛ']) || getColByKeywords(r, ['mail', 'ταχυδρομ', 'ηλεκτρον']);

        if (phone || email) {
            const existing = state.epalContactMap.get(epalKey) || { phone: '', email: '' };
            state.epalContactMap.set(epalKey, {
                phone: existing.phone || phone,
                email: existing.email || email
            });
        }

        if (id) {
            map.set(id, {
                id, pde, dde, epal,
                day: cleanKey(r['ΗΜΕΡΑ ΔΙΔΑΣΚΑΛΙΑΣ']),
                eid: cleanKey(r['ΕΙΔΙΚΟΤΗΤΑ']) || 'ΧΩΡΙΣ ΕΙΔΙΚΟΤΗΤΑ',
                egkrisi: cleanKey(r['ΠΛΗΘΟΣ ΕΓΚΡΙΣΗΣ ΤΜΗΜΑΤΟΣ'])
            });
        }
    });

    // 2. Αναθέσεις
    (state.datasets.d2 || []).forEach(r => {
        const id = cleanKey(r['Αρ. Τμήματος']);
        if (!id) return;

        if (!state.teachersBySection.has(id)) state.teachersBySection.set(id, []);
        state.teachersBySection.get(id).push(r);

        if (!map.has(id)) {
            map.set(id, {
                id,
                pde: cleanKey(r['Π.Δ.Ε.'] || r['ΠΔΕ']) || getColByKeywords(r, ['πδε']) || 'ΧΩΡΙΣ ΠΔΕ',
                dde: cleanKey(r['Δ.Δ.Ε.']) || 'ΧΩΡΙΣ ΔΔΕ',
                epal: cleanKey(r['ΕΠΑΛ']) || 'ΧΩΡΙΣ ΕΠΑΛ',
                day: '',
                eid: cleanKey(r['Ειδικότητα']) || 'ΧΩΡΙΣ ΕΙΔΙΚΟΤΗΤΑ',
                egkrisi: ''
            });
        }
    });

    // 3. Μαθητευόμενοι
    (state.datasets.d3 || []).forEach(r => {
        const id = cleanKey(r['Α/Α ΤΜΗΜΑΤΟΣ']);
        if (!id) return;

        if (!state.studentsBySection.has(id)) state.studentsBySection.set(id, []);
        state.studentsBySection.get(id).push(r);

        if (!map.has(id)) {
            map.set(id, {
                id,
                pde: cleanKey(r['ΠΔΕ'] || r['Π.Δ.Ε.']) || getColByKeywords(r, ['πδε']) || 'ΧΩΡΙΣ ΠΔΕ',
                dde: cleanKey(r['ΔΔΕ']) || 'ΧΩΡΙΣ ΔΔΕ',
                epal: cleanKey(r['ΕΠΑΛ']) || 'ΧΩΡΙΣ ΕΠΑΛ',
                day: '',
                eid: cleanKey(r['ΕΙΔΙΚΟΤΗΤΑ']) || 'ΧΩΡΙΣ ΕΙΔΙΚΟΤΗΤΑ',
                egkrisi: ''
            });
        }
    });

      state.allSections = Array.from(map.values()).map(sec => {
        const teachers = state.teachersBySection.get(sec.id) || [];
        const students = state.studentsBySection.get(sec.id) || [];
        const stoppedCount = students.filter(isStudentStopped).length;
        const eneegylCount = students.filter(isStudentEneegyl).length;

        const teacherTokens = teachers.map(t => `${t['Εκπαιδευτικός'] || ''} ${t['ΑΜ'] || ''} ${t['Κλάδος'] || ''}`).join(' ');
        const studentTokens = students.map(s => `${s['ΕΠΩΝΥΜΟ'] || ''} ${s['ΟΝΟΜΑ'] || ''}`).join(' ');

        return {
            ...sec,
            stuCount: students.length,
            stoppedCount,
            eneegylCount,
            teaCount: teachers.length,
            searchIndex: normalizeText(`${sec.id} ${sec.eid} ${sec.epal} ${sec.dde} ${sec.pde} ${sec.day} ${teacherTokens} ${studentTokens}`)
        };
    }).sort((a, b) =>
        a.pde.localeCompare(b.pde, 'el') ||
        a.dde.localeCompare(b.dde, 'el') ||
        a.epal.localeCompare(b.epal, 'el') ||
        a.eid.localeCompare(b.eid, 'el')
    );

    populateFilterDropdowns();
    applyFilter();
}









// ==========================================
// 5. ΦΙΛΤΡΑΡΙΣΜΑ & DROPDOWNS
// ==========================================
function fillSelectOptions(selectEl, items, defaultLabel, prevValue) {
    selectEl.innerHTML = `<option value="">${defaultLabel}</option>`;
    items.forEach(val => {
        const opt = document.createElement('option');
        opt.value = val;
        opt.textContent = val;
        selectEl.appendChild(opt);
    });
    if (prevValue && items.includes(prevValue)) selectEl.value = prevValue;
}

function populateFilterDropdowns() {
    const pdeSel = document.getElementById('pdeFilter');
    const pdes = Array.from(new Set(state.allSections.map(s => s.pde))).sort((a, b) => a.localeCompare(b, 'el'));
    fillSelectOptions(pdeSel, pdes, '-- Όλες οι ΠΔΕ --', pdeSel.value);
    updateDependentDropdowns();
}

function updateDependentDropdowns() {
    const selectedPde = document.getElementById('pdeFilter').value;
    const ddeSel = document.getElementById('ddeFilter');
    const epalSel = document.getElementById('epalFilter');
    const eidSel = document.getElementById('eidFilter');

    const byPde = selectedPde ? state.allSections.filter(s => s.pde === selectedPde) : state.allSections;
    const ddes = Array.from(new Set(byPde.map(s => s.dde))).sort((a, b) => a.localeCompare(b, 'el'));
    fillSelectOptions(ddeSel, ddes, '-- Όλες οι ΔΔΕ --', ddeSel.value);

    const activeDde = ddeSel.value;
    const byDde = activeDde ? byPde.filter(s => s.dde === activeDde) : byPde;
    const epals = Array.from(new Set(byDde.map(s => s.epal))).sort((a, b) => a.localeCompare(b, 'el'));
    fillSelectOptions(epalSel, epals, '-- Όλα τα ΕΠΑΛ --', epalSel.value);

    const activeEpal = epalSel.value;
    const byEpal = activeEpal ? byDde.filter(s => s.epal === activeEpal) : byDde;
    const eids = Array.from(new Set(byEpal.map(s => s.eid))).sort((a, b) => a.localeCompare(b, 'el'));
    fillSelectOptions(eidSel, eids, '-- Όλες οι Ειδικότητες --', eidSel.value);
}

function applyFilter() {
    const pdeVal = document.getElementById('pdeFilter').value;
    const ddeVal = document.getElementById('ddeFilter').value;
    const epalVal = document.getElementById('epalFilter').value;
    const eidVal = document.getElementById('eidFilter').value;
    const statusVal = document.getElementById('statusFilter').value;
    const searchVal = normalizeText(document.getElementById('searchInput').value);

    state.currentFiltered = state.allSections.filter(s => {
        if (pdeVal && s.pde !== pdeVal) return false;
        if (ddeVal && s.dde !== ddeVal) return false;
        if (epalVal && s.epal !== epalVal) return false;
        if (eidVal && s.eid !== eidVal) return false;

        if (statusVal === 'has_eneegyl' && s.eneegylCount === 0) return false;
        if (statusVal === 'has_stopped' && s.stoppedCount === 0) return false;
        if (statusVal === 'eneegyl_and_stopped' && (s.eneegylCount === 0 || s.stoppedCount === 0)) return false;
        if (statusVal === 'zero_students' && s.stuCount !== 0) return false;
        if (statusVal === 'no_teachers' && s.teaCount !== 0) return false;

        if (searchVal && !s.searchIndex.includes(searchVal)) return false;
        return true;
    });

    closeModal();
    renderGroupedTable(state.currentFiltered);
}






// ==========================================
// 6. ΑΠΕΙΚΟΝΙΣΗ ΠΙΝΑΚΑ & ΣΤΑΤΙΣΤΙΚΩΝ
// ==========================================
function updateStatsBar(list) {
    const bar = document.getElementById('statsBar');
    if (!bar) return;

    if (state.allSections.length === 0) {
        bar.style.display = 'none';
        return;
    }

    bar.style.display = 'grid';
    const uniqueTeachers = new Set();
    let totalStudents = 0;

    list.forEach(sec => {
        totalStudents += sec.stuCount;
        (state.teachersBySection.get(sec.id) || []).forEach(t => {
            const key = normalizeAm(t['ΑΜ']) || normalizeText(t['Εκπαιδευτικός']);
            if (key) uniqueTeachers.add(key);
        });
    });

    document.getElementById('statPde').textContent = new Set(list.map(x => x.pde)).size;
    document.getElementById('statDde').textContent = new Set(list.map(x => x.dde)).size;
    document.getElementById('statEpal').textContent = new Set(list.map(x => `${x.dde}||${x.epal}`)).size;
    document.getElementById('statEid').textContent = new Set(list.map(x => x.eid)).size;
    document.getElementById('statTmimata').textContent = list.length;
    document.getElementById('statTeachers').textContent = uniqueTeachers.size;
    document.getElementById('statStudents').textContent = totalStudents;
}

function renderGroupedTable(list) {
    updateStatsBar(list);

    const card = document.getElementById('allSectionsCard');
    const wrap = document.getElementById('allSectionsTable');

    if (list.length === 0) {
        card.style.display = state.allSections.length > 0 ? 'block' : 'none';
        wrap.innerHTML = '<div class="empty-msg">Δεν βρέθηκαν τμήματα με αυτά τα κριτήρια αναζήτησης.</div>';
        return;
    }
    card.style.display = 'block';

    const pdeGroups = new Map();
    list.forEach(item => {
        if (!pdeGroups.has(item.pde)) pdeGroups.set(item.pde, new Map());
        const ddeMap = pdeGroups.get(item.pde);
        if (!ddeMap.has(item.dde)) ddeMap.set(item.dde, new Map());
        const epalMap = ddeMap.get(item.dde);
        if (!epalMap.has(item.epal)) epalMap.set(item.epal, []);
        epalMap.get(item.epal).push(item);
    });

    const formatBadge = (secCount, stuCount, stopCount, eneegylCount) => {
        const parts = [`Τμήματα: ${secCount}`, `Μαθητές: ${stuCount}`];
        if (stopCount > 0) parts.push(`⛔ ${stopCount}`);
        if (eneegylCount > 0) parts.push(`🟢 ${eneegylCount}`);
        return `<span class="group-badge">${parts.join(' | ')}</span>`;
    };

    let html = `
        <table class="main-sections-table">
            <colgroup>
                <col class="col-eid">
                <col class="col-tmima">
                <col class="col-students">
                <col class="col-lesson">
            </colgroup>
            <thead>
                <tr>
                    <th>Ειδικότητα</th>
                    <th>Τμήμα</th>
                    <th>Μαθητές</th>
                    <th>Μάθημα</th>
                </tr>
            </thead>
            <tbody>
    `;

    let pdeIdx = 0, ddeIdx = 0, epalIdx = 0;

    pdeGroups.forEach((ddeMap, pdeName) => {
        pdeIdx++;
        const pdeKey = `pde-${pdeIdx}`;
        let pdeSections = 0, pdeStudents = 0, pdeStopped = 0, pdeEneegyl = 0;

        ddeMap.forEach(epalMap => {
            epalMap.forEach(items => {
                pdeSections += items.length;
                pdeStudents += items.reduce((sum, x) => sum + x.stuCount, 0);
                pdeStopped += items.reduce((sum, x) => sum + x.stoppedCount, 0);
                pdeEneegyl += items.reduce((sum, x) => sum + x.eneegylCount, 0);
            });
        });

        html += `
            <tr class="group-pde" data-pde="${pdeKey}" data-collapsed="false">
                <td colspan="4">
                    <span class="toggle-arrow">▼</span> 🗺️ ${escapeHTML(pdeName)}
                    ${formatBadge(pdeSections, pdeStudents, pdeStopped, pdeEneegyl)}
                </td>
            </tr>
        `;

        ddeMap.forEach((epalMap, ddeName) => {
            ddeIdx++;
            const ddeKey = `dde-${ddeIdx}`;
            let ddeSections = 0, ddeStudents = 0, ddeStopped = 0, ddeEneegyl = 0;

            epalMap.forEach(items => {
                ddeSections += items.length;
                ddeStudents += items.reduce((sum, x) => sum + x.stuCount, 0);
                ddeStopped += items.reduce((sum, x) => sum + x.stoppedCount, 0);
                ddeEneegyl += items.reduce((sum, x) => sum + x.eneegylCount, 0);
            });

            html += `
                <tr class="group-dde child-of-${pdeKey}" data-dde="${ddeKey}" data-collapsed="false">
                    <td colspan="4">
                        <span class="toggle-arrow">▼</span> 🏛️ ${escapeHTML(ddeName)}
                        ${formatBadge(ddeSections, ddeStudents, ddeStopped, ddeEneegyl)}
                    </td>
                </tr>
            `;

            epalMap.forEach((items, epalName) => {
                epalIdx++;
                const epalKey = `epal-${epalIdx}`;
                const epalStudents = items.reduce((sum, x) => sum + x.stuCount, 0);
                const epalStopped = items.reduce((sum, x) => sum + x.stoppedCount, 0);
                const epalEneegyl = items.reduce((sum, x) => sum + x.eneegylCount, 0);

                const contact = state.epalContactMap.get(`${ddeName}||${epalName}`);
                const contactBadges = [];
                if (contact?.phone) {
                    const cleanTel = escapeHTML(contact.phone.replace(/\s+/g, ''));
                    contactBadges.push(`<a href="tel:${cleanTel}" class="contact-link">📞 ${escapeHTML(contact.phone)}</a>`);
                }
                if (contact?.email) {
                    const safeMail = escapeHTML(contact.email);
                    contactBadges.push(`<a href="mailto:${safeMail}" class="contact-link">✉️ ${safeMail}</a>`);
                }
                const epalContactHtml = contactBadges.length > 0
                    ? `<span class="epal-contact-info">${contactBadges.join('')}</span>`
                    : '';

                html += `
                    <tr class="group-epal child-of-${pdeKey} child-of-${ddeKey}" data-epal="${epalKey}" data-collapsed="false">
                        <td colspan="4">
                            <span class="toggle-arrow">▼</span> 🏫 ${escapeHTML(epalName)}
                            ${epalContactHtml}
                            ${formatBadge(items.length, epalStudents, epalStopped, epalEneegyl)}
                        </td>
                    </tr>
                `;

                items.forEach(item => {
                    const isSelected = item.id === state.selectedTmimaId;
                    const safeId = escapeHTML(item.id);

                    const aaBtn = `<span class="click-pill aa-pill" data-action="teachers" data-id="${safeId}" title="Προβολή αναθέσεων εκπαιδευτικών (${item.teaCount})">${safeId}</span>`;

                    const stoppedBadge = item.stoppedCount > 0
                        ? `<span class="click-pill seats-stop-pill" data-action="students" data-id="${safeId}" title="${item.stoppedCount} διακοπές μαθητείας">⛔ ${item.stoppedCount}</span>`
                        : `<span class="click-pill seats-stop-pill pill-placeholder">⛔ 0</span>`;

                    const eneegylBadge = item.eneegylCount > 0
                        ? `<span class="click-pill seats-eneegyl-pill" data-action="students" data-id="${safeId}" title="${item.eneegylCount} μαθητευόμενοι ΕΝΕΕΓΥΛ">🟢 ${item.eneegylCount}</span>`
                        : `<span class="click-pill seats-eneegyl-pill pill-placeholder">🟢 0</span>`;

                    const seatsBtn = `
                        <div class="students-cell-wrap">
                            <span class="click-pill seats-pill ${item.stuCount === 0 ? 'seats-warn' : ''}" data-action="students" data-id="${safeId}" title="Σύνολο: ${item.stuCount} (Ενεργοί: ${item.stuCount - item.stoppedCount})">${item.stuCount}</span>
                            ${stoppedBadge}
                            ${eneegylBadge}
                        </div>
                    `;

                    html += `
                        <tr class="tmima-row child-of-${pdeKey} child-of-${ddeKey} child-of-${epalKey} ${isSelected ? 'active-row' : ''}" data-id="${safeId}">
                            <td>${escapeHTML(item.eid)}</td>
                            <td>${aaBtn}</td>
                            <td>${seatsBtn}</td>
                            <td>${escapeHTML(item.day) || '—'}</td>
                        </tr>
                    `;
                });
            });
        });
    });

    html += '</tbody></table>';
    wrap.innerHTML = html;
}





// ==========================================
// 7. ΔΙΑΧΕΙΡΙΣΗ ΟΜΑΔΩΝ & MODALS
// ==========================================
function toggleEpal(epalKey, rowEl) {
    const nextState = rowEl.getAttribute('data-collapsed') !== 'true';
    rowEl.setAttribute('data-collapsed', String(nextState));
    rowEl.querySelector('.toggle-arrow').textContent = nextState ? '▶' : '▼';

    document.querySelectorAll(`.child-of-${epalKey}`).forEach(el => {
        el.style.display = nextState ? 'none' : '';
    });
}

function toggleDde(ddeKey, rowEl) {
    const nextState = rowEl.getAttribute('data-collapsed') !== 'true';
    rowEl.setAttribute('data-collapsed', String(nextState));
    rowEl.querySelector('.toggle-arrow').textContent = nextState ? '▶' : '▼';

    document.querySelectorAll(`.child-of-${ddeKey}`).forEach(el => {
        if (nextState) {
            el.style.display = 'none';
        } else if (el.classList.contains('group-epal')) {
            el.style.display = '';
        } else {
            const epalClass = Array.from(el.classList).find(c => c.startsWith('child-of-epal-'));
            const epalKey = epalClass ? epalClass.replace('child-of-', '') : '';
            const epalRow = document.querySelector(`tr.group-epal[data-epal="${epalKey}"]`);
            const epalCollapsed = epalRow?.getAttribute('data-collapsed') === 'true';
            el.style.display = epalCollapsed ? 'none' : '';
        }
    });
}

function togglePde(pdeKey, rowEl) {
    const nextState = rowEl.getAttribute('data-collapsed') !== 'true';
    rowEl.setAttribute('data-collapsed', String(nextState));
    rowEl.querySelector('.toggle-arrow').textContent = nextState ? '▶' : '▼';

    document.querySelectorAll(`.child-of-${pdeKey}`).forEach(el => {
        if (nextState) {
            el.style.display = 'none';
        } else if (el.classList.contains('group-dde')) {
            el.style.display = '';
        } else {
            const ddeClass = Array.from(el.classList).find(c => c.startsWith('child-of-dde-'));
            const ddeKey = ddeClass ? ddeClass.replace('child-of-', '') : '';
            const ddeRow = document.querySelector(`tr.group-dde[data-dde="${ddeKey}"]`);
            if (ddeRow?.getAttribute('data-collapsed') === 'true') {
                el.style.display = 'none';
                return;
            }

            if (el.classList.contains('group-epal')) {
                el.style.display = '';
            } else {
                const epalClass = Array.from(el.classList).find(c => c.startsWith('child-of-epal-'));
                const epalKey = epalClass ? epalClass.replace('child-of-', '') : '';
                const epalRow = document.querySelector(`tr.group-epal[data-epal="${epalKey}"]`);
                const epalCollapsed = epalRow?.getAttribute('data-collapsed') === 'true';
                el.style.display = epalCollapsed ? 'none' : '';
            }
        }
    });
}

function toggleAllGroups(collapse) {
    document.querySelectorAll('tr.group-pde, tr.group-dde, tr.group-epal').forEach(row => {
        row.setAttribute('data-collapsed', String(collapse));
        const arrow = row.querySelector('.toggle-arrow');
        if (arrow) arrow.textContent = collapse ? '▶' : '▼';
    });
    document.querySelectorAll('tr.group-dde, tr.group-epal, tr.tmima-row').forEach(row => {
        row.style.display = collapse ? 'none' : '';
    });
}

function openModal(titleText, bodyHtml, tmimaId) {
    state.selectedTmimaId = tmimaId;
    document.querySelectorAll('tr.tmima-row').forEach(tr => {
        tr.classList.toggle('active-row', tr.getAttribute('data-id') === tmimaId);
    });
    document.getElementById('modalTitle').textContent = titleText;
    document.getElementById('modalContent').innerHTML = bodyHtml;
    document.getElementById('modalOverlay').style.display = 'flex';
}

function closeModal() {
    document.getElementById('modalOverlay').style.display = 'none';
}

function showTeachers(tmimaId) {
    const sec = state.allSections.find(s => s.id === tmimaId);
    const subtitle = sec ? ` — ${sec.epal} (${sec.eid})` : '';
    const teachers = state.teachersBySection.get(tmimaId) || [];

    const bodyHtml = teachers.length > 0 ? `
        <div class="table-wrap">
            <table>
                <thead>
                    <tr>
                        <th>Εκπαιδευτικός</th><th>Κλάδος</th><th>Ιδιότητα</th><th>ΑΜ</th>
                        <th>Αρ. Πρακτικού</th><th>Ημερομηνία</th><th>Ώρες</th><th>Τύπος</th><th>Έγκριση</th>
                    </tr>
                </thead>
                <tbody>
                    ${teachers.map(t => {
                        const rawAm = cleanKey(t['ΑΜ']);
                        const info = state.teachersInfoByAm.get(normalizeAm(rawAm));
                        const contactParts = [];
                        if (info?.phone) {
                            const cleanTel = escapeHTML(info.phone.replace(/\s+/g, ''));
                            contactParts.push(`<a href="tel:${cleanTel}" class="contact-link">📱 ${escapeHTML(info.phone)}</a>`);
                        }
                        if (info?.email) {
                            const safeMail = escapeHTML(info.email);
                            contactParts.push(`<a href="mailto:${safeMail}" class="contact-link">✉️ ${safeMail}</a>`);
                        }
                        const contactHtml = contactParts.length ? `<div class="teacher-contact">${contactParts.join('')}</div>` : '';

                        return `
                            <tr>
                                <td><strong>${escapeHTML(t['Εκπαιδευτικός'])}</strong>${contactHtml}</td>
                                <td>${escapeHTML(t['Κλάδος'])}</td>
                                <td>${escapeHTML(t['Ιδιότητα'])}</td>
                                <td>${escapeHTML(rawAm)}</td>
                                <td>${escapeHTML(t['Αρ. Πρακτικού'])}</td>
                                <td>${escapeHTML(formatDate(t['Ημερομηνία']))}</td>
                                <td>${escapeHTML(t['Ώρες'])}</td>
                                <td>${escapeHTML(t['Τύπος'])}</td>
                                <td><span class="badge ${cleanKey(t['Έγκριση']) === 'ΝΑΙ' ? 'badge-ok' : 'badge-stop'}">${escapeHTML(t['Έγκριση'])}</span></td>
                            </tr>
                        `;
                    }).join('')}
                </tbody>
            </table>
        </div>`
        : '<div class="empty-msg">Δεν βρέθηκαν αναθέσεις εκπαιδευτικών για αυτό το τμήμα (ή δεν έχει φορτωθεί το 2ο αρχείο).</div>';

    openModal(`👨‍🏫 Αναθέσεις Εκπαιδευτικών — Τμήμα ${tmimaId} (${teachers.length})${subtitle}`, bodyHtml, tmimaId);
}

function showStudents(tmimaId) {
    const sec = state.allSections.find(s => s.id === tmimaId);
    const subtitle = sec ? ` — ${sec.epal} (${sec.eid})` : '';
    const students = state.studentsBySection.get(tmimaId) || [];

    const bodyHtml = students.length > 0 ? `
        <div class="table-wrap">
            <table>
                <thead>
                    <tr>
                        <th>Α/Α</th><th>Επώνυμο</th><th>Όνομα</th><th>Φύλο</th>
                        <th>Κινητό</th><th>Email</th><th>Ενεργός</th>
                    </tr>
                </thead>
                <tbody>
                    ${students.map((s, idx) => {
                        const rawStop = cleanKey(
                            s['ΗΜΕΡΟΜΗΝΙΑ ΔΙΑΚΟΠΗΣ'] ||
                            s['ΗΜ/ΝΙΑ ΔΙΑΚΟΠΗΣ'] ||
                            s['ΗΜ. ΔΙΑΚΟΠΗΣ'] ||
                            s['ΔΙΑΚΟΠΗ']
                        ) || getColByKeywords(s, ['διακοπ']);

                        const formattedStop = formatDate(rawStop);
                        const isActive = !formattedStop || formattedStop === '-' || formattedStop === '—';
                        const isEneegyl = cleanKey(s['ΕΝΕΕΓΥΛ']).toUpperCase() === 'ΝΑΙ';
                        const gender = cleanKey(s['ΦΥΛΟ'] || s['Φύλο']) || getColByKeywords(s, ['φυλ']) || '—';

                        const rawPhone = cleanKey(s['ΚΙΝΗΤΟ'] || s['ΤΗΛΕΦΩΝΟ']);
                        const phoneHtml = rawPhone
                            ? `<a href="tel:${escapeHTML(rawPhone.replace(/\s+/g, ''))}" class="contact-link">${escapeHTML(rawPhone)}</a>`
                            : '—';

                        const rawEmail = cleanKey(s['EMAIL']);
                        const emailHtml = rawEmail
                            ? `<a href="mailto:${escapeHTML(rawEmail)}" class="contact-link">${escapeHTML(rawEmail)}</a>`
                            : '—';

                        // Χρωματισμός γραμμής: Κόκκινο για διακοπή, Πράσινο για ΕΝΕΕΓΥΛ
                        let rowStyle = '';
                        let firstCellStyle = '';
                        if (!isActive) {
                            rowStyle = 'style="background-color: #fef2f2;"';
                            firstCellStyle = 'style="border-left: 5px solid #dc2626; font-weight: 700; color: #991b1b;"';
                        } else if (isEneegyl) {
                            rowStyle = 'style="background-color: #dcfce7;"';
                            firstCellStyle = 'style="border-left: 5px solid #16a34a; font-weight: 700; color: #166534;"';
                        }

                        const eneegylTag = isEneegyl
                            ? `<span class="badge" style="background:#16a34a; color:#fff; margin-left:8px;">ΕΝΕΕΓΥΛ</span>`
                            : '';

                        return `
                            <tr ${rowStyle}>
                                <td ${firstCellStyle}>${idx + 1}</td>
                                <td>
                                    <button type="button" class="student-link" data-tmima="${escapeHTML(tmimaId)}" data-idx="${idx}">
                                        ${escapeHTML(s['ΕΠΩΝΥΜΟ'])}
                                    </button>
                                </td>
                                <td>
                                    <button type="button" class="student-link" data-tmima="${escapeHTML(tmimaId)}" data-idx="${idx}">
                                        ${escapeHTML(s['ΟΝΟΜΑ'])}
                                    </button>
                                    ${eneegylTag}
                                </td>
                                <td>${escapeHTML(gender)}</td>
                                <td>${phoneHtml}</td>
                                <td>${emailHtml}</td>
                                <td>
                                    <span class="badge ${isActive ? 'badge-ok' : 'badge-stop'}" title="${!isActive ? 'Διακοπή: ' + escapeHTML(formattedStop) : 'Ενεργός μαθητευόμενος'}">
                                        ${isActive ? 'ΝΑΙ' : 'ΟΧΙ (' + escapeHTML(formattedStop) + ')'}
                                    </span>
                                </td>
                            </tr>
                        `;
                    }).join('')}
                </tbody>
            </table>
        </div>`
        : '<div class="empty-msg">Δεν βρέθηκαν μαθητευόμενοι για αυτό το τμήμα (ή δεν έχει φορτωθεί το 3ο αρχείο).</div>';

    openModal(`🎓 Μαθητευόμενοι — Τμήμα ${tmimaId} (${students.length})${subtitle}`, bodyHtml, tmimaId);
}

function showStudentDetails(tmimaId, studentIdx) {
    const students = state.studentsBySection.get(tmimaId) || [];
    const student = students[studentIdx];
    if (!student) return;

    const fullName = `${cleanKey(student['ΕΠΩΝΥΜΟ'])} ${cleanKey(student['ΟΝΟΜΑ'])}`.trim();
    const stopped = isStudentStopped(student);

    // Κανονικοποίηση κλειδιών (μετατροπή τελικού 'ς' σε 'σ')
    const cleanFieldKey = (k) => normalizeText(k).replace(/ς/g, 'σ');

    const excludedFields = [
        'κωδικοσ isced',
        'παρουσιεσ εργ',
        'κανονικεσ αδειεσ',
        'απουσιεσ',
        'κλειστοσ εργοδοτησ',
        'παρουσιεσ διδασκ',
        'δικαιολογημενεσ',
        'αδικαιολογητεσ',
        'ημ/νια πιθανησ ολοκληρωσησ',
        'ημ/νια ολοκληρωσησ μαθηματων',
        'mis',
        'αναρρ. σε εργασιμη',
        'αργια/μη εργασιμη',
        'τελευταια καταχωρημενη ημ/νια',
        'υπολοιπο παρουσιων',
        'πληρωτεο ποσο',
        'υπολοιπο ποσο',
        'συνολικο ποσο'
    ];

    const fieldsHtml = Object.entries(student)
        .filter(([key]) => {
            const norm = cleanFieldKey(key);
            return norm !== '' && !excludedFields.some(ex => norm === ex || norm.includes(ex));
        })
        .map(([key, rawVal]) => {
            const normKey = cleanFieldKey(key);
            let val = cleanKey(rawVal);

            // 1. Το πεδίο ΑΙΤΗΣΗ μένει αυστηρά αριθμός
            if (normKey === 'αιτηση' || normKey.includes('αιτηση')) {
                val = val ? String(Number(val) || val) : '';
            }
            // 2. Μετατροπή ημερομηνιών σε ηη/μμ/εεεε
            else if (val && (/^\d{5}$/.test(val) || /^\d{4}-\d{1,2}-\d{1,2}/.test(val))) {
                val = formatDate(val);
            }

            let displayVal = val ? escapeHTML(val) : '—';

            // 3. Ενεργοί σύνδεσμοι (link, τηλέφωνο, email)
            if (val && normKey.includes('αναλυτικη εικονα')) {
                const href = /^https?:\/\//i.test(val) ? val : `https://${val}`;
                displayVal = `<a href="${escapeHTML(href)}" target="_blank" rel="noopener noreferrer" class="contact-link">🔗 Προβολή Αναλυτικής Εικόνας</a>`;
            } else if (val && (normKey.includes('τηλεφ') || normKey.includes('κινητ'))) {
                displayVal = `<a href="tel:${escapeHTML(val.replace(/\s+/g, ''))}" class="contact-link">📱 ${escapeHTML(val)}</a>`;
            } else if (val && normKey.includes('mail')) {
                displayVal = `<a href="mailto:${escapeHTML(val)}" class="contact-link">✉️ ${escapeHTML(val)}</a>`;
            }

            return `
                <div class="detail-box">
                    <div class="detail-label">${escapeHTML(key)}</div>
                    <div class="detail-val">${displayVal}</div>
                </div>
            `;
        }).join('');

    const bodyHtml = `
        <div class="student-card-header">
            <button type="button" class="btn btn-secondary" id="backToStudentsBtn" data-tmima="${escapeHTML(tmimaId)}">
                ⬅ Πίσω στους μαθητευόμενους του τμήματος
            </button>
            <span class="badge ${stopped ? 'badge-stop' : 'badge-ok'}">
                ${stopped ? '⛔ Διακοπή Μαθητείας' : '✔ Ενεργός Μαθητευόμενος'}
            </span>
        </div>
        <div class="student-details-grid">
            ${fieldsHtml}
        </div>
    `;

    openModal(`👤 Καρτέλα Μαθητευόμενου — ${fullName}`, bodyHtml, tmimaId);
}


// ==========================================
// 8. EVENT LISTENERS
// ==========================================
window.addEventListener('DOMContentLoaded', autoLoadExcelFiles);

document.getElementById('multiFiles').addEventListener('change', async (e) => {
    const files = Array.from(e.target.files);
    if (!files.length) return;

    for (const file of files) {
        try {
            const buf = await file.arrayBuffer();
            const parsed = await parseXlsxBuffer(buf);
            const fname = normalizeText(file.name);
            const firstRow = parsed.at(0) || {};

            const matchedCfg = FILE_CONFIGS.find(cfg => cfg.match(fname, firstRow));
            if (matchedCfg) {
                state.datasets[matchedCfg.key] = parsed;
                updateFileStatus(matchedCfg.index, matchedCfg.label, parsed.length, true);
            }
        } catch (err) {
            alert(`Σφάλμα στο αρχείο ${file.name}: ${err.message}`);
        }
    }
    rebuildData();
});

document.getElementById('allSectionsTable').addEventListener('click', (e) => {
    if (e.target.closest('.epal-contact-info')) return;

    const pill = e.target.closest('.click-pill');
    if (pill) {
        const id = pill.getAttribute('data-id');
        if (pill.getAttribute('data-action') === 'teachers') showTeachers(id);
        if (pill.getAttribute('data-action') === 'students') showStudents(id);
        return;
    }

    const epalRow = e.target.closest('tr.group-epal');
    if (epalRow) {
        toggleEpal(epalRow.getAttribute('data-epal'), epalRow);
        return;
    }

    const ddeRow = e.target.closest('tr.group-dde');
    if (ddeRow) {
        toggleDde(ddeRow.getAttribute('data-dde'), ddeRow);
        return;
    }

    const pdeRow = e.target.closest('tr.group-pde');
    if (pdeRow) {
        togglePde(pdeRow.getAttribute('data-pde'), pdeRow);
    }
});

['pdeFilter', 'ddeFilter', 'epalFilter'].forEach(id => {
    document.getElementById(id).addEventListener('change', () => {
        updateDependentDropdowns();
        applyFilter();
    });
});

document.getElementById('eidFilter').addEventListener('change', applyFilter);
document.getElementById('statusFilter').addEventListener('change', applyFilter);
document.getElementById('searchInput').addEventListener('input', debounce(applyFilter, 150));

document.getElementById('resetFiltersBtn').addEventListener('click', () => {
    ['pdeFilter', 'ddeFilter', 'epalFilter', 'eidFilter', 'statusFilter', 'searchInput'].forEach(id => {
        document.getElementById(id).value = '';
    });
    updateDependentDropdowns();
    applyFilter();
});

document.getElementById('exportCsvBtn').addEventListener('click', () => {
    if (state.currentFiltered.length === 0) {
        alert('Δεν υπάρχουν δεδομένα για εξαγωγή.');
        return;
    }

    const headers = ['ΠΔΕ', 'ΔΔΕ', 'ΕΠΑΛ', 'Ειδικότητα', 'Τμήμα', 'Μαθητές', 'Διακοπές', 'Αναθέσεις Εκπ/κών', 'Μάθημα'];
    const rows = state.currentFiltered.map(s => [
        s.pde, s.dde, s.epal, s.eid, s.id, s.stuCount, s.stoppedCount, s.teaCount, s.day || ''
    ]);

    downloadCsvFile(`Τμηματα_Μαθητειας_${new Date().toISOString().slice(0, 10)}.csv`, headers, rows);
});

document.getElementById('exportFullCsvBtn').addEventListener('click', () => {
    if (state.currentFiltered.length === 0) {
        alert('Δεν υπάρχουν δεδομένα για εξαγωγή.');
        return;
    }

    const headers = [
        'ΠΔΕ', 'ΔΔΕ', 'ΕΠΑΛ', 'Ειδικότητα', 'Τμήμα', 'Ημέρα Διδασκαλίας',
        'Ρόλος', 'Ονοματεπώνυμο', 'ΑΜ / Φύλο', 'Κλάδος / Κατάσταση', 'Τηλέφωνο', 'Email'
    ];
    const rows = [];

    state.currentFiltered.forEach(s => {
        const teachers = state.teachersBySection.get(s.id) || [];
        const students = state.studentsBySection.get(s.id) || [];

        teachers.forEach(t => {
            const rawAm = cleanKey(t['ΑΜ']);
            const info = state.teachersInfoByAm.get(normalizeAm(rawAm)) || {};
            rows.push([
                s.pde, s.dde, s.epal, s.eid, s.id, s.day || '',
                'Εκπαιδευτικός', t['Εκπαιδευτικός'] || '', rawAm,
                `${t['Κλάδος'] || ''} (${t['Ιδιότητα'] || ''})`.trim(),
                info.phone || '', info.email || ''
            ]);
        });

        students.forEach(st => {
            const stopped = isStudentStopped(st);
            const gender = cleanKey(st['ΦΥΛΟ'] || st['Φύλο']) || getColByKeywords(st, ['φυλ']);
            rows.push([
                s.pde, s.dde, s.epal, s.eid, s.id, s.day || '',
                'Μαθητευόμενος', `${st['ΕΠΩΝΥΜΟ'] || ''} ${st['ΟΝΟΜΑ'] || ''}`.trim(), gender,
                stopped ? 'Διακοπή' : 'Ενεργός',
                st['ΚΙΝΗΤΟ'] || st['ΤΗΛΕΦΩΝΟ'] || '', st['EMAIL'] || ''
            ]);
        });

        if (teachers.length === 0 && students.length === 0) {
            rows.push([s.pde, s.dde, s.epal, s.eid, s.id, s.day || '', 'Χωρίς εγγραφές', '', '', '', '', '']);
        }
    });

    downloadCsvFile(`Αναλυτικη_Κατασταση_Μαθητειας_${new Date().toISOString().slice(0, 10)}.csv`, headers, rows);
});

document.getElementById('printBtn').addEventListener('click', () => window.print());
document.getElementById('expandAllBtn').addEventListener('click', () => toggleAllGroups(false));
document.getElementById('collapseAllBtn').addEventListener('click', () => toggleAllGroups(true));
document.getElementById('modalCloseBtn').addEventListener('click', closeModal);
document.getElementById('modalOverlay').addEventListener('click', (e) => {if (e.target.id === 'modalOverlay') closeModal();});
document.getElementById('modalContent').addEventListener('click', (e) => {
    const stuBtn = e.target.closest('.student-link');
    if (stuBtn) {
        const tmimaId = stuBtn.getAttribute('data-tmima');
        const idx = Number(stuBtn.getAttribute('data-idx'));
        showStudentDetails(tmimaId, idx);
        return;
    }

    const backBtn = e.target.closest('#backToStudentsBtn');
    if (backBtn) {
        showStudents(backBtn.getAttribute('data-tmima'));
    }
});


document.addEventListener('keydown', (e) => {
    if (e.key === 'Escape') closeModal();
});






// ==========================================
// 9. ΔΗΜΙΟΥΡΓΙΑ ΕΙΚΟΝΑΣ ΣΤΑΤΙΣΤΙΚΩΝ (.PNG - HD 2x)
// ==========================================
function openStatsImageModal() {
    const list = state.currentFiltered;
    if (!list.length) {
        alert('Δεν υπάρχουν δεδομένα για δημιουργία εικόνας.');
        return;
    }

    const pdeCount = new Set(list.map(x => x.pde)).size;
    const ddeCount = new Set(list.map(x => x.dde)).size;
    const epalCount = new Set(list.map(x => `${x.dde}||${x.epal}`)).size;
    const tmimataCount = list.length;

    const uniqueTeachers = new Set();
    let totalStudents = 0;
    let stoppedStudents = 0;
    let menCount = 0;
    let womenCount = 0;

    const eidMap = new Map();
    const dayMap = new Map();
    const pdeStatsMap = new Map();

    list.forEach(sec => {
        const stus = state.studentsBySection.get(sec.id) || [];
        const stopped = sec.stoppedCount;
        const active = stus.length - stopped;

        totalStudents += stus.length;
        stoppedStudents += stopped;

        stus.forEach(s => {
            const g = normalizeText(cleanKey(s['ΦΥΛΟ'] || s['Φύλο']) || getColByKeywords(s, ['φυλ']));
            if (g.startsWith('α')) menCount++;
            else if (g.startsWith('γ')) womenCount++;
        });

        eidMap.set(sec.eid, (eidMap.get(sec.eid) || 0) + stus.length);

        const dayName = sec.day ? sec.day.split(',').at(0).trim() : 'Χωρίς ημέρα';
        dayMap.set(dayName, (dayMap.get(dayName) || 0) + 1);

        const pdeObj = pdeStatsMap.get(sec.pde) || { active: 0, stopped: 0, total: 0 };
        pdeObj.active += active;
        pdeObj.stopped += stopped;
        pdeObj.total += stus.length;
        pdeStatsMap.set(sec.pde, pdeObj);

        (state.teachersBySection.get(sec.id) || []).forEach(t => {
            const key = normalizeAm(t['ΑΜ']) || normalizeText(t['Εκπαιδευτικός']);
            if (key) uniqueTeachers.add(key);
        });
    });

    const activeStudents = totalStudents - stoppedStudents;
    const topEids = Array.from(eidMap.entries()).sort((a, b) => b.at(1) - a.at(1)).slice(0, 5);
    const topDays = Array.from(dayMap.entries()).sort((a, b) => b.at(1) - a.at(1)).slice(0, 5);
    const topPdes = Array.from(pdeStatsMap.entries()).sort((a, b) => b.at(1).total - a.at(1).total).slice(0, 5);

    const activeFilters = ['pdeFilter', 'ddeFilter', 'epalFilter', 'eidFilter']
        .map(id => document.getElementById(id).value)
        .filter(Boolean);
    const filterSubtitle = activeFilters.length ? activeFilters.join('  •  ') : 'Συνολική Εικόνα (Όλες οι δομές)';

    const W = 1100, H = 880, SCALE = 2;
    const canvas = document.createElement('canvas');
    canvas.width = W * SCALE;
    canvas.height = H * SCALE;
    const ctx = canvas.getContext('2d');
    ctx.scale(SCALE, SCALE);

    ctx.fillStyle = '#f1f5f9';
    ctx.fillRect(0, 0, W, H);

    const grad = ctx.createLinearGradient(0, 0, W, 0);
    grad.addColorStop(0, '#1e3a8a');
    grad.addColorStop(1, '#2563eb');
    ctx.fillStyle = grad;
    ctx.fillRect(0, 0, W, 92);

    ctx.fillStyle = '#ffffff';
    ctx.font = 'bold 25px Segoe UI, sans-serif';
    ctx.fillText('Αναλυτικό Ταμπλό Μαθητείας ΕΠΑΛ', 32, 40);

    ctx.fillStyle = '#dbeafe';
    ctx.font = '15px Segoe UI, sans-serif';
    ctx.fillText(filterSubtitle, 32, 68);

    ctx.textAlign = 'right';
    ctx.fillText(`Ημερομηνία: ${new Date().toLocaleDateString('el-GR')}`, W - 32, 40);
    ctx.textAlign = 'left';

    const kpis = [
        { label: 'ΠΔΕ', val: pdeCount },
        { label: 'ΔΔΕ', val: ddeCount },
        { label: 'ΕΠΑΛ', val: epalCount },
        { label: 'ΤΜΗΜΑΤΑ', val: tmimataCount },
        { label: 'ΕΚΠΑΙΔΕΥΤΙΚΟΙ', val: uniqueTeachers.size },
        { label: 'ΜΑΘΗΤΕΥΟΜΕΝΟΙ', val: totalStudents }
    ];

    kpis.forEach((k, i) => {
        const x = 32 + i * 174;
        const y = 110;
        ctx.fillStyle = '#ffffff';
        ctx.strokeStyle = '#cbd5e1';
        ctx.lineWidth = 1;
        ctx.beginPath();
        ctx.roundRect(x, y, 162, 78, 8);
        ctx.fill();
        ctx.stroke();

        ctx.fillStyle = '#1e40af';
        ctx.fillRect(x, y + 74, 162, 4);

        ctx.fillStyle = '#64748b';
        ctx.font = 'bold 11px Segoe UI, sans-serif';
        ctx.textAlign = 'center';
        ctx.fillText(k.label, x + 81, y + 26);

        ctx.fillStyle = '#0f172a';
        ctx.font = 'bold 25px Segoe UI, sans-serif';
        ctx.fillText(String(k.val), x + 81, y + 58);
    });
    ctx.textAlign = 'left';






    ctx.fillStyle = '#ffffff';
    ctx.strokeStyle = '#cbd5e1';
    ctx.beginPath();
    ctx.roundRect(32, 202, W - 64, 66, 8);
    ctx.fill();
    ctx.stroke();

    ctx.fillStyle = '#1e293b';
    ctx.font = 'bold 14px Segoe UI, sans-serif';
    ctx.fillText(`Κατάσταση Μαθητευομένων:   🟢 Ενεργοί: ${activeStudents} (${totalStudents ? Math.round((activeStudents / totalStudents) * 100) : 0}%)     🔴 Διακοπές: ${stoppedStudents} (${totalStudents ? Math.round((stoppedStudents / totalStudents) * 100) : 0}%)`, 48, 224);

    const barX = 48, barY = 236, barW = W - 96, barH = 18;
    const activeRatio = totalStudents > 0 ? activeStudents / totalStudents : 1;

    ctx.fillStyle = '#fca5a5';
    ctx.beginPath();
    ctx.roundRect(barX, barY, barW, barH, 5);
    ctx.fill();

    if (activeRatio > 0) {
        ctx.fillStyle = '#16a34a';
        ctx.beginPath();
        ctx.roundRect(barX, barY, barW * activeRatio, barH, 5);
        ctx.fill();
    }

    const cardY = 284, cardH = 265, cardW = 334;
    const cardXs = new Array(32, 383, 734);

    cardXs.forEach(cx => {
        ctx.fillStyle = '#ffffff';
        ctx.strokeStyle = '#cbd5e1';
        ctx.beginPath();
        ctx.roundRect(cx, cardY, cardW, cardH, 10);
        ctx.fill();
        ctx.stroke();
    });

    // Γράφημα 1: Φύλο
    ctx.fillStyle = '#0f172a';
    ctx.font = 'bold 15px Segoe UI, sans-serif';
    ctx.fillText('Κατανομή Μαθητευομένων ανά Φύλο', cardXs.at(0) + 18, cardY + 28);

    const genderTotal = menCount + womenCount;
    const centerX = cardXs.at(0) + cardW / 2;
    const centerY = cardY + 132;
    const radius = 64;

    if (genderTotal > 0) {
        const menAngle = (menCount / genderTotal) * Math.PI * 2;
        ctx.beginPath();
        ctx.arc(centerX, centerY, radius, -Math.PI / 2, -Math.PI / 2 + menAngle);
        ctx.lineWidth = 24;
        ctx.strokeStyle = '#2563eb';
        ctx.stroke();

        ctx.beginPath();
        ctx.arc(centerX, centerY, radius, -Math.PI / 2 + menAngle, -Math.PI / 2 + Math.PI * 2);
        ctx.lineWidth = 24;
        ctx.strokeStyle = '#0d9488';
        ctx.stroke();

        const menPct = Math.round((menCount / genderTotal) * 100);
        const womenPct = 100 - menPct;

        ctx.fillStyle = '#2563eb';
        ctx.beginPath();
        ctx.arc(cardXs.at(0) + 55, cardY + 228, 6, 0, Math.PI * 2);
        ctx.fill();
        ctx.fillStyle = '#1e293b';
        ctx.font = 'bold 13px Segoe UI, sans-serif';
        ctx.fillText(`Άνδρες: ${menPct}% (${menCount})`, cardXs.at(0) + 67, cardY + 233);

        ctx.fillStyle = '#0d9488';
        ctx.beginPath();
        ctx.arc(cardXs.at(0) + 195, cardY + 228, 6, 0, Math.PI * 2);
        ctx.fill();
        ctx.fillStyle = '#1e293b';
        ctx.fillText(`Γυναίκες: ${womenPct}% (${womenCount})`, cardXs.at(0) + 207, cardY + 233);
    } else {
        ctx.fillStyle = '#64748b';
        ctx.font = '13px Segoe UI, sans-serif';
        ctx.textAlign = 'center';
        ctx.fillText('Δεν υπάρχουν στοιχεία φύλου', centerX, centerY);
        ctx.textAlign = 'left';
    }

    // Γράφημα 2: Ημέρα Διδασκαλίας
    ctx.fillStyle = '#0f172a';
    ctx.font = 'bold 15px Segoe UI, sans-serif';
    ctx.fillText('Τμήματα ανά Ημέρα Διδασκαλίας', cardXs.at(1) + 18, cardY + 28);

    const maxDayVal = topDays.length ? topDays.at(0).at(1) : 1;
    topDays.forEach(([day, count], i) => {
        const ry = cardY + 58 + i * 38;
        ctx.fillStyle = '#334155';
        ctx.font = '13px Segoe UI, sans-serif';
        ctx.fillText(day.slice(0, 13), cardXs.at(1) + 18, ry + 15);

        const bx = cardXs.at(1) + 115, bw = 165;
        ctx.fillStyle = '#e2e8f0';
        ctx.beginPath();
        ctx.roundRect(bx, ry + 2, bw, 16, 4);
        ctx.fill();

        ctx.fillStyle = '#2563eb';
        ctx.beginPath();
        ctx.roundRect(bx, ry + 2, Math.max(6, (count / maxDayVal) * bw), 16, 4);
        ctx.fill();

        ctx.fillStyle = '#0f172a';
        ctx.font = 'bold 13px Segoe UI, sans-serif';
        ctx.fillText(String(count), bx + bw + 8, ry + 15);
    });

    // Γράφημα 3: Top ΠΔΕ
    ctx.fillStyle = '#0f172a';
    ctx.font = 'bold 15px Segoe UI, sans-serif';
    ctx.fillText('Μαθητές & Διακοπές (Top ΠΔΕ)', cardXs.at(2) + 18, cardY + 28);

    const maxPdeVal = topPdes.length ? topPdes.at(0).at(1).total : 1;
    const chartBottom = cardY + 215;
    const chartMaxH = 140;

    topPdes.forEach(([pdeName, stats], i) => {
        const colX = cardXs.at(2) + 28 + i * 58;
        const colW = 34;
        const totalH = Math.max(6, (stats.total / maxPdeVal) * chartMaxH);
        const activeH = stats.total > 0 ? (stats.active / stats.total) * totalH : totalH;
        const stopH = totalH - activeH;

        ctx.fillStyle = '#2563eb';
        ctx.fillRect(colX, chartBottom - activeH, colW, activeH);

        if (stopH > 0) {
            ctx.fillStyle = '#ef4444';
            ctx.fillRect(colX, chartBottom - totalH, colW, stopH);
        }

        ctx.fillStyle = '#0f172a';
        ctx.font = 'bold 11px Segoe UI, sans-serif';
        ctx.textAlign = 'center';
        ctx.fillText(String(stats.total), colX + colW / 2, chartBottom - totalH - 5);

        const shortPde = pdeName.replace(/ΠΔΕ\s*/i, '').slice(0, 7);
        ctx.fillStyle = '#475569';
        ctx.font = '10px Segoe UI, sans-serif';
        ctx.fillText(shortPde, colX + colW / 2, chartBottom + 15);
    });
    ctx.textAlign = 'left';

    ctx.fillStyle = '#2563eb';
    ctx.fillRect(cardXs.at(2) + 85, cardY + 242, 10, 10);
    ctx.fillStyle = '#334155';
    ctx.font = '11px Segoe UI, sans-serif';
    ctx.fillText('Ενεργοί', cardXs.at(2) + 100, cardY + 251);

    ctx.fillStyle = '#ef4444';
    ctx.fillRect(cardXs.at(2) + 165, cardY + 242, 10, 10);
    ctx.fillStyle = '#334155';
    ctx.fillText('Διακοπές', cardXs.at(2) + 180, cardY + 251);

    // Γράφημα 4: Top 5 Ειδικότητες
    const botY = 566, botH = 285;
    ctx.fillStyle = '#ffffff';
    ctx.strokeStyle = '#cbd5e1';
    ctx.beginPath();
    ctx.roundRect(32, botY, W - 64, botH, 10);
    ctx.fill();
    ctx.stroke();

    ctx.fillStyle = '#0f172a';
    ctx.font = 'bold 16px Segoe UI, sans-serif';
    ctx.fillText('Top 5 Ειδικότητες (βάσει πλήθους μαθητευομένων)', 50, botY + 32);

    const maxEidVal = topEids.length ? topEids.at(0).at(1) : 1;
    topEids.forEach(([name, count], idx) => {
        const rowY = botY + 55 + idx * 42;
        const shortName = name.length > 62 ? name.slice(0, 60) + '...' : name;

        ctx.fillStyle = '#334155';
        ctx.font = '13px Segoe UI, sans-serif';
        ctx.fillText(shortName, 50, rowY + 18);

        const trackX = 520, trackW = 480;
        ctx.fillStyle = '#e2e8f0';
        ctx.beginPath();
        ctx.roundRect(trackX, rowY + 4, trackW, 18, 5);
        ctx.fill();

        ctx.fillStyle = '#1d4ed8';
        ctx.beginPath();
        ctx.roundRect(trackX, rowY + 4, Math.max(8, (count / maxEidVal) * trackW), 18, 5);
        ctx.fill();

        ctx.fillStyle = '#0f172a';
        ctx.font = 'bold 13px Segoe UI, sans-serif';
        ctx.fillText(String(count), trackX + trackW + 12, rowY + 18);
    });

    const bodyHtml = `
        <div class="stats-canvas-wrap">
            <div id="canvasHost"></div>
            <button type="button" class="btn btn-success" id="downloadPngBtn">💾 Λήψη Εικόνας Ταμπλό (.png)</button>
        </div>
    `;
    openModal('📊 Αναλυτικό Ταμπλό Στατιστικών', bodyHtml, state.selectedTmimaId);
    document.getElementById('canvasHost').appendChild(canvas);

    document.getElementById('downloadPngBtn').addEventListener('click', () => {
        const link = document.createElement('a');
        link.download = `Analytiko_Tamplo_Mathiteias_${new Date().toISOString().slice(0, 10)}.png`;
        link.href = canvas.toDataURL('image/png');
        link.click();
    });
}

document.getElementById('statsImgBtn')?.addEventListener('click', openStatsImageModal);


























