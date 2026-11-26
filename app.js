let data1 = [], data2 = [], data3 = [], data4 = [];
let allSections = [];
let currentFiltered = [];
let teachersBySection = new Map();
let studentsBySection = new Map();
let teachersInfoByAm = new Map();
let selectedTmimaId = '';
let epalContactMap = new Map();


const cleanKey = (val) => String(val ?? '').trim();

const normalizeText = (str) =>
    cleanKey(str)
        .toLowerCase()
        .normalize('NFD')
        .replace(/[\u0300-\u036f]/g, '');

function escapeHTML(str) {
    return String(str ?? '')
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;');
}

function formatDate(val) {
    const s = cleanKey(val);
    if (!s) return '';
    if (/^\d{5}$/.test(s)) {
        const d = new Date(Math.round((Number(s) - 25569) * 86400 * 1000));
        return d.toLocaleDateString('el-GR');
    }
    return s;
}

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

    const sheetPath = Object.keys(entries).filter(k => /^xl\/worksheets\/sheet\d+\.xml$/i.test(k)).sort().shift();
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







async function autoLoadExcelFiles() {
    const filesConfig = [
        { url: 'Τμήματα.xlsx', index: 1, name: '1. Τμήματα' },
        { url: 'Αναθέσεις.xlsx', index: 2, name: '2. Αναθέσεις' },
        { url: 'Μαθητευόμενοι.xlsx', index: 3, name: '3. Μαθητευόμενοι' },
        { url: 'Εκπαιδευτικοί.xlsx', index: 4, name: '4. Εκπαιδευτικοί' }
    ];

    for (const item of filesConfig) {
        const statusEl = document.getElementById(`status${item.index}`);
        try {
            const res = await fetch(`${encodeURI(item.url)}?t=${Date.now()}`);
            if (!res.ok) throw new Error('Δεν βρέθηκε');
            const buf = await res.arrayBuffer();
            const parsed = await parseXlsxBuffer(buf);

            if (item.index === 1) data1 = parsed;
            if (item.index === 2) data2 = parsed;
            if (item.index === 3) data3 = parsed;
            if (item.index === 4) data4 = parsed;

            statusEl.textContent = `${item.name}: ✔ ${parsed.length}`;
            statusEl.className = 'status-pill loaded';
        } catch (err) {
            statusEl.textContent = `${item.name}: ⚠️ Επιλέξτε το χειροκίνητα`;
            statusEl.className = 'status-pill warn';
        }
    }
    rebuildData();
}

window.addEventListener('DOMContentLoaded', autoLoadExcelFiles);

document.getElementById('multiFiles').addEventListener('change', async (e) => {
    const files = Array.from(e.target.files);
    if (files.length === 0) return;

    for (const file of files) {
        try {
            const buf = await file.arrayBuffer();
            const parsed = await parseXlsxBuffer(buf);
            const fname = normalizeText(file.name);
            const firstRow = parsed || {};

            if (fname.includes('τμημα') || ('ΠΛΗΘΟΣ ΕΓΚΡΙΣΗΣ ΤΜΗΜΑΤΟΣ' in firstRow)) {
                data1 = parsed;
                const el = document.getElementById('status1');
                el.textContent = `1. Τμήματα: ✔ ${parsed.length}`;
                el.className = 'status-pill loaded';
            } else if (fname.includes('αναθεσ') || ('Αρ. Τμήματος' in firstRow)) {
                data2 = parsed;
                const el = document.getElementById('status2');
                el.textContent = `2. Αναθέσεις: ✔ ${parsed.length}`;
                el.className = 'status-pill loaded';
            } else if (fname.includes('μαθητ') || ('Α/Α ΤΜΗΜΑΤΟΣ' in firstRow && 'ΕΠΩΝΥΜΟ' in firstRow)) {
                data3 = parsed;
                const el = document.getElementById('status3');
                el.textContent = `3. Μαθητευόμενοι: ✔ ${parsed.length}`;
                el.className = 'status-pill loaded';
            } else if (fname.includes('εκπαιδ') || ('ΑΜ' in firstRow && !('Αρ. Τμήματος' in firstRow))) {
                data4 = parsed;
                const el = document.getElementById('status4');
                el.textContent = `4. Εκπαιδευτικοί: ✔ ${parsed.length}`;
                el.className = 'status-pill loaded';
            }
        } catch (err) {
            alert(`Σφάλμα στο αρχείο ${file.name}: ` + err.message);
        }
    }
    rebuildData();
});

// Βοηθητική συνάρτηση για κανονικοποίηση ΑΜ (αφαιρεί κενά και αρχικά μηδενικά)
const normalizeAm = (val) => cleanKey(val).replace(/^0+/, '');

// Βοηθητική συνάρτηση που βρίσκει τιμή στήλης ανεξαρτήτως τόνων/πεζών-κεφαλαίων
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

function rebuildData() {
    const map = new Map();
    teachersBySection = new Map();
    studentsBySection = new Map();
    teachersInfoByAm = new Map();
    epalContactMap = new Map();

    // 4. Ευρετήριο στοιχείων επικοινωνίας εκπαιδευτικών ανά ΑΜ
    data4.forEach(r => {
        const rawAm = cleanKey(r['ΑΜ'] || r['Α.Μ.'] || r['Αριθμός Μητρώου']) || getColByKeywords(r, ['αμ', 'μητρω']);
        const am = normalizeAm(rawAm);
        if (!am) return;

        const phone = cleanKey(r['ΚΙΝΗΤΟ'] || r['Κινητό'] || r['ΤΗΛΕΦΩΝΟ'] || r['Τηλέφωνο']) || getColByKeywords(r, ['κινητ', 'τηλεφ']);
        const email = cleanKey(r['EMAIL'] || r['Email'] || r['E-MAIL'] || r['E-mail']) || getColByKeywords(r, ['mail', 'ταχυδρομ', 'ηλεκτρον']);

        teachersInfoByAm.set(am, { phone, email });
    });

    // 1. Αρχείο Τμημάτων + στοιχεία επικοινωνίας ΕΠΑΛ
    data1.forEach(r => {
        const id = cleanKey(r['Α/Α ΤΜΗΜΑΤΟΣ']);
        const dde = cleanKey(r['ΔΔΕ']) || 'ΧΩΡΙΣ ΔΔΕ';
        const epal = cleanKey(r['ΕΠΑΛ']) || 'ΧΩΡΙΣ ΕΠΑΛ';

        const epalKey = `${dde}||${epal}`;
        const phone = cleanKey(r['ΤΗΛΕΦΩΝΟ'] || r['Τηλέφωνο'] || r['ΤΗΛ'] || r['ΤΗΛ. ΕΠΑΛ'] || r['ΤΗΛΕΦΩΝΟ ΕΠΑΛ']) || getColByKeywords(r, ['τηλεφ', 'τηλ', 'κινητ']);
        const email = cleanKey(r['EMAIL'] || r['Email'] || r['E-MAIL'] || r['EMAIL ΕΠΑΛ'] || r['E-MAIL ΕΠΑΛ']) || getColByKeywords(r, ['mail', 'ταχυδρομ', 'ηλεκτρον']);

        if (phone || email) {
            const existing = epalContactMap.get(epalKey) || { phone: '', email: '' };
            epalContactMap.set(epalKey, {
                phone: existing.phone || phone,
                email: existing.email || email
            });
        }

        if (id) {
            map.set(id, {
                id,
                dde,
                epal,
                day: cleanKey(r['ΗΜΕΡΑ ΔΙΔΑΣΚΑΛΙΑΣ']),
                eid: cleanKey(r['ΕΙΔΙΚΟΤΗΤΑ']) || 'ΧΩΡΙΣ ΕΙΔΙΚΟΤΗΤΑ',
                egkrisi: cleanKey(r['ΠΛΗΘΟΣ ΕΓΚΡΙΣΗΣ ΤΜΗΜΑΤΟΣ'])
            });
        }
    });

    // 2. Αρχείο Αναθέσεων
    data2.forEach(r => {
        const id = cleanKey(r['Αρ. Τμήματος']);
        if (!id) return;

        if (!teachersBySection.has(id)) teachersBySection.set(id, []);
        teachersBySection.get(id).push(r);

        if (!map.has(id)) {
            map.set(id, {
                id,
                dde: cleanKey(r['Δ.Δ.Ε.']) || 'ΧΩΡΙΣ ΔΔΕ',
                epal: cleanKey(r['ΕΠΑΛ']) || 'ΧΩΡΙΣ ΕΠΑΛ',
                day: '',
                eid: cleanKey(r['Ειδικότητα']) || 'ΧΩΡΙΣ ΕΙΔΙΚΟΤΗΤΑ',
                egkrisi: ''
            });
        }
    });

    // 3. Αρχείο Μαθητευομένων
    data3.forEach(r => {
        const id = cleanKey(r['Α/Α ΤΜΗΜΑΤΟΣ']);
        if (!id) return;

        if (!studentsBySection.has(id)) studentsBySection.set(id, []);
        studentsBySection.get(id).push(r);

        if (!map.has(id)) {
            map.set(id, {
                id,
                dde: cleanKey(r['ΔΔΕ']) || 'ΧΩΡΙΣ ΔΔΕ',
                epal: cleanKey(r['ΕΠΑΛ']) || 'ΧΩΡΙΣ ΕΠΑΛ',
                day: '',
                eid: cleanKey(r['ΕΙΔΙΚΟΤΗΤΑ']) || 'ΧΩΡΙΣ ΕΙΔΙΚΟΤΗΤΑ',
                egkrisi: ''
            });
        }
    });

    allSections = Array.from(map.values()).sort((a, b) =>
        a.dde.localeCompare(b.dde, 'el') ||
        a.epal.localeCompare(b.epal, 'el') ||
        a.eid.localeCompare(b.eid, 'el')
    );

    populateFilterDropdowns();
    applyFilter();
}







function populateFilterDropdowns() {
    const ddeSel = document.getElementById('ddeFilter');
    const prevDde = ddeSel.value;

    const ddes = Array.from(new Set(allSections.map(s => s.dde))).sort((a, b) => a.localeCompare(b, 'el'));
    ddeSel.innerHTML = '<option value="">-- Όλες οι ΔΔΕ --</option>';
    ddes.forEach(d => {
        const opt = document.createElement('option');
        opt.value = d;
        opt.textContent = d;
        ddeSel.appendChild(opt);
    });
    if (prevDde && ddes.includes(prevDde)) ddeSel.value = prevDde;

    updateDependentDropdowns();
}

function updateDependentDropdowns() {
    const selectedDde = document.getElementById('ddeFilter').value;
    const epalSel = document.getElementById('epalFilter');
    const eidSel = document.getElementById('eidFilter');
    const prevEpal = epalSel.value;
    const prevEid = eidSel.value;

    const byDde = selectedDde ? allSections.filter(s => s.dde === selectedDde) : allSections;
    const epals = Array.from(new Set(byDde.map(s => s.epal))).sort((a, b) => a.localeCompare(b, 'el'));

    epalSel.innerHTML = '<option value="">-- Όλα τα ΕΠΑΛ --</option>';
    epals.forEach(ep => {
        const opt = document.createElement('option');
        opt.value = ep;
        opt.textContent = ep;
        epalSel.appendChild(opt);
    });
    if (prevEpal && epals.includes(prevEpal)) epalSel.value = prevEpal;

    const activeEpal = epalSel.value;
    const byEpal = activeEpal ? byDde.filter(s => s.epal === activeEpal) : byDde;
    const eids = Array.from(new Set(byEpal.map(s => s.eid))).sort((a, b) => a.localeCompare(b, 'el'));

    eidSel.innerHTML = '<option value="">-- Όλες οι Ειδικότητες --</option>';
    eids.forEach(ei => {
        const opt = document.createElement('option');
        opt.value = ei;
        opt.textContent = ei;
        eidSel.appendChild(opt);
    });
    if (prevEid && eids.includes(prevEid)) eidSel.value = prevEid;
}

function applyFilter() {
    const ddeVal = document.getElementById('ddeFilter').value;
    const epalVal = document.getElementById('epalFilter').value;
    const eidVal = document.getElementById('eidFilter').value;
    const searchVal = normalizeText(document.getElementById('searchInput').value);

    currentFiltered = allSections.filter(s => {
        if (ddeVal && s.dde !== ddeVal) return false;
        if (epalVal && s.epal !== epalVal) return false;
        if (eidVal && s.eid !== eidVal) return false;
        if (searchVal) {
            const haystack = normalizeText(`${s.id} ${s.eid} ${s.epal} ${s.dde} ${s.day}`);
            if (!haystack.includes(searchVal)) return false;
        }
        return true;
    });

    closeModal();
    renderGroupedTable(currentFiltered);
}

document.getElementById('ddeFilter').addEventListener('change', () => {
    updateDependentDropdowns();
    applyFilter();
});

document.getElementById('epalFilter').addEventListener('change', () => {
    updateDependentDropdowns();
    applyFilter();
});

document.getElementById('eidFilter').addEventListener('change', applyFilter);
document.getElementById('searchInput').addEventListener('input', applyFilter);

document.getElementById('resetFiltersBtn').addEventListener('click', () => {
    document.getElementById('ddeFilter').value = '';
    document.getElementById('searchInput').value = '';
    updateDependentDropdowns();
    document.getElementById('epalFilter').value = '';
    updateDependentDropdowns();
    document.getElementById('eidFilter').value = '';
    applyFilter();
});

document.getElementById('exportCsvBtn').addEventListener('click', () => {
    if (currentFiltered.length === 0) {
        alert('Δεν υπάρχουν δεδομένα για εξαγωγή.');
        return;
    }

    const headers = ['ΔΔΕ', 'ΕΠΑΛ', 'Ειδικότητα', 'Τμήμα', 'Μαθητές', 'Αναθέσεις Εκπ/κών', 'Μάθημα'];
    const rows = currentFiltered.map(s => {
        const stuCount = studentsBySection.get(s.id)?.length || 0;
        const teaCount = teachersBySection.get(s.id)?.length || 0;
        return [s.dde, s.epal, s.eid, s.id, stuCount, teaCount, s.day || '']
            .map(v => `"${String(v).replace(/"/g, '""')}"`)
            .join(';');
    });

    const csvContent = '\uFEFF' + [headers.join(';'), ...rows].join('\r\n');
    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `Τμηματα_Μαθητειας_${new Date().toISOString().slice(0, 10)}.csv`;
    a.click();
    URL.revokeObjectURL(url);
});

function renderGroupedTable(list) {
    const card = document.getElementById('allSectionsCard');
    const wrap = document.getElementById('allSectionsTable');

    if (list.length === 0) {
        card.style.display = allSections.length > 0 ? 'block' : 'none';
        wrap.innerHTML = '<div class="empty-msg">Δεν βρέθηκαν τμήματα με αυτά τα κριτήρια αναζήτησης.</div>';
        return;
    }
    card.style.display = 'block';

    const ddeGroups = new Map();
    list.forEach(item => {
        if (!ddeGroups.has(item.dde)) ddeGroups.set(item.dde, new Map());
        const epalMap = ddeGroups.get(item.dde);
        if (!epalMap.has(item.epal)) epalMap.set(item.epal, []);
        epalMap.get(item.epal).push(item);
    });

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

    let ddeIdx = 0;
    let epalIdx = 0;

    ddeGroups.forEach((epalMap, ddeName) => {
        ddeIdx++;
        const ddeKey = 'dde-' + ddeIdx;
        let ddeSections = 0;
        let ddeStudents = 0;
        epalMap.forEach(items => {
            ddeSections += items.length;
            ddeStudents += items.reduce((sum, x) => sum + (studentsBySection.get(x.id)?.length || 0), 0);
        });

        html += `
            <tr class="group-dde" data-dde="${ddeKey}" data-collapsed="false" onclick="toggleDde('${ddeKey}', this)">
                <td colspan="4">
                    <span class="toggle-arrow">▼</span> 🏛️ ${escapeHTML(ddeName)}
                    <span class="group-badge">Τμήματα: ${ddeSections} | Μαθητές: ${ddeStudents}</span>
                </td>
            </tr>
        `;

        epalMap.forEach((items, epalName) => {
            epalIdx++;
            const epalKey = 'epal-' + epalIdx;
            const epalSections = items.length;
            const epalStudents = items.reduce((sum, x) => sum + (studentsBySection.get(x.id)?.length || 0), 0);

            const contact = epalContactMap.get(`${ddeName}||${epalName}`);
            const contactBadges = [];
            if (contact?.phone) contactBadges.push(`<span>📞 ${escapeHTML(contact.phone)}</span>`);
            if (contact?.email) contactBadges.push(`<span>✉️ ${escapeHTML(contact.email)}</span>`);
            const epalContactHtml = contactBadges.length > 0
                ? `<span class="epal-contact-info" onclick="event.stopPropagation()">${contactBadges.join('')}</span>`
                : '';

            html += `
                <tr class="group-epal child-of-${ddeKey}" data-epal="${epalKey}" data-collapsed="false" onclick="toggleEpal('${epalKey}', this)">
                    <td colspan="4">
                        <span class="toggle-arrow">▼</span> 🏫 ${escapeHTML(epalName)}
                        ${epalContactHtml}
                        <span class="group-badge">Τμήματα: ${epalSections} | Μαθητές: ${epalStudents}</span>
                    </td>
                </tr>
            `;

            items.forEach(item => {
                const isSelected = item.id === selectedTmimaId;
                const stuCount = studentsBySection.get(item.id)?.length || 0;

                const aaBtn = `<span class="click-pill aa-pill" onclick="event.stopPropagation(); showTeachers('${escapeHTML(item.id)}')" title="Προβολή αναθέσεων εκπαιδευτικών">${escapeHTML(item.id)}</span>`;
                const seatsBtn = `<span class="click-pill seats-pill ${stuCount === 0 ? 'seats-warn' : ''}" onclick="event.stopPropagation(); showStudents('${escapeHTML(item.id)}')" title="Προβολή μαθητευομένων">${stuCount}</span>`;

                html += `
                    <tr class="tmima-row child-of-${ddeKey} child-of-${epalKey} ${isSelected ? 'active-row' : ''}" data-id="${escapeHTML(item.id)}">
                        <td>${escapeHTML(item.eid)}</td>
                        <td>${aaBtn}</td>
                        <td>${seatsBtn}</td>
                        <td>${escapeHTML(item.day) || '—'}</td>
                    </tr>
                `;
            });
        });
    });

    html += '</tbody></table>';
    wrap.innerHTML = html;
}







window.toggleDde = function(ddeKey, rowEl) {
    const isCollapsed = rowEl.getAttribute('data-collapsed') === 'true';
    const nextState = !isCollapsed;
    rowEl.setAttribute('data-collapsed', String(nextState));
    rowEl.querySelector('.toggle-arrow').textContent = nextState ? '▶' : '▼';

    document.querySelectorAll('.child-of-' + ddeKey).forEach(el => {
        if (nextState) {
            el.style.display = 'none';
        } else if (el.classList.contains('group-epal')) {
            el.style.display = '';
        } else {
            const epalClass = Array.from(el.classList).find(c => c.startsWith('child-of-epal-'));
            const epalKey = epalClass ? epalClass.replace('child-of-', '') : '';
            const epalRow = document.querySelector(`tr.group-epal[data-epal="${epalKey}"]`);
            const epalCollapsed = epalRow && epalRow.getAttribute('data-collapsed') === 'true';
            el.style.display = epalCollapsed ? 'none' : '';
        }
    });
};

window.toggleEpal = function(epalKey, rowEl) {
    const isCollapsed = rowEl.getAttribute('data-collapsed') === 'true';
    const nextState = !isCollapsed;
    rowEl.setAttribute('data-collapsed', String(nextState));
    rowEl.querySelector('.toggle-arrow').textContent = nextState ? '▶' : '▼';

    document.querySelectorAll('.child-of-' + epalKey).forEach(el => {
        el.style.display = nextState ? 'none' : '';
    });
};

window.toggleAllGroups = function(collapse) {
    document.querySelectorAll('tr.group-dde, tr.group-epal').forEach(row => {
        row.setAttribute('data-collapsed', String(collapse));
        const arrow = row.querySelector('.toggle-arrow');
        if (arrow) arrow.textContent = collapse ? '▶' : '▼';
    });
    document.querySelectorAll('tr.group-epal, tr.tmima-row').forEach(row => {
        row.style.display = collapse ? 'none' : '';
    });
};

function openModal(titleText, bodyHtml, tmimaId) {
    selectedTmimaId = tmimaId;
    document.querySelectorAll('tr.tmima-row').forEach(tr => {
        tr.classList.toggle('active-row', tr.getAttribute('data-id') === tmimaId);
    });
    document.getElementById('modalTitle').textContent = titleText;
    document.getElementById('modalContent').innerHTML = bodyHtml;
    document.getElementById('modalOverlay').style.display = 'flex';
}

window.closeModal = function() {
    document.getElementById('modalOverlay').style.display = 'none';
};

window.closeModalOnBg = function(e) {
    if (e.target.id === 'modalOverlay') closeModal();
};

document.addEventListener('keydown', (e) => {
    if (e.key === 'Escape') closeModal();
});

window.showTeachers = function(tmimaId) {
    const sec = allSections.find(s => s.id === tmimaId);
    const subtitle = sec ? ` — ${sec.epal} (${sec.eid})` : '';
    const teachers = teachersBySection.get(tmimaId) || [];

    let bodyHtml = '';
    if (teachers.length > 0) {
        bodyHtml = `
            <div class="table-wrap">
                <table>
                    <thead>
                        <tr>
                            <th>Εκπαιδευτικός</th>
                            <th>Κλάδος</th>
                            <th>Ιδιότητα</th>
                            <th>ΑΜ</th>
                            <th>Αρ. Πρακτικού</th>
                            <th>Ημερομηνία</th>
                            <th>Ώρες</th>
                            <th>Τύπος</th>
                            <th>Έγκριση</th>
                        </tr>
                    </thead>
                    <tbody>
                        ${teachers.map(t => {
                            const rawAm = cleanKey(t['ΑΜ']);
                            const info = teachersInfoByAm.get(normalizeAm(rawAm));
                            const contactParts = [];
                            if (info?.phone) contactParts.push(`<span>📱 ${escapeHTML(info.phone)}</span>`);
                            if (info?.email) contactParts.push(`<span>✉️ ${escapeHTML(info.email)}</span>`);
                            const contactHtml = contactParts.length > 0
                                ? `<div class="teacher-contact">${contactParts.join('')}</div>`
                                : '';

                            return `
                                <tr>
                                    <td>
                                        <strong>${escapeHTML(t['Εκπαιδευτικός'])}</strong>
                                        ${contactHtml}
                                    </td>
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
            </div>`;
    } else {
        bodyHtml = '<div class="empty-msg">Δεν βρέθηκαν αναθέσεις εκπαιδευτικών για αυτό το τμήμα (ή δεν έχει φορτωθεί το 2ο αρχείο).</div>';
    }

    openModal(`👨‍🏫 Αναθέσεις Εκπαιδευτικών — Τμήμα ${tmimaId} (${teachers.length})${subtitle}`, bodyHtml, tmimaId);
};

window.showStudents = function(tmimaId) {
    const sec = allSections.find(s => s.id === tmimaId);
    const subtitle = sec ? ` — ${sec.epal} (${sec.eid})` : '';
    const students = studentsBySection.get(tmimaId) || [];

    let bodyHtml = '';
    if (students.length > 0) {
        bodyHtml = `
            <div class="table-wrap">
                <table>
                    <thead>
                        <tr>
                            <th>Α/Α</th>
                            <th>Επώνυμο</th>
                            <th>Όνομα</th>
                            <th>Φύλο</th>
                            <th>Κινητό</th>
                            <th>Email</th>
                            <th>Ενεργός</th>
                        </tr>
                    </thead>
                    <tbody>
                        ${students.map((s, idx) => {
                            const stopDate = cleanKey(
                                s['ΗΜΕΡΟΜΗΝΙΑ ΔΙΑΚΟΠΗΣ'] ||
                                s['ΗΜ/ΝΙΑ ΔΙΑΚΟΠΗΣ'] ||
                                s['ΗΜ. ΔΙΑΚΟΠΗΣ'] ||
                                s['ΔΙΑΚΟΠΗ']
                            );
                            const isActive = !stopDate;
                            const gender = cleanKey(s['ΦΥΛΟ'] || s['Φύλο']) || '—';

                            return `
                                <tr>
                                    <td>${idx + 1}</td>
                                    <td><strong>${escapeHTML(s['ΕΠΩΝΥΜΟ'])}</strong></td>
                                    <td>${escapeHTML(s['ΟΝΟΜΑ'])}</td>
                                    <td>${escapeHTML(gender)}</td>
                                    <td>${escapeHTML(s['ΚΙΝΗΤΟ'] || s['ΤΗΛΕΦΩΝΟ']) || '—'}</td>
                                    <td>${escapeHTML(s['EMAIL']) || '—'}</td>
                                    <td>
                                        <span class="badge ${isActive ? 'badge-ok' : 'badge-stop'}" title="${stopDate ? 'Διακοπή: ' + escapeHTML(formatDate(stopDate)) : 'Ενεργός μαθητευόμενος'}">
                                            ${isActive ? 'ΝΑΙ' : 'ΟΧΙ'}
                                        </span>
                                    </td>
                                </tr>
                            `;
                        }).join('')}
                    </tbody>
                </table>
            </div>`;
    } else {
        bodyHtml = '<div class="empty-msg">Δεν βρέθηκαν μαθητευόμενοι για αυτό το τμήμα (ή δεν έχει φορτωθεί το 3ο αρχείο).</div>';
    }

    openModal(`🎓 Μαθητευόμενοι — Τμήμα ${tmimaId} (${students.length})${subtitle}`, bodyHtml, tmimaId);
};


// Αυτόματη ενημέρωση της νέας μπάρας αριθμών (ΔΔΕ, ΕΠΑΛ, Ειδικότητες, Τμήματα, Εκπαιδευτικοί, Μαθητευόμενοι)
const _origRenderGroupedTable = renderGroupedTable;
renderGroupedTable = function(list) {
    const bar = document.getElementById('statsBar');
    if (bar) {
        if (allSections.length === 0) {
            bar.style.display = 'none';
        } else {
            bar.style.display = 'grid';
            const uniqueTeachers = new Set();
            let totalStudents = 0;

            list.forEach(sec => {
                totalStudents += (studentsBySection.get(sec.id)?.length || 0);
                (teachersBySection.get(sec.id) || []).forEach(t => {
                    const key = normalizeAm(t['ΑΜ']) || normalizeText(t['Εκπαιδευτικός']);
                    if (key) uniqueTeachers.add(key);
                });
            });

            document.getElementById('statDde').textContent = new Set(list.map(x => x.dde)).size;
            document.getElementById('statEpal').textContent = new Set(list.map(x => `${x.dde}||${x.epal}`)).size;
            document.getElementById('statEid').textContent = new Set(list.map(x => x.eid)).size;
            document.getElementById('statTmimata').textContent = list.length;
            document.getElementById('statTeachers').textContent = uniqueTeachers.size;
            document.getElementById('statStudents').textContent = totalStudents;
        }
    }
    _origRenderGroupedTable(list);
};









