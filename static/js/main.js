/* ══════════════════════════════════════════════════════════════
   ProTest — main.js
   Brilliant PDF approach: send rendered HTML to Flask/WeasyPrint
   → server returns a true vector PDF with 100% selectable text.
   Falls back to browser print if WeasyPrint is unavailable.
   ══════════════════════════════════════════════════════════════ */

'use strict';

// ── DOM refs ──────────────────────────────────────────────────
const form         = document.getElementById('config-form');
const pagesContainer = document.getElementById('pages-container');
const emptyPaper   = document.getElementById('empty-paper');
const exportZone   = document.getElementById('export-zone');
const exportBtn    = document.getElementById('export-pdf-btn');
const exportNote   = document.getElementById('export-note');
const errorBanner  = document.getElementById('error-banner');
const srLive       = document.getElementById('sr-live');
const genBtn       = document.getElementById('generate-btn');
const genLabel     = document.getElementById('gen-label');
const genSpinner   = document.getElementById('gen-spinner');
const expLabel     = document.getElementById('exp-label');
const expSpinner   = document.getElementById('exp-spinner');
const logoInput    = document.getElementById('logo-upload');
const chapterSelect = document.getElementById('chapter-select');

// ── State ──────────────────────────────────────────────────────
let currentData   = null;
let uploadedLogo  = null;

// ── Utility ───────────────────────────────────────────────────
const announce = msg => { srLive.textContent = ''; srLive.textContent = msg; };

const showError = msg => {
    errorBanner.textContent = msg;
    errorBanner.hidden = false;
    announce('Error: ' + msg);
};

const clearError = () => { errorBanner.hidden = true; errorBanner.textContent = ''; };

function toRoman(n) {
    const val = [1000,900,500,400,100,90,50,40,10,9,5,4,1];
    const sym = ['m','cm','d','cd','c','xc','l','xl','x','ix','v','iv','i'];
    let out = '';
    val.forEach((v, i) => { while (n >= v) { out += sym[i]; n -= v; } });
    return out;
}

function roundUp5(n) { return Math.ceil(n / 5) * 5; }

function escapeHtml(str) {
    return String(str)
        .replace(/&/g,'&amp;')
        .replace(/</g,'&lt;')
        .replace(/>/g,'&gt;')
        .replace(/"/g,'&quot;');
}

// ── Logo upload ────────────────────────────────────────────────
logoInput.addEventListener('change', e => {
    const file = e.target.files[0];
    if (!file) { uploadedLogo = null; return; }
    const reader = new FileReader();
    reader.onload = ev => {
        uploadedLogo = ev.target.result;
        if (currentData) renderTest(currentData);
    };
    reader.readAsDataURL(file);
});

// ── Real-time settings update ─────────────────────────────────
['academy-name','academy-location','subject-name','test-instructions','font-size','chapter-select']
    .forEach(id => {
        document.getElementById(id).addEventListener('input', () => {
            if (currentData) renderTest(currentData);
        });
    });

// ── Form Submit: Generate Test ─────────────────────────────────
form.addEventListener('submit', async e => {
    e.preventDefault();
    clearError();

    const mcqCount   = Math.max(0, +document.getElementById('mcq-count').value   || 0);
    const shortCount = Math.max(0, +document.getElementById('short-count').value  || 0);
    const longCount  = Math.max(0, +document.getElementById('long-count').value   || 0);
    const chapter    = chapterSelect ? chapterSelect.value : 'All';

    if (mcqCount + shortCount + longCount === 0) {
        showError('Please set at least one question count above zero.');
        return;
    }

    // UI loading state
    genBtn.disabled = true;
    genLabel.textContent = 'Generating…';
    genSpinner.hidden = false;
    announce('Generating test, please wait.');

    try {
        const res = await fetch('/api/generate', {
            method:  'POST',
            headers: { 'Content-Type': 'application/json' },
            body:    JSON.stringify({ mcq_count: mcqCount, short_count: shortCount, long_count: longCount, chapter }),
        });

        if (!res.ok) throw new Error(`Server error ${res.status}`);
        const data = await res.json();

        // Validate we got at least something
        const total = (data.mcqs?.length || 0) + (data.short_questions?.length || 0) + (data.long_questions?.length || 0);
        if (total === 0) throw new Error('No questions returned. The Google Sheet may be empty or unreachable.');

        currentData = data;
        renderTest(data);
        exportZone.hidden = false;
        announce('Test generated successfully.');

    } catch (err) {
        showError(err.message);
        currentData = null;
        exportZone.hidden = true;
    } finally {
        genBtn.disabled = false;
        genLabel.textContent = 'Generate Test';
        genSpinner.hidden = true;
    }
});

// ── Export: WeasyPrint server-side PDF ─────────────────────────
exportBtn.addEventListener('click', async () => {
    if (!currentData) return;

    expLabel.textContent = 'Generating PDF…';
    expSpinner.hidden = false;
    exportBtn.disabled = true;
    exportNote.textContent = 'Server rendering…';
    announce('Generating PDF, please wait.');

    // Build the clean print-ready HTML to send to the server
    const printHtml = buildPrintHtml(currentData);

    try {
        const res = await fetch('/api/export-pdf', {
            method:  'POST',
            headers: { 'Content-Type': 'application/json' },
            body:    JSON.stringify({ html: printHtml }),
        });

        if (res.ok) {
            // WeasyPrint succeeded — download the binary PDF blob
            const blob  = await res.blob();
            const url   = URL.createObjectURL(blob);
            const link  = document.createElement('a');
            link.href   = url;
            link.download = 'ProTest_Assessment.pdf';
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
            URL.revokeObjectURL(url);
            exportNote.textContent = '✓ Vector PDF · Selectable text';
            announce('PDF downloaded successfully.');
        } else {
            // WeasyPrint not available — fall back to browser print
            const errData = await res.json().catch(() => ({}));
            console.warn('WeasyPrint unavailable, falling back to print dialog:', errData.error);
            exportNote.textContent = 'Using print dialog (install WeasyPrint for direct download)';
            await document.fonts.ready;
            window.print();
            announce('Print dialog opened.');
        }
    } catch (err) {
        console.error('Export error:', err);
        // Network-level fallback
        exportNote.textContent = 'Using print dialog as fallback';
        await document.fonts.ready;
        window.print();
    } finally {
        expLabel.textContent = 'Save as PDF';
        expSpinner.hidden = true;
        exportBtn.disabled = false;
    }
});

// ── Build clean print HTML (sent to xhtml2pdf on server) ──────
function buildPrintHtml(data) {
    const body = buildTestBodyHtml(data);
    const fontSize = +(document.getElementById('font-size').value) || 13;

    return `<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<title>ProTest Assessment</title>
<style>
@page { size: A4; margin: 18mm; }
* { box-sizing: border-box; }
body { font-family: Helvetica, Arial, sans-serif; font-size: ${fontSize}px; color: #000; line-height: 1.45; margin: 0; padding: 0; }
p { margin: 0 0 8px 0; padding: 0; }
table { width: 100%; border-collapse: collapse; table-layout: fixed; }
td, th { border: 1.5px solid #000; padding: 6px 9px; vertical-align: top; word-wrap: break-word; overflow-wrap: break-word; }
th { background: #f1f5f9; font-weight: bold; font-size: 0.88em; }
strong, b { font-weight: bold; }
table.no-border, table.no-border td, table.no-border th { border: none !important; padding: 2px 0; background: none; }
table.section-hdr, table.section-hdr td, table.section-hdr th { border: none !important; padding: 0 !important; background: none !important; font-weight: bold; font-size: 1em; }
table.section-hdr { margin: 12px 0 6px; }
.q-block { margin-bottom: 8px; }
</style>
</head>
<body>${buildPdfBodyHtml(data)}</body>
</html>`;
}

// ── Simplified HTML for xhtml2pdf (no flex/grid - uses tables for layout) ──
function buildPdfBodyHtml(data) {
    const subject    = escapeHtml(document.getElementById('subject-name').value     || 'Assessment');
    const academy    = escapeHtml(document.getElementById('academy-name').value     || '');
    const location   = escapeHtml(document.getElementById('academy-location').value || '');
    const instruct   = escapeHtml(document.getElementById('test-instructions').value || '');

    const mcqs   = data.mcqs            || [];
    const shorts = data.short_questions || [];
    const longs  = data.long_questions  || [];

    const today     = new Date();
    const dateStr   = `${String(today.getDate()).padStart(2,'0')}/${String(today.getMonth()+1).padStart(2,'0')}/${today.getFullYear()}`;
    const timeMins  = mcqs.length * 1 + shorts.length * 3 + longs.length * 10;
    const timeStr   = `${roundUp5(timeMins)} Minutes`;
    const totalMarks = mcqs.length * 1 + shorts.length * 2 + longs.length * 5;

    let html = '';

    // Header (table-based for xhtml2pdf compat)
    html += `<table style="width:100%;border:none;margin-bottom:14px;">
        <tr>
            <td style="border:none;width:80px;vertical-align:middle;">${uploadedLogo ? `<img src="${uploadedLogo}" style="max-width:75px;max-height:60px;">` : ''}</td>
            <td style="border:none;text-align:center;vertical-align:middle;">
                <p style="font-size:1.7em;font-weight:bold;text-transform:uppercase;letter-spacing:0.05em;margin:0;">${academy}</p>
                <p style="font-size:1em;color:#334155;margin:2px 0 0;">${location}</p>
            </td>
            <td style="border:none;width:80px;"></td>
        </tr>
    </table>`;

    // Info table
    html += `<table style="margin-bottom:14px;">
        <tr>
            <td style="width:34%;"><strong>Name:</strong></td>
            <td style="width:33%;"><strong>Roll No:</strong></td>
            <td style="width:33%;"><strong>Subject:</strong> ${subject}</td>
        </tr>
        <tr>
            <td><strong>Date:</strong> ${dateStr}</td>
            <td><strong>Time:</strong> ${timeStr}</td>
            <td><strong>Total Marks:</strong> ${totalMarks}</td>
        </tr>
    </table>`;

    if (instruct) {
        html += `<p style="font-style:italic;font-size:0.9em;margin-bottom:16px;border-bottom:1px dashed #ccc;padding-bottom:6px;"><strong>Instructions:</strong> ${instruct}</p>`;
    }

    // Q1 MCQs
    if (mcqs.length > 0) {
        html += `<table class="section-hdr"><tr><td>Question 1: Attempt all Multiple Choice Questions.</td><td style="text-align:right;white-space:nowrap;width:1%;">(${mcqs.length} Marks)</td></tr></table>`;
        html += `<table style="margin-bottom:20px;page-break-inside:auto;">
            <thead><tr>
                <th style="width:34%;">Question</th>
                <th style="width:16.5%;">A</th>
                <th style="width:16.5%;">B</th>
                <th style="width:16.5%;">C</th>
                <th style="width:16.5%;">D</th>
            </tr></thead>
            <tbody>`;
        mcqs.forEach((q, i) => {
            html += `<tr style="page-break-inside:avoid;">
                <td><strong>${i+1}.</strong> ${escapeHtml(q.Question || '')}</td>
                <td>${escapeHtml(q['Option A'] || '')}</td>
                <td>${escapeHtml(q['Option B'] || '')}</td>
                <td>${escapeHtml(q['Option C'] || '')}</td>
                <td>${escapeHtml(q['Option D'] || '')}</td>
            </tr>`;
        });
        html += `</tbody></table>`;
    }

    // Q2 Short
    if (shorts.length > 0) {
        html += `<table class="section-hdr"><tr><td>Question 2: Provide concise answers to the following short questions.</td><td style="text-align:right;white-space:nowrap;width:1%;">(${shorts.length * 2} Marks)</td></tr></table>`;
        shorts.forEach((q, i) => {
            html += `<p class="q-block"><strong>${toRoman(i+1)}.</strong> ${escapeHtml(q.Question || '')}</p>`;
        });
    }

    // Q3 Long
    if (longs.length > 0) {
        html += `<table class="section-hdr"><tr><td>Question 3: Provide detailed answers for the following questions.</td><td style="text-align:right;white-space:nowrap;width:1%;">(${longs.length * 5} Marks)</td></tr></table>`;
        longs.forEach((q, i) => {
            html += `<p class="q-block"><strong>${i+1}.</strong> ${escapeHtml(q['Extended Response Task'] || q.Question || '')}</p>`;
        });
    }

    // Answer Key — always on its own page
    if (mcqs.length > 0) {
        html += `<div style="page-break-before:always;padding-top:4px;">`;
        html += `<table class="no-border" style="margin-bottom:16px;"><tr>
            <td style="border:none;text-align:center;">
                <p style="font-size:1.4em;font-weight:bold;text-transform:uppercase;letter-spacing:0.04em;margin:0;">${academy}</p>
                <p style="font-size:0.9em;color:#475569;margin:2px 0 0;">Answer Key &mdash; ${subject}</p>
            </td>
        </tr></table>`;
        html += `<table class="section-hdr"><tr><td>Question 1: Answer Key (MCQs)</td></tr></table>`;
        html += `<p>`;
        mcqs.forEach((q, i) => {
            let ans = (q['Correct Answer'] || '').replace(/option\s*/i, '').trim();
            html += `<span style="display:inline-block;border:1.5px solid #94a3b8;border-radius:3px;padding:4px 8px;margin:2px;background:#f8fafc;font-size:0.85em;font-weight:500;min-width:48px;text-align:center;"><strong>${i+1}.</strong> ${escapeHtml(ans)}</span>`;
        });
        html += `</p></div>`;
    }

    return html;
}

// ── Render Test: multi-page PDF viewer style ───────────────────
function renderTest(data) {
    // Hide empty state placeholder
    if (emptyPaper) emptyPaper.hidden = true;

    const fontSize   = +(document.getElementById('font-size').value) || 13;
    const subject    = escapeHtml(document.getElementById('subject-name').value     || 'Assessment');
    const academy    = escapeHtml(document.getElementById('academy-name').value     || '');
    const location   = escapeHtml(document.getElementById('academy-location').value || '');
    const instruct   = escapeHtml(document.getElementById('test-instructions').value || '');

    const mcqs   = data.mcqs            || [];
    const shorts = data.short_questions || [];
    const longs  = data.long_questions  || [];

    const today      = new Date();
    const dateStr    = `${String(today.getDate()).padStart(2,'0')}/${String(today.getMonth()+1).padStart(2,'0')}/${today.getFullYear()}`;
    const timeMins   = mcqs.length + shorts.length * 3 + longs.length * 10;
    const timeStr    = `${roundUp5(timeMins)} Minutes`;
    const totalMarks = mcqs.length + shorts.length * 2 + longs.length * 5;

    // Remove previously rendered pages (but keep the emptyPaper node)
    const oldPages = pagesContainer.querySelectorAll('.paper-page');
    oldPages.forEach(p => p.remove());

    /* ══ Build page content blocks ══════════════════════════════
       We use a "virtual page" approach:
       - Page 1: Header + info table + Q1 + Q2 + Q3
       - Last page: Answer Key (always separate)
    ═══════════════════════════════════════════════════════════ */

    // ── Header HTML (reused on every question page) ─────────────
    const headerHtml = `
        <div style="text-align:center;margin-bottom:14px;">
            <div style="position:relative;display:flex;justify-content:center;align-items:center;min-height:${uploadedLogo ? '70px' : '0'};">
                ${uploadedLogo ? `<img src="${uploadedLogo}" style="position:absolute;left:0;max-width:80px;max-height:65px;object-fit:contain;" alt="Academy Logo">` : ''}
                <div>
                    <div style="font-size:1.65em;font-weight:800;text-transform:uppercase;letter-spacing:0.06em;color:#000;">${academy}</div>
                    <div style="font-size:0.95em;color:#334155;margin-top:2px;">${location}</div>
                </div>
            </div>
        </div>
        <table style="width:100%;border-collapse:collapse;margin-bottom:12px;table-layout:fixed;">
            <tr>
                <td style="border:1.5px solid #000;padding:6px 10px;width:34%;"><strong>Name:</strong></td>
                <td style="border:1.5px solid #000;padding:6px 10px;width:33%;"><strong>Roll No:</strong></td>
                <td style="border:1.5px solid #000;padding:6px 10px;width:33%;"><strong>Subject:</strong> ${subject}</td>
            </tr>
            <tr>
                <td style="border:1.5px solid #000;padding:6px 10px;"><strong>Date:</strong> ${dateStr}</td>
                <td style="border:1.5px solid #000;padding:6px 10px;"><strong>Time:</strong> ${timeStr}</td>
                <td style="border:1.5px solid #000;padding:6px 10px;"><strong>Total Marks:</strong> ${totalMarks}</td>
            </tr>
        </table>
        ${instruct ? `<div style="font-style:italic;font-size:0.88em;color:#1e293b;margin-bottom:14px;padding-bottom:5px;border-bottom:1px dashed #cbd5e1;"><strong>Instructions:</strong> ${instruct}</div>` : ''}
    `;

    // ── Q1 MCQ HTML ──────────────────────────────────────────────
    let q1Html = '';
    if (mcqs.length > 0) {
        q1Html += `<div style="margin-bottom:18px;">
            <div style="font-size:1em;font-weight:700;color:#000;margin-bottom:10px;">
                Question 1: Attempt all Multiple Choice Questions.
                <span style="float:right;">(${mcqs.length} Marks)</span>
            </div>
            <table style="width:100%;border-collapse:collapse;table-layout:fixed;word-break:break-word;">
                <thead><tr>
                    <th style="border:1.5px solid #000;padding:6px 7px;background:#f1f5f9;width:34%;text-align:left;font-size:0.87em;">Question</th>
                    <th style="border:1.5px solid #000;padding:6px 7px;background:#f1f5f9;width:16.5%;font-size:0.87em;">A</th>
                    <th style="border:1.5px solid #000;padding:6px 7px;background:#f1f5f9;width:16.5%;font-size:0.87em;">B</th>
                    <th style="border:1.5px solid #000;padding:6px 7px;background:#f1f5f9;width:16.5%;font-size:0.87em;">C</th>
                    <th style="border:1.5px solid #000;padding:6px 7px;background:#f1f5f9;width:16.5%;font-size:0.87em;">D</th>
                </tr></thead>
                <tbody>`;
        mcqs.forEach((q, i) => {
            q1Html += `<tr>
                <td style="border:1.5px solid #000;padding:6px 7px;vertical-align:top;"><strong>${i+1}.</strong> ${escapeHtml(q.Question||'')}</td>
                <td style="border:1.5px solid #000;padding:6px 7px;vertical-align:top;">${escapeHtml(q['Option A']||'')}</td>
                <td style="border:1.5px solid #000;padding:6px 7px;vertical-align:top;">${escapeHtml(q['Option B']||'')}</td>
                <td style="border:1.5px solid #000;padding:6px 7px;vertical-align:top;">${escapeHtml(q['Option C']||'')}</td>
                <td style="border:1.5px solid #000;padding:6px 7px;vertical-align:top;">${escapeHtml(q['Option D']||'')}</td>
            </tr>`;
        });
        q1Html += `</tbody></table></div>`;
    }

    // ── Q2 Short HTML ────────────────────────────────────────────
    let q2Html = '';
    if (shorts.length > 0) {
        q2Html += `<div style="margin-bottom:18px;">
            <div style="font-size:1em;font-weight:700;color:#000;margin-bottom:10px;">
                Question 2: Provide concise answers to the following short questions.
                <span style="float:right;">(${shorts.length * 2} Marks)</span>
            </div>`;
        shorts.forEach((q, i) => {
            q2Html += `<div style="display:flex;gap:8px;margin-bottom:11px;">
                <span style="font-weight:700;min-width:24px;flex-shrink:0;">${toRoman(i+1)}.</span>
                <span style="flex:1;overflow-wrap:break-word;">${escapeHtml(q.Question||'')}</span>
            </div>`;
        });
        q2Html += `</div>`;
    }

    // ── Q3 Long HTML ─────────────────────────────────────────────
    let q3Html = '';
    if (longs.length > 0) {
        q3Html += `<div style="margin-bottom:18px;">
            <div style="font-size:1em;font-weight:700;color:#000;margin-bottom:10px;">
                Question 3: Provide detailed answers for the following questions.
                <span style="float:right;">(${longs.length * 5} Marks)</span>
            </div>`;
        longs.forEach((q, i) => {
            q3Html += `<div style="display:flex;gap:8px;margin-bottom:11px;">
                <span style="font-weight:700;min-width:24px;flex-shrink:0;">${i+1}.</span>
                <span style="flex:1;overflow-wrap:break-word;">${escapeHtml(q['Extended Response Task']||q.Question||'')}</span>
            </div>`;
        });
        q3Html += `</div>`;
    }

    // ── Answer Key HTML ──────────────────────────────────────────
    let akHtml = '';
    if (mcqs.length > 0) {
        akHtml += `
        <div style="text-align:center;margin-bottom:22px;">
            <div style="font-size:1.2em;font-weight:800;color:#000;text-transform:uppercase;letter-spacing:0.04em;">${academy}</div>
            <div style="font-size:0.9em;color:#475569;">Answer Key — ${subject}</div>
        </div>
        <div style="font-size:1em;font-weight:700;color:#000;margin-bottom:14px;">
            Question 1: Answer Key (MCQs)
        </div>
        <div style="display:grid;grid-template-columns:repeat(5,1fr);gap:10px;">`;
        mcqs.forEach((q, i) => {
            let ans = (q['Correct Answer'] || '').replace(/option\s*/i, '').trim();
            akHtml += `<div style="border:1.5px solid #94a3b8;border-radius:4px;padding:6px 8px;background:#f8fafc;font-size:0.88em;font-weight:500;text-align:center;">
                <strong>${i+1}.</strong> ${escapeHtml(ans)}
            </div>`;
        });
        akHtml += `</div>`;
    }

    /* ══ Assemble pages ══════════════════════════════════════════
       Page 1 = header + all questions
       Last page = answer key (always separate)
    ═══════════════════════════════════════════════════════════ */

    // Main question page
    const mainPage = document.createElement('div');
    mainPage.className = 'paper-page';
    mainPage.style.fontSize = fontSize + 'px';
    mainPage.innerHTML = headerHtml + q1Html + q2Html + q3Html;
    pagesContainer.appendChild(mainPage);

    // Answer key page (separate, last)
    if (mcqs.length > 0) {
        const akPage = document.createElement('div');
        akPage.className = 'paper-page';
        akPage.style.fontSize = fontSize + 'px';
        akPage.innerHTML = akHtml;
        pagesContainer.appendChild(akPage);
    }
}

// ── Core HTML Builder for PDF export (shared) ──────────────────
function buildTestBodyHtml(data) {
    // For PDF export, returns the same content but as a flat HTML
    // block suitable for xhtml2pdf (no fixed-height page divs)
    const fontSize   = +(document.getElementById('font-size').value) || 13;
    const subject    = escapeHtml(document.getElementById('subject-name').value     || 'Assessment');
    const academy    = escapeHtml(document.getElementById('academy-name').value     || '');
    const location   = escapeHtml(document.getElementById('academy-location').value || '');
    const instruct   = escapeHtml(document.getElementById('test-instructions').value || '');

    const mcqs   = data.mcqs            || [];
    const shorts = data.short_questions || [];
    const longs  = data.long_questions  || [];

    const today      = new Date();
    const dateStr    = `${String(today.getDate()).padStart(2,'0')}/${String(today.getMonth()+1).padStart(2,'0')}/${today.getFullYear()}`;
    const timeMins   = mcqs.length + shorts.length * 3 + longs.length * 10;
    const timeStr    = `${roundUp5(timeMins)} Minutes`;
    const totalMarks = mcqs.length + shorts.length * 2 + longs.length * 5;

    let html = `<div style="font-family:Helvetica,Arial,sans-serif;font-size:${fontSize}px;color:#0f172a;line-height:1.6;">`;

    // Header
    html += `<div style="text-align:center;margin-bottom:14px;">
        <p style="font-size:1.65em;font-weight:bold;text-transform:uppercase;letter-spacing:0.06em;margin:0;">${academy}</p>
        <p style="font-size:0.95em;color:#334155;margin:2px 0 0;">${location}</p>
    </div>`;

    // Info table
    html += `<table style="width:100%;border-collapse:collapse;margin-bottom:12px;table-layout:fixed;">
        <tr>
            <td style="border:1.5px solid #000;padding:6px 10px;width:34%;"><b>Name:</b></td>
            <td style="border:1.5px solid #000;padding:6px 10px;width:33%;"><b>Roll No:</b></td>
            <td style="border:1.5px solid #000;padding:6px 10px;width:33%;"><b>Subject:</b> ${subject}</td>
        </tr>
        <tr>
            <td style="border:1.5px solid #000;padding:6px 10px;"><b>Date:</b> ${dateStr}</td>
            <td style="border:1.5px solid #000;padding:6px 10px;"><b>Time:</b> ${timeStr}</td>
            <td style="border:1.5px solid #000;padding:6px 10px;"><b>Total Marks:</b> ${totalMarks}</td>
        </tr>
    </table>`;

    if (instruct) {
        html += `<p style="font-style:italic;font-size:0.88em;margin-bottom:14px;border-bottom:1px dashed #ccc;padding-bottom:5px;"><b>Instructions:</b> ${instruct}</p>`;
    }

    // Q1 MCQs
    if (mcqs.length > 0) {
        html += `<p style="font-size:1em;font-weight:bold;margin:0 0 10px;"><b>Question 1: Attempt all Multiple Choice Questions.</b> <span style="float:right;">(${mcqs.length} Marks)</span></p>
        <table style="width:100%;border-collapse:collapse;table-layout:fixed;word-break:break-word;margin-bottom:18px;page-break-inside:auto;">
            <thead><tr>
                <th style="border:1.5px solid #000;padding:6px 7px;background:#f1f5f9;width:34%;font-size:0.87em;text-align:left;">Question</th>
                <th style="border:1.5px solid #000;padding:6px 7px;background:#f1f5f9;width:16.5%;font-size:0.87em;">A</th>
                <th style="border:1.5px solid #000;padding:6px 7px;background:#f1f5f9;width:16.5%;font-size:0.87em;">B</th>
                <th style="border:1.5px solid #000;padding:6px 7px;background:#f1f5f9;width:16.5%;font-size:0.87em;">C</th>
                <th style="border:1.5px solid #000;padding:6px 7px;background:#f1f5f9;width:16.5%;font-size:0.87em;">D</th>
            </tr></thead>
            <tbody>`;
        mcqs.forEach((q, i) => {
            html += `<tr style="page-break-inside:avoid;">
                <td style="border:1.5px solid #000;padding:6px 7px;vertical-align:top;"><b>${i+1}.</b> ${escapeHtml(q.Question||'')}</td>
                <td style="border:1.5px solid #000;padding:6px 7px;vertical-align:top;">${escapeHtml(q['Option A']||'')}</td>
                <td style="border:1.5px solid #000;padding:6px 7px;vertical-align:top;">${escapeHtml(q['Option B']||'')}</td>
                <td style="border:1.5px solid #000;padding:6px 7px;vertical-align:top;">${escapeHtml(q['Option C']||'')}</td>
                <td style="border:1.5px solid #000;padding:6px 7px;vertical-align:top;">${escapeHtml(q['Option D']||'')}</td>
            </tr>`;
        });
        html += `</tbody></table>`;
    }

    // Q2 Short
    if (shorts.length > 0) {
        html += `<p style="font-weight:bold;margin:0 0 10px;"><b>Question 2: Provide concise answers to the following short questions.</b> <span style="float:right;">(${shorts.length * 2} Marks)</span></p>`;
        shorts.forEach((q, i) => {
            html += `<p style="margin:0 0 10px;"><b>${toRoman(i+1)}.</b> ${escapeHtml(q.Question||'')}</p>`;
        });
    }

    // Q3 Long
    if (longs.length > 0) {
        html += `<p style="font-weight:bold;margin:16px 0 10px;"><b>Question 3: Provide detailed answers for the following questions.</b> <span style="float:right;">(${longs.length * 5} Marks)</span></p>`;
        longs.forEach((q, i) => {
            html += `<p style="margin:0 0 10px;"><b>${i+1}.</b> ${escapeHtml(q['Extended Response Task']||q.Question||'')}</p>`;
        });
    }

    // Answer Key on separate page
    if (mcqs.length > 0) {
        html += `<div style="page-break-before:always;padding-top:10px;">
            <p style="font-weight:bold;font-size:1em;margin-bottom:14px;"><b>Question 1: Answer Key (MCQs)</b></p>
            <p>`;
        mcqs.forEach((q, i) => {
            let ans = (q['Correct Answer'] || '').replace(/option\s*/i, '').trim();
            html += `<span style="display:inline-block;border:1.5px solid #94a3b8;border-radius:4px;padding:5px 10px;margin:3px;background:#f8fafc;font-size:0.88em;font-weight:500;min-width:50px;text-align:center;"><b>${i+1}.</b> ${escapeHtml(ans)}</span>`;
        });
        html += `</p></div>`;
    }

    html += `</div>`;
    return html;
}


