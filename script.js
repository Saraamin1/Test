(function () {
    // ---------- Single-select checkbox groups (status summary) ----------
    var statusBoxes = document.querySelectorAll('.status-summary .checkbox-large');
    statusBoxes.forEach(function (box) {
        box.addEventListener('click', function () {
            var wasChecked = box.classList.contains('checked');
            statusBoxes.forEach(function (b) { b.classList.remove('checked'); });
            if (!wasChecked) box.classList.add('checked');
        });
    });

    // ---------- Single-select checkbox groups (OK / Issue / N/A per row) ----------
    document.querySelectorAll('.checkbox-container').forEach(function (container) {
        var boxes = container.querySelectorAll('.checkbox-small');
        boxes.forEach(function (box) {
            box.addEventListener('click', function () {
                var wasChecked = box.classList.contains('checked');
                boxes.forEach(function (b) { b.classList.remove('checked'); });
                if (!wasChecked) box.classList.add('checked');
            });
        });
    });

    // ---------- Reset form ----------
    document.getElementById('resetBtn').addEventListener('click', function () {
        if (!confirm('Clear all entries and selections?')) return;
        document.querySelectorAll('#printable input, #printable textarea').forEach(function (el) { el.value = ''; });
        document.querySelectorAll('#printable .checked').forEach(function (el) { el.classList.remove('checked'); });
    });

    // ---------- Push live values into attributes/content so cloneNode() carries them ----------
    // Inputs only reflect typed/picked values in their live "value" property, not in the
    // HTML attribute, and cloneNode() only copies attributes/content — so we sync first.
    function syncValuesForExport(root) {
        root.querySelectorAll('input[type="text"], input[type="date"], input[type="time"], input[type="datetime-local"]')
            .forEach(function (input) { input.setAttribute('value', input.value); });
        root.querySelectorAll('textarea').forEach(function (ta) { ta.textContent = ta.value; });
    }

    // ---------- Format date / time / datetime-local values as readable text ----------
    function formatDateInput(input) {
        if (!input.value) return '______________________';

        if (input.type === 'date') {
            var d = new Date(input.value + 'T00:00:00');
            return d.toLocaleDateString('en-US', { year: 'numeric', month: '2-digit', day: '2-digit' });
        }

        if (input.type === 'time') {
            var t = input.value.split(':');
            var dt = new Date();
            dt.setHours(parseInt(t[0], 10), parseInt(t[1], 10), 0, 0);
            return dt.toLocaleTimeString('en-US', { hour: '2-digit', minute: '2-digit' });
        }

        if (input.type === 'datetime-local') {
            var dtl = new Date(input.value);
            var datePart = dtl.toLocaleDateString('en-US', { year: 'numeric', month: '2-digit', day: '2-digit' });
            var timePart = dtl.toLocaleTimeString('en-US', { hour: '2-digit', minute: '2-digit' });
            return datePart + '  ' + timePart;
        }

        return input.value;
    }

    // ---------- Build a static, export-ready clone of the form ----------
    // Shared by both PDF and Word export so the two stay visually in sync.
    function buildExportClone() {
        var source = document.getElementById('printable');
        syncValuesForExport(source);

        var clone = source.cloneNode(true);
        clone.id = 'printable';

        // Plain text inputs -> span
        clone.querySelectorAll('input[type="text"]').forEach(function (input) {
            var span = document.createElement('span');
            span.className = 'info-input';
            span.textContent = input.value && input.value.trim() !== '' ? input.value : '______________________';
            input.parentNode.replaceChild(span, input);
        });

        // Date / time / datetime-local inputs -> formatted span
        clone.querySelectorAll('input[type="date"], input[type="time"], input[type="datetime-local"]').forEach(function (input) {
            var span = document.createElement('span');
            span.className = 'info-input';
            span.textContent = formatDateInput(input);
            input.parentNode.replaceChild(span, input);
        });

        // Textareas -> div (keeps the same visual box via the shared class)
        clone.querySelectorAll('textarea').forEach(function (ta) {
            var div = document.createElement('div');
            div.className = 'notes-area';
            div.textContent = ta.value && ta.value.trim() !== '' ? ta.value : ' ';
            ta.parentNode.replaceChild(div, ta);
        });

        // Keep only the glyph on checked boxes; empty out unchecked ones
        clone.querySelectorAll('.checkbox-large, .checkbox-small').forEach(function (box) {
            if (!box.classList.contains('checked')) box.textContent = '';
        });

        // Defensive: #printable never contains the toolbar, but strip it if present
        clone.querySelectorAll('.toolbar').forEach(function (t) { t.remove(); });

        return clone;
    }

    // ---------- PDF download (renders the export clone, not the live form) ----------
    document.getElementById('downloadPdfBtn').addEventListener('click', function () {
        var btn = this;
        btn.disabled = true;
        btn.textContent = 'Preparing PDF…';

        var liveEl = document.getElementById('printable');
        var clone = buildExportClone();

        // Render off-screen at the same width as the live form
        clone.style.position = 'fixed';
        clone.style.left = '-10000px';
        clone.style.top = '0';
        clone.style.width = liveEl.scrollWidth + 'px';
        clone.style.background = '#ffffff';
        document.body.appendChild(clone);

        var opt = {
            margin: 0.4,
            filename: 'Excavation-Trenching-Safety-Inspection.pdf',
            image: { type: 'jpeg', quality: 0.98 },
            html2canvas: { scale: 2, useCORS: true, windowWidth: clone.scrollWidth },
            jsPDF: { unit: 'in', format: 'letter', orientation: 'portrait' },
            pagebreak: { mode: ['avoid-all', 'css', 'legacy'] }
        };

        html2pdf().set(opt).from(clone).save().then(function () {
            document.body.removeChild(clone);
            btn.disabled = false;
            btn.textContent = '⬇ Download PDF';
        }).catch(function (err) {
            console.error(err);
            if (clone.parentNode) document.body.removeChild(clone);
            btn.disabled = false;
            btn.textContent = '⬇ Download PDF';
            alert('PDF generation failed. Please try again.');
        });
    });

    // ---------- Word (.docx) download — real Word file, same layout as the page ----------
    // Uses the html-docx-js library (loaded in index.html) to convert real HTML+CSS
    // into an actual .docx (OOXML/zip) file — not a renamed .doc.
    document.getElementById('downloadWordBtn').addEventListener('click', async function () {
        var btn = this;
        btn.disabled = true;
        var originalLabel = btn.textContent;
        btn.textContent = 'Preparing Word file…';

        try {
            if (!window.htmlDocx) {
                throw new Error('html-docx-js library not loaded');
            }

            var clone = buildExportClone();

            // Pull the real styling straight from style.css so the Word file matches the page
            var cssText = '';
            try {
                var cssResponse = await fetch('style.css');
                if (cssResponse.ok) cssText = await cssResponse.text();
            } catch (fetchErr) {
                console.warn('Could not fetch style.css, falling back to minimal styling.', fetchErr);
            }

            if (!cssText) {
                cssText =
                    'body{font-family:Calibri,Arial,sans-serif;font-size:11pt;color:#343A40;}' +
                    'h1{color:#FF6B35;} h2{color:#fff;background:#1E3A5F;padding:8pt;}' +
                    '.checkbox-large,.checkbox-small{border:2px solid #000;display:inline-block;text-align:center;}';
            }

            // Word doesn't support CSS grid/flexbox, so patch just the layout properties
            // on top of the real style.css (colors/fonts/borders all still come from it).
            var wordLayoutPatch =
                '.header-info,.status-summary,.atmospheric-grid,.signature-section{display:block;width:100%;}' +
                '.info-field,.status-item,.signature-box{display:inline-block;vertical-align:top;width:47%;margin:4pt 1%;}' +
                '.date-time-row{display:block;} .date-time-row input, .date-time-row span{display:inline-block;width:47%;margin-right:2%;}' +
                '.checklist-item{display:block;border-bottom:1px solid #E9ECEF;padding:8pt 0;}' +
                '.checkbox-container{display:inline-block;width:38%;vertical-align:top;}' +
                '.item-text{display:inline-block;width:58%;vertical-align:top;}' +
                '.checkbox-option{display:inline-block;margin-right:10pt;}' +
                '.footer-grid{display:block;width:100%;} .footer-grid div{display:inline-block;width:32%;}';

            var fullHtml =
                '<!DOCTYPE html><html><head><meta charset="utf-8">' +
                '<title>Excavation & Trenching Safety Inspection</title>' +
                '<style>' + cssText + wordLayoutPatch + '</style>' +
                '</head><body>' + clone.outerHTML + '</body></html>';

            var docxBlob = window.htmlDocx.asBlob(fullHtml);

            var link = document.createElement('a');
            link.href = URL.createObjectURL(docxBlob);
            link.download = 'Excavation-Trenching-Safety-Inspection.docx';
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
            URL.revokeObjectURL(link.href);
        } catch (err) {
            console.error(err);
            alert('Word file generation failed: ' + err.message);
        } finally {
            btn.disabled = false;
            btn.textContent = originalLabel;
        }
    });
})();
