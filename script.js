(function () {
  // ---------- Single-select checkbox groups (status summary) ----------
  var statusBoxes = document.querySelectorAll(
    ".status-summary .checkbox-large"
  );
  statusBoxes.forEach(function (box) {
    box.addEventListener("click", function () {
      var wasChecked = box.classList.contains("checked");
      statusBoxes.forEach(function (b) {
        b.classList.remove("checked");
      });
      if (!wasChecked) box.classList.add("checked");
    });
  });

  // ---------- Single-select checkbox groups (OK / Issue / N/A per row) ----------
  document
    .querySelectorAll(".checkbox-container")
    .forEach(function (container) {
      var boxes = container.querySelectorAll(".checkbox-small");
      boxes.forEach(function (box) {
        box.addEventListener("click", function () {
          var wasChecked = box.classList.contains("checked");
          boxes.forEach(function (b) {
            b.classList.remove("checked");
          });
          if (!wasChecked) box.classList.add("checked");
        });
      });
    });

  // ---------- Reset form ----------
  document.getElementById("resetBtn").addEventListener("click", function () {
    if (!confirm("Clear all entries and selections?")) return;
    document
      .querySelectorAll('#printable input[type="text"], #printable textarea')
      .forEach(function (el) {
        el.value = "";
      });
    document.querySelectorAll("#printable .checked").forEach(function (el) {
      el.classList.remove("checked");
    });
  });

  // ---------- Sync current input values onto the DOM before export ----------
  // (inputs already reflect their live values automatically; textareas need
  // the value pushed onto the visible textContent for html2canvas/doc export)
  function syncValuesForExport(root) {
    root.querySelectorAll('input[type="text"]').forEach(function (input) {
      input.setAttribute("value", input.value);
    });
    root.querySelectorAll("textarea").forEach(function (ta) {
      ta.textContent = ta.value;
    });
  }

  // ---------- PDF download (keeps the exact visual layout) ----------
  document
    .getElementById("downloadPdfBtn")
    .addEventListener("click", function () {
      var btn = this;
      btn.disabled = true;
      btn.textContent = "Preparing PDF…";

      var el = document.getElementById("printable");
      syncValuesForExport(el);

      var opt = {
        margin: 0.4,
        filename: "Excavation-Trenching-Safety-Inspection.pdf",
        image: { type: "jpeg", quality: 0.98 },
        html2canvas: { scale: 2, useCORS: true, windowWidth: el.scrollWidth },
        jsPDF: { unit: "in", format: "letter", orientation: "portrait" },
        pagebreak: { mode: ["avoid-all", "css", "legacy"] },
      };

      html2pdf()
        .set(opt)
        .from(el)
        .save()
        .then(function () {
          btn.disabled = false;
          btn.textContent = "⬇ Download PDF";
        })
        .catch(function (err) {
          console.error(err);
          btn.disabled = false;
          btn.textContent = "⬇ Download PDF";
          alert("PDF generation failed. Please try again.");
        });
    });

  // ---------- Word (.doc) download — same shape & order as the form ----------
  document
    .getElementById("downloadWordBtn")
    .addEventListener("click", function () {
      var source = document.getElementById("printable");
      var clone = source.cloneNode(true);

      // Replace inputs with plain text so Word renders the entered values
      clone.querySelectorAll('input[type="text"]').forEach(function (input) {
        var span = document.createElement("span");
        span.textContent =
          input.value && input.value.trim() !== ""
            ? input.value
            : "______________________";
        span.style.borderBottom = "1px solid #343A40";
        span.style.display = "inline-block";
        span.style.minWidth = "120px";
        input.parentNode.replaceChild(span, input);
      });

      // Replace textarea with formatted paragraph
      clone.querySelectorAll("textarea").forEach(function (ta) {
        var div = document.createElement("div");
        div.style.border = "1px solid #6C757D";
        div.style.minHeight = "80px";
        div.style.padding = "8pt";
        div.textContent = ta.value && ta.value.trim() !== "" ? ta.value : " ";
        ta.parentNode.replaceChild(div, ta);
      });

      // Render checkbox state as plain characters (Word can't do CSS ::checked tricks reliably)
      clone
        .querySelectorAll(".checkbox-large, .checkbox-small")
        .forEach(function (box) {
          var isChecked = box.classList.contains("checked");
          box.textContent = isChecked ? box.textContent : "";
          box.style.color = "#000000";
          box.style.textAlign = "center";
        });

      var styles =
        "" +
        "body{font-family:Calibri,Arial,sans-serif;font-size:11pt;color:#343A40;}" +
        "h1{color:#FF6B35;font-size:22pt;text-transform:uppercase;}" +
        "h2{color:#FFFFFF;background:#1E3A5F;padding:8pt;font-size:16pt;margin-top:20pt;}" +
        "h3{font-size:13pt;margin-top:16pt;}" +
        ".header{border-bottom:3px solid #FF6B35;padding-bottom:10pt;margin-bottom:16pt;}" +
        ".info-label{font-size:9pt;font-weight:bold;color:#6C757D;display:block;}" +
        ".status-summary{background:#E9ECEF;padding:10pt;margin-bottom:16pt;}" +
        ".status-item{display:inline-block;width:32%;vertical-align:top;margin-bottom:8pt;}" +
        ".checkbox-large{border:2px solid #000;width:30px;height:30px;display:inline-block;text-align:center;}" +
        ".checkbox-small{border:2px solid #000;width:18px;height:18px;display:inline-block;text-align:center;font-size:11pt;}" +
        ".checklist-item{border-bottom:1px solid #E9ECEF;padding:8pt 0;}" +
        ".checkbox-container{display:inline-block;width:35%;vertical-align:top;}" +
        ".checkbox-option{display:inline-block;margin-right:10pt;}" +
        ".item-text{display:inline-block;width:60%;vertical-align:top;}" +
        ".alert-box{background:#F8D7DA;border-left:4px solid #DC143C;padding:8pt;margin:10pt 0;}" +
        ".info-box{background:#D1ECF1;border-left:4px solid #17A2B8;padding:8pt;margin:10pt 0;}" +
        ".atmospheric-grid{background:#E9ECEF;padding:10pt;margin:10pt 0;}" +
        ".signature-section{margin-top:20pt;}" +
        ".signature-box{border:2px dashed #6C757D;padding:10pt;width:45%;display:inline-block;vertical-align:top;margin-right:2%;}" +
        ".footer{margin-top:20pt;border-top:1px solid #6C757D;font-size:8pt;color:#6C757D;text-align:center;}";

      var html =
        '<html xmlns:o="urn:schemas-microsoft-com:office:office" ' +
        'xmlns:w="urn:schemas-microsoft-com:office:word" xmlns="http://www.w3.org/TR/REC-html40">' +
        '<head><meta charset="utf-8"><title>Excavation & Trenching Safety Inspection</title>' +
        "<style>" +
        styles +
        "</style></head><body>" +
        clone.innerHTML +
        "</body></html>";

      var blob = new Blob(["\ufeff", html], { type: "application/msword" });
      var link = document.createElement("a");
      link.href = URL.createObjectURL(blob);
      link.download = "Excavation-Trenching-Safety-Inspection.doc";
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
    });
})();
