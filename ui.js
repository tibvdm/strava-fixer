/* Page behaviour: pick or drop a file, repair it, download the result. */
(function () {
  "use strict";

  const dropzone = document.getElementById("dropzone");
  const input = document.getElementById("file");
  const status = document.getElementById("status");
  let lastResult = null;

  function el(tag, attrs, children) {
    const node = document.createElement(tag);
    Object.assign(node, attrs || {});
    for (const child of [].concat(children || [])) {
      node.append(child);
    }
    return node;
  }

  function show(kind, title, body) {
    status.className = "status " + kind;
    status.replaceChildren(el("h2", { textContent: title }), ...body);
    status.hidden = false;
  }

  function download() {
    XLSX.writeFile(lastResult.workbook, lastResult.fileName, { compression: true });
  }

  function againButton() {
    return el("button", { type: "button", className: "button secondary", textContent: "Nog een bestand herstellen", onclick: () => input.click() });
  }

  function showSuccess(result) {
    const body = [
      el("p", { textContent: `${result.activities} activiteiten hersteld. Het bestand “${result.fileName}” staat in je map Downloads.` }),
    ];
    if (result.problems.length) {
      body.push(el("div", { className: "warning" }, [
        `Let op: ${result.problems.length} activiteit(en) konden niet volledig hersteld worden. Kijk die even na in Excel:`,
        el("ul", {}, result.problems.map((p) => el("li", { textContent: `Rij ${p.row + 1}: ${p.preview}` }))),
      ]));
    }
    body.push(el("div", { className: "actions" }, [
      el("button", { type: "button", className: "button", textContent: "Opnieuw downloaden", onclick: download }),
      againButton(),
    ]));
    show("ok", "Klaar!", body);
  }

  function showError(message) {
    show("error", "Dat lukte niet", [
      el("p", { textContent: message }),
      el("div", { className: "actions" }, [againButton()]),
    ]);
  }

  async function handleFile(file) {
    if (!file) return;
    if (!/\.xlsx$/i.test(file.name)) {
      showError(`“${file.name}” is geen Excel-bestand. Kies een bestand dat eindigt op .xlsx.`);
      return;
    }
    dropzone.classList.add("busy");
    show("", "Bezig met herstellen…", [el("p", { textContent: file.name })]);
    try {
      const data = new Uint8Array(await file.arrayBuffer());
      // Let the browser paint the "busy" state before the (short) blocking work.
      await new Promise((resolve) => setTimeout(resolve, 30));
      const result = StravaFixer.fixWorkbook(XLSX, Papa, data);
      result.fileName = file.name.replace(/\.xlsx$/i, "") + "_fixed.xlsx";
      lastResult = result;
      download();
      showSuccess(result);
    } catch (e) {
      console.error(e);
      showError(e instanceof StravaFixer.FixError ? e.message : "Er ging iets onverwachts mis bij het lezen van dit bestand.");
    } finally {
      dropzone.classList.remove("busy");
      input.value = "";
    }
  }

  input.addEventListener("change", () => handleFile(input.files[0]));

  // Accept drops anywhere on the page, so a slightly missed drop does not open the file in the browser.
  window.addEventListener("dragover", (e) => {
    e.preventDefault();
    dropzone.classList.add("dragging");
  });
  window.addEventListener("dragleave", (e) => {
    if (!e.relatedTarget) dropzone.classList.remove("dragging");
  });
  window.addEventListener("drop", (e) => {
    e.preventDefault();
    dropzone.classList.remove("dragging");
    handleFile(e.dataTransfer.files[0]);
  });
})();
