/* Page behaviour: pick or drop activities.csv, convert it, download the xlsx. */
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
    return el("button", { type: "button", className: "button secondary", textContent: "Ander bestand kiezen", onclick: () => input.click() });
  }

  function showSuccess(result) {
    const body = [
      el("p", { textContent: `${result.activities} activiteiten omgezet. Het bestand “${result.fileName}” staat in je map Downloads.` }),
    ];
    const notes = [
      ...result.problems.map((p) => `Rij ${p.row}: ${p.fields} in plaats van ${result.width} gegevens (${p.preview})`),
      ...result.warnings,
    ];
    if (notes.length) {
      body.push(el("div", { className: "warning" }, [
        "Let op: kijk deze rijen even na in Excel.",
        el("ul", {}, notes.map((text) => el("li", { textContent: text }))),
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
    if (!/\.csv$/i.test(file.name)) {
      showError(StravaFixer.NOT_STRAVA);
      return;
    }
    dropzone.classList.add("busy");
    show("", "Bezig met omzetten…", [el("p", { textContent: file.name })]);
    try {
      const data = new Uint8Array(await file.arrayBuffer());
      // Let the browser paint the "busy" state before the (short) blocking work.
      await new Promise((resolve) => setTimeout(resolve, 30));
      const result = StravaFixer.convertCsv(XLSX, Papa, data);
      result.fileName = file.name.replace(/\.csv$/i, "") + ".xlsx";
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
