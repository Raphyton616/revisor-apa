document.addEventListener("DOMContentLoaded", () => {
  const overlay = document.getElementById("upload-overlay");
  const dropZone = document.getElementById("drop-zone");
  const fileInput = document.getElementById("file-input");
  const uploadError = document.getElementById("upload-error");
  const uploadLoading = document.getElementById("upload-loading");
  const analysisContainer = document.getElementById("analysis-container");
  const btnOpenUpload = document.getElementById("btn-open-upload");
  const btnCloseModal = document.getElementById("btn-close-modal");

  let currentFile = null;

  function showOverlay() {
    overlay.classList.remove("closed");
  }

  function hideOverlay() {
    overlay.classList.add("closed");
  }

  function showError(message) {
    uploadError.textContent = message;
    uploadError.classList.remove("hidden");
  }

  function clearError() {
    uploadError.textContent = "";
    uploadError.classList.add("hidden");
  }

  function setLoading(isLoading) {
    uploadLoading.classList.toggle("hidden", !isLoading);
    if (isLoading) {
      uploadLoading.setAttribute("aria-busy", "true");
    } else {
      uploadLoading.removeAttribute("aria-busy");
    }
  }

  async function readError(response, fallback) {
    try {
      const data = await response.json();
      return data.detail || fallback;
    } catch {
      return fallback;
    }
  }

  async function handleFile(file) {
    clearError();

    if (!file || !file.name.toLowerCase().endsWith(".docx")) {
      currentFile = null;
      showError("Solo se admite formato .docx.");
      return;
    }

    if (file.size > 10 * 1024 * 1024) {
      currentFile = null;
      showError("El archivo no puede superar los 10 MB.");
      return;
    }

    currentFile = file;
    setLoading(true);

    const formData = new FormData();
    formData.append("file", file);

    try {
      const response = await fetch("/api/analyze", {
        method: "POST",
        body: formData,
      });

      if (!response.ok) {
        throw new Error(await readError(response, "No fue posible analizar el archivo."));
      }

      const data = await response.json();
      renderAnalysis(data);
      hideOverlay();
    } catch (error) {
      currentFile = null;
      showError(error.message || "No fue posible analizar el archivo.");
    } finally {
      setLoading(false);
    }
  }

  function renderAnalysis(data) {
    analysisContainer.classList.remove("hidden");

    const safeNumber = (value) => Number(value || 0).toLocaleString("es-ES");
    const checks = Array.isArray(data.checks) ? data.checks : [];
    const categories = ["Formato", "Citas", "Referencias"];
    const iconMap = { ok: "✓", warning: "!", error: "×" };

    const headerHtml = `
      <div class="results-header">
        <div>
          <h2>Informe de Verificación APA 7</h2>
          <p class="document-name">
            Documento: <strong>${escapeHtml(currentFile?.name || "")}</strong>
          </p>
        </div>
        <button id="btn-correct-doc" class="btn-primary" type="button">
          Corregir documento
        </button>
      </div>
    `;

    const metricsHtml = `
      <div class="metrics-grid">
        <div class="metric-card">
          <div class="value">${safeNumber(data.ok_count)}</div>
          <div class="label">Correctos</div>
        </div>
        <div class="metric-card">
          <div class="value">${safeNumber(data.issue_count)}</div>
          <div class="label">Para revisar</div>
        </div>
        <div class="metric-card">
          <div class="value">${safeNumber(data.word_count)}</div>
          <div class="label">Palabras</div>
        </div>
        <div class="metric-card">
          <div class="value">${safeNumber(data.reference_count)}</div>
          <div class="label">Referencias</div>
        </div>
      </div>
    `;

    let checksHtml = "";
    categories.forEach((category) => {
      const categoryChecks = checks.filter((check) => check.category === category);
      if (!categoryChecks.length) return;

      checksHtml += `<section class="check-section"><h3>${escapeHtml(category)}</h3>`;
      categoryChecks.forEach((check) => {
        const status = ["ok", "warning", "error"].includes(check.status)
          ? check.status
          : "warning";
        const recommendation = check.recommendation && status !== "ok"
          ? `<div class="check-recommendation">Recomendación: ${escapeHtml(check.recommendation)}</div>`
          : "";

        checksHtml += `
          <div class="check-card check-${status}">
            <div class="check-icon" aria-hidden="true">${iconMap[status]}</div>
            <div>
              <div class="check-title">${escapeHtml(check.title)}</div>
              <div class="check-detail">${escapeHtml(check.detail)}</div>
              ${recommendation}
            </div>
          </div>
        `;
      });
      checksHtml += "</section>";
    });

    const missing = Array.isArray(data.citations_without_reference)
      ? data.citations_without_reference
      : [];
    const missingRefsHtml = missing.length
      ? `
        <section class="check-section">
          <h3>Citas sin referencia coincidente</h3>
          <div class="check-card check-warning">
            <div class="check-icon" aria-hidden="true">!</div>
            <div>
              <div class="check-title">Estas citas no tienen una entrada coincidente</div>
              <ul class="detail-list">
                ${missing.map((citation) => `
                  <li>
                    <strong>${escapeHtml(citation.autor)}</strong>
                    (${escapeHtml(citation.anio)}) —
                    <em>${escapeHtml(citation.tipo)}</em>
                  </li>
                `).join("")}
              </ul>
            </div>
          </div>
        </section>
      `
      : "";

    const uncited = Array.isArray(data.uncited_references)
      ? data.uncited_references
      : [];
    const uncitedHtml = uncited.length
      ? `
        <section class="check-section">
          <h3>Referencias no citadas en el texto</h3>
          <div class="check-card check-warning">
            <div class="check-icon" aria-hidden="true">!</div>
            <div>
              <div class="check-title">Estas entradas no aparecen en las citas detectadas</div>
              <ul class="detail-list">
                ${uncited.map((reference) => `
                  <li>
                    <strong>${escapeHtml(reference.autor)}</strong>
                    ${reference.anio ? `(${escapeHtml(reference.anio)})` : ""}
                  </li>
                `).join("")}
              </ul>
            </div>
          </div>
        </section>
      `
      : "";

    analysisContainer.innerHTML = headerHtml + metricsHtml + checksHtml + missingRefsHtml + uncitedHtml;

    document
      .getElementById("btn-correct-doc")
      .addEventListener("click", downloadCorrectedDocument);
  }

  async function downloadCorrectedDocument() {
    if (!currentFile) {
      showError("Primero debes cargar un documento DOCX.");
      showOverlay();
      return;
    }

    const button = document.getElementById("btn-correct-doc");
    button.disabled = true;
    button.textContent = "Generando archivo...";

    const formData = new FormData();
    formData.append("file", currentFile);

    try {
      const response = await fetch("/api/correct", {
        method: "POST",
        body: formData,
      });

      if (!response.ok) {
        throw new Error(await readError(response, "No se pudo generar la versión corregida."));
      }

      const blob = await response.blob();
      const url = window.URL.createObjectURL(blob);
      const link = document.createElement("a");
      link.href = url;
      link.download = currentFile.name.replace(/\.docx$/i, "") + "_APA7_Corregido.docx";
      document.body.appendChild(link);
      link.click();
      link.remove();
      window.URL.revokeObjectURL(url);
    } catch (error) {
      alert("Error: " + (error.message || "No se pudo descargar el documento."));
    } finally {
      button.disabled = false;
      button.textContent = "Corregir documento";
    }
  }

  function escapeHtml(value) {
    return String(value ?? "")
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#039;");
  }

  btnOpenUpload.addEventListener("click", () => {
    clearError();
    showOverlay();
  });

  btnCloseModal.addEventListener("click", hideOverlay);

  dropZone.addEventListener("click", () => fileInput.click());

  dropZone.addEventListener("dragover", (event) => {
    event.preventDefault();
    dropZone.classList.add("dragover");
  });

  dropZone.addEventListener("dragleave", () => {
    dropZone.classList.remove("dragover");
  });

  dropZone.addEventListener("drop", (event) => {
    event.preventDefault();
    dropZone.classList.remove("dragover");
    if (event.dataTransfer.files.length) {
      handleFile(event.dataTransfer.files[0]);
    }
  });

  fileInput.addEventListener("change", () => {
    if (fileInput.files.length) {
      handleFile(fileInput.files[0]);
    }
  });

  document.addEventListener("keydown", (event) => {
    if (event.key === "Escape") {
      hideOverlay();
    }
  });
});
