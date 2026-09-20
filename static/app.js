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

  btnOpenUpload.addEventListener("click", showOverlay);
  btnCloseModal.addEventListener("click", hideOverlay);

  // Drag and Drop
  dropZone.addEventListener("click", () => fileInput.click());

  dropZone.addEventListener("dragover", (e) => {
    e.preventDefault();
    dropZone.classList.add("dragover");
  });

  dropZone.addEventListener("dragleave", () => {
    dropZone.classList.remove("dragover");
  });

  dropZone.addEventListener("drop", (e) => {
    e.preventDefault();
    dropZone.classList.remove("dragover");
    if (e.dataTransfer.files.length) {
      handleFile(e.dataTransfer.files[0]);
    }
  });

  fileInput.addEventListener("change", () => {
    if (fileInput.files.length) {
      handleFile(fileInput.files[0]);
    }
  });

  function showError(msg) {
    uploadError.textContent = msg;
    uploadError.classList.remove("hidden");
  }

  function clearError() {
    uploadError.textContent = "";
    uploadError.classList.add("hidden");
  }

  async function handleFile(file) {
    clearError();

    if (!file.name.toLowerCase().endsWith(".docx")) {
      showError("¡Solo se admite formato .docx! ⚠️");
      return;
    }

    currentFile = file;
    uploadLoading.classList.remove("hidden");

    const formData = new FormData();
    formData.append("file", file);

    try {
      const response = await fetch("/api/analyze", {
        method: "POST",
        body: formData,
      });

      const data = await response.json();

      if (!response.ok) {
        throw new Error(data.detail || "Error al procesar el archivo.");
      }

      renderAnalysis(data);
      uploadLoading.classList.add("hidden");
      hideOverlay(); // Cierre animado hacia la izquierda
    } catch (err) {
      uploadLoading.classList.add("hidden");
      showError(err.message);
    }
  }

  function renderAnalysis(data) {
    analysisContainer.classList.remove("hidden");
    analysisContainer.innerHTML = "";

    // Header del Análisis
    const headerHtml = `
      <div class="results-header">
        <div>
          <h2>Informe de Verificación APA 7</h2>
          <p style="color: var(--ink-soft); font-size: 0.9rem;">Documento: <strong>${escapeHtml(currentFile.name)}</strong></p>
        </div>
        <button id="btn-correct-doc" class="btn-primary">✨ Corregir documento</button>
      </div>
    `;

    // Métricas
    const metricsHtml = `
      <div class="metrics-grid">
        <div class="metric-card">
          <div class="value">${data.ok_count}</div>
          <div class="label">Correctos</div>
        </div>
        <div class="metric-card">
          <div class="value">${data.issue_count}</div>
          <div class="label">Para revisar</div>
        </div>
        <div class="metric-card">
          <div class="value">${data.word_count.toLocaleString()}</div>
          <div class="label">Palabras</div>
        </div>
        <div class="metric-card">
          <div class="value">${data.reference_count}</div>
          <div class="label">Referencias</div>
        </div>
      </div>
    `;

    // Categorías de checks
    const categories = ["Formato", "Citas", "Referencias"];
    let checksHtml = "";

    categories.forEach((cat) => {
      const categoryChecks = data.checks.filter((c) => c.category === cat);
      if (categoryChecks.length === 0) return;

      checksHtml += `<div class="check-section"><h3>${cat}</h3>`;
      categoryChecks.forEach((c) => {
        const iconMap = { ok: "✓", warning: "!", error: "×" };
        checksHtml += `
          <div class="check-card check-${c.status}">
            <div class="check-icon">${iconMap[c.status] || "!"}</div>
            <div>
              <div class="check-title">${escapeHtml(c.title)}</div>
              <div class="check-detail">${escapeHtml(c.detail)}</div>
              ${c.recommendation && c.status !== "ok" ? `<div class="check-recommendation">Recomendación: ${escapeHtml(c.recommendation)}</div>` : ""}
            </div>
          </div>
        `;
      });
      checksHtml += `</div>`;
    });

    analysisContainer.innerHTML = headerHtml + metricsHtml + checksHtml;

    document.getElementById("btn-correct-doc").addEventListener("click", downloadCorrectedDocument);
  }

  async function downloadCorrectedDocument() {
    if (!currentFile) return;

    const btn = document.getElementById("btn-correct-doc");
    btn.disabled = true;
    btn.textContent = "Generando archivo...";

    const formData = new FormData();
    formData.append("file", currentFile);

    try {
      const response = await fetch("/api/correct", {
        method: "POST",
        body: formData,
      });

      if (!response.ok) {
        const errData = await response.json();
        throw new Error(errData.detail || "No se pudo generar la versión corregida.");
      }

      const blob = await response.blob();
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = currentFile.name.replace(".docx", "") + "_APA7_Corregido.docx";
      document.body.appendChild(a);
      a.click();
      a.remove();
      window.URL.revokeObjectURL(url);

      btn.disabled = false;
      btn.textContent = "✨ Corregir documento";
    } catch (err) {
      alert("Error: " + err.message);
      btn.disabled = false;
      btn.textContent = "✨ Corregir documento";
    }
  }

  function escapeHtml(str) {
    return str.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
  }
});
