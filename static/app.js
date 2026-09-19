(() => {
  const uploadZone = document.getElementById("uploadZone");
  const fileInput = document.getElementById("fileInput");
  const fileInfo = document.getElementById("fileInfo");
  const fileName = document.getElementById("fileName");
  const btnClear = document.getElementById("btnClear");
  const btnAnalyze = document.getElementById("btnAnalyze");
  const btnCorrect = document.getElementById("btnCorrect");
  const emptyState = document.getElementById("emptyState");
  const loading = document.getElementById("loading");
  const results = document.getElementById("results");
  const metrics = document.getElementById("metrics");
  const summaryAlert = document.getElementById("summaryAlert");
  const checksContainer = document.getElementById("checksContainer");
  const detailSection = document.getElementById("detailSection");
  const detailContent = document.getElementById("detailContent");

  let selectedFile = null;

  // --- Upload handling ---
  uploadZone.addEventListener("click", () => fileInput.click());

  uploadZone.addEventListener("dragover", (e) => {
    e.preventDefault();
    uploadZone.classList.add("dragover");
  });

  uploadZone.addEventListener("dragleave", () => {
    uploadZone.classList.remove("dragover");
  });

  uploadZone.addEventListener("drop", (e) => {
    e.preventDefault();
    uploadZone.classList.remove("dragover");
    const file = e.dataTransfer.files[0];
    if (file) handleFile(file);
  });

  fileInput.addEventListener("change", () => {
    if (fileInput.files[0]) handleFile(fileInput.files[0]);
  });

  btnClear.addEventListener("click", clearFile);

  function handleFile(file) {
    if (!file.name.toLowerCase().endsWith(".docx")) {
      alert("Solo se aceptan archivos .docx");
      return;
    }
    selectedFile = file;
    fileName.textContent = file.name;
    fileInfo.style.display = "flex";
    btnAnalyze.disabled = false;
    btnCorrect.disabled = false;
    emptyState.style.display = "block";
    results.style.display = "none";
  }

  function clearFile() {
    selectedFile = null;
    fileInput.value = "";
    fileInfo.style.display = "none";
    btnAnalyze.disabled = true;
    btnCorrect.disabled = true;
    emptyState.style.display = "block";
    results.style.display = "none";
  }

  // --- Analyze ---
  btnAnalyze.addEventListener("click", async () => {
    if (!selectedFile) return;

    emptyState.style.display = "none";
    results.style.display = "none";
    loading.style.display = "block";
    btnAnalyze.disabled = true;
    btnCorrect.disabled = true;

    const formData = new FormData();
    formData.append("file", selectedFile);

    try {
      const response = await fetch("/api/analyze", {
        method: "POST",
        body: formData,
      });

      if (!response.ok) {
        const err = await response.json().catch(() => ({}));
        throw new Error(err.detail || "Error al analizar el documento");
      }

      const data = await response.json();
      renderResults(data);
    } catch (err) {
      alert("Error: " + err.message);
      emptyState.style.display = "block";
    } finally {
      loading.style.display = "none";
      btnAnalyze.disabled = false;
      btnCorrect.disabled = false;
    }
  });

  // --- Correct / Download ---
  btnCorrect.addEventListener("click", async () => {
    if (!selectedFile) return;

    btnCorrect.disabled = true;
    btnCorrect.textContent = "Generando…";

    const formData = new FormData();
    formData.append("file", selectedFile);

    try {
      const response = await fetch("/api/correct", {
        method: "POST",
        body: formData,
      });

      if (!response.ok) {
        const err = await response.json().catch(() => ({}));
        throw new Error(err.detail || "Error al generar el documento corregido");
      }

      const blob = await response.blob();
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = selectedFile.name.replace(/\.docx$/i, "") + "_APA7_corregido.docx";
      document.body.appendChild(a);
      a.click();
      a.remove();
      URL.revokeObjectURL(url);
    } catch (err) {
      alert("Error: " + err.message);
    } finally {
      btnCorrect.disabled = false;
      btnCorrect.textContent = "Descargar corregido";
    }
  });

  // --- Render results ---
  function renderResults(data) {
    results.style.display = "block";

    // Metrics
    metrics.innerHTML = `
      <div class="metric">
        <span class="metric-label">Párrafos</span>
        <span class="metric-value">${data.paragraph_count ?? 0}</span>
      </div>
      <div class="metric">
        <span class="metric-label">Palabras</span>
        <span class="metric-value">${data.word_count ?? 0}</span>
      </div>
      <div class="metric">
        <span class="metric-label">Citas</span>
        <span class="metric-value">${data.citation_count ?? 0}</span>
      </div>
      <div class="metric">
        <span class="metric-label">Referencias</span>
        <span class="metric-value">${data.reference_count ?? 0}</span>
      </div>
    `;

    // Summary alert
    const ok = data.ok_count ?? 0;
    const issues = data.issue_count ?? 0;
    if (issues === 0) {
      summaryAlert.className = "alert success";
      summaryAlert.textContent = `Excelente. Se revisaron ${ok} criterios y no se detectaron problemas importantes.`;
    } else {
      summaryAlert.className = "alert info";
      summaryAlert.textContent = `Se revisaron ${ok + issues} criterios: ${ok} correctos y ${issues} con observaciones.`;
    }

    // Group checks by category
    const categories = {};
    (data.checks || []).forEach((check) => {
      if (!categories[check.category]) categories[check.category] = [];
      categories[check.category].push(check);
    });

    checksContainer.innerHTML = "";
    Object.entries(categories).forEach(([category, checks]) => {
      const section = document.createElement("div");
      section.className = "category";
      section.innerHTML = `<h3>${category}</h3>`;

      checks.forEach((check) => {
        const statusClass =
          check.status === "ok" ? "check-ok" :
          check.status === "warning" ? "check-warning" : "check-error";

        const icon =
          check.status === "ok" ? "✓" :
          check.status === "warning" ? "!" : "✕";

        const card = document.createElement("div");
        card.className = `check-card ${statusClass}`;
        card.innerHTML = `
          <div class="check-icon">${icon}</div>
          <div>
            <div class="check-title">${check.title}</div>
            <div class="check-detail">${check.detail || ""}</div>
            \( {check.recommendation ? `<div class="check-recommendation"> \){check.recommendation}</div>` : ""}
          </div>
        `;
        section.appendChild(card);
      });

      checksContainer.appendChild(section);
    });

    // Detail section (citations & references)
    const hasDetails =
      (data.citations && data.citations.length) ||
      (data.references && data.references.length) ||
      (data.citations_without_reference && data.citations_without_reference.length) ||
      (data.uncited_references && data.uncited_references.length);

    if (hasDetails) {
      detailSection.style.display = "block";
      let html = "";

      if (data.citations && data.citations.length) {
        html += `<div class="detail-block"><h4>Citas detectadas (${data.citations.length})</h4><ul>`;
        data.citations.forEach((c) => {
          html += `<li>\( {c.texto} <em>( \){c.tipo})</em></li>`;
        });
        html += `</ul></div>`;
      }

      if (data.citations_without_reference && data.citations_without_reference.length) {
        html += `<div class="detail-block"><h4>Citas sin referencia coincidente</h4><ul>`;
        data.citations_without_reference.forEach((c) => {
          html += `<li>${c.texto}</li>`;
        });
        html += `</ul></div>`;
      }

      if (data.references && data.references.length) {
        html += `<div class="detail-block"><h4>Referencias detectadas (${data.references.length})</h4><ul>`;
        data.references.forEach((r) => {
          html += `<li>${r.texto}</li>`;
        });
        html += `</ul></div>`;
      }

      if (data.uncited_references && data.uncited_references.length) {
        html += `<div class="detail-block"><h4>Referencias sin cita en el texto</h4><ul>`;
        data.uncited_references.forEach((r) => {
          html += `<li>${r.texto}</li>`;
        });
        html += `</ul></div>`;
      }

      detailContent.innerHTML = html;
    } else {
      detailSection.style.display = "none";
    }
  }
})();
