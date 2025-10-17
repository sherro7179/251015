(() => {
  const catalogEl = document.getElementById("catalog-data");
  if (!catalogEl) {
    return;
  }

  const catalog = JSON.parse(catalogEl.textContent || "{}");
  const form = document.getElementById("validation-form");
  const docTypeSelect = document.getElementById("doc-type-select");
  const uploadBlock = document.getElementById("upload-block");
  const fileInput = document.getElementById("doc-file");
  const fileNameLabel = document.getElementById("selected-file-name");
  const submitBtn = document.getElementById("submit-btn");
  const resetBtn = document.getElementById("reset-form");

  const guideContainer = document.getElementById("template-guide");
  const guideDocName = document.getElementById("guide-doc-name");
  const guideSummary = document.getElementById("guide-summary");
  const guideFields = document.getElementById("guide-fields");
  const guideAttachments = document.getElementById("guide-attachments");
  const guideNotes = document.getElementById("guide-notes");
  const guideNotesBlock = document.getElementById("guide-notes-block");
  const guideApprovalsBlock = document.getElementById("guide-approvals-block");
  const guideApprovals = document.getElementById("guide-approvals");

  const resultCard = document.getElementById("result-card");
  const resultPlaceholder = document.getElementById("result-placeholder");
  const resultTitle = document.getElementById("result-title");
  const resultSummary = document.getElementById("result-summary");
  const resultType = document.getElementById("result-type");
  const resultFilename = document.getElementById("result-filename");
  const resultDetails = document.getElementById("result-details");
  const statusPill = document.querySelector(".status-pill");

  const docTemplates = catalog.doc_templates || {};
  const docTypes = catalog.doc_types || [];
  const roleLabels = (catalog.roles || []).reduce((acc, item) => {
    acc[item.code] = item.label;
    return acc;
  }, {});

  let currentDocType = "";
  let currentFile = null;

  function renderList(element, items, emptyText = "내용 없음") {
    element.innerHTML = "";
    if (!items || !items.length) {
      const li = document.createElement("li");
      li.textContent = emptyText;
      element.appendChild(li);
      return;
    }
    items.forEach((item) => {
      const li = document.createElement("li");
      li.textContent = item;
      element.appendChild(li);
    });
  }

  function renderFields(element, items) {
    element.innerHTML = "";
    if (!items || !items.length) {
      const li = document.createElement("li");
      li.textContent = "No guidance provided.";
      element.appendChild(li);
      return;
    }
    items.forEach((item) => {
      const li = document.createElement("li");
      li.innerHTML = `<strong>${item.label}</strong> ${item.value}`;
      element.appendChild(li);
    });
  }

  function formatApprovalFlow(flow = []) {
    if (!flow.length) {
      return "Not specified";
    }
    return flow.map((code) => roleLabels[code] || code).join(" → ");
  }

  function renderGuide(template) {
    if (!guideContainer) {
      return;
    }
    if (!template) {
      guideContainer.classList.add("hidden");
      return;
    }
    guideContainer.classList.remove("hidden");
    guideDocName.textContent = template.name || "";
    guideSummary.textContent = template.summary || "";
    renderFields(guideFields, template.fields || []);
    renderList(
      guideAttachments,
      template.attachments || [],
      "No attachment guidance"
    );

    if (template.approval_flow && template.approval_flow.length) {
      guideApprovalsBlock.classList.remove("hidden");
      guideApprovals.textContent = formatApprovalFlow(template.approval_flow);
    } else {
      guideApprovalsBlock.classList.add("hidden");
      guideApprovals.textContent = "";
    }

    if (template.tips && template.tips.length) {
      guideNotesBlock.classList.remove("hidden");
      renderList(guideNotes, template.tips);
    } else {
      guideNotesBlock.classList.add("hidden");
      guideNotes.innerHTML = "";
    }
  }

  function resetFormState() {
    fileInput.value = "";
    fileNameLabel.textContent = "선택된 파일이 없습니다.";
    uploadBlock.classList.add("hidden");
    submitBtn.disabled = true;
    currentFile = null;
  }

  function resetAll() {
    form?.reset();
    docTypeSelect.value = "";
    currentDocType = "";
    resetFormState();
    renderGuide(null);
    resultCard.classList.add("hidden");
    resultDetails.innerHTML = "";
    resultPlaceholder.classList.remove("hidden");
    resultPlaceholder.innerHTML =
      '<h3>검토 대기 중</h3><p>문서 유형을 선택하고 DOCX 파일을 첨부한 뒤 "규정 검토 실행"을 눌러주세요.</p>';
  }

  function handleDocTypeChange() {
    currentDocType = docTypeSelect.value;
    resetFormState();
    if (!currentDocType) {
      renderGuide(null);
      return;
    }
    renderGuide(docTemplates[currentDocType]);
    uploadBlock.classList.remove("hidden");
  }

  function handleFileChange() {
    const [file] = fileInput.files;
    currentFile = file || null;
    fileNameLabel.textContent = currentFile
      ? currentFile.name
      : "선택된 파일이 없습니다.";
    submitBtn.disabled = !(currentDocType && currentFile);
  }

  function setLoading(isLoading) {
    submitBtn.disabled = isLoading || !(currentDocType && currentFile);
    docTypeSelect.disabled = isLoading;
    fileInput.disabled = isLoading;
    if (isLoading) {
      resultPlaceholder.classList.remove("hidden");
      resultPlaceholder.innerHTML =
        "<h3>검토 중...</h3><p>AI 엔진이 양식과 규정을 확인하고 있습니다.</p>";
      resultCard.classList.add("hidden");
    }
  }

  function renderMissing(items = []) {
    if (!items.length) {
      return "모든 필수 항목이 충족되었습니다.";
    }
    return `보완 필요 항목:<br/>${items
      .map((item) => `• ${item}`)
      .join("<br/>")}`;
  }

  function renderResult(payload) {
    resultPlaceholder.classList.add("hidden");
    resultCard.classList.remove("hidden");

    const passed = Boolean(payload?.passed);
    resultCard.classList.toggle("success", passed);
    resultCard.classList.toggle("fail", !passed);
    statusPill.textContent = passed ? "PASS" : "REVIEW";
    resultTitle.textContent = passed ? "결재 진행 가능" : "양식/규정 보완 필요";
    resultSummary.textContent = passed
      ? "업로드한 결재 서류가 템플릿과 규정을 모두 충족합니다."
      : "아래 항목을 확인하고 문서를 보완한 뒤 다시 검토해 주세요.";

    const docTypeInfo = docTypes.find(
      (item) => item.code === payload.doc_type
    );
    resultType.textContent = docTypeInfo ? docTypeInfo.label : payload.doc_type;
    resultFilename.textContent = payload.filename || "-";

    const structure = payload.structure || { ok: false, missing: [], checked: [] };
    const regulation =
      payload.regulation || { ok: false, missing: [], checked: [] };

    const coverage = Math.round((structure.coverage || 0) * 100);

    let structureMessage = "";
    if (!structure.checked || !structure.checked.length) {
      structureMessage = "샘플 양식 정보가 없어 구조 비교를 수행하지 못했습니다.";
    } else if (structure.ok) {
      structureMessage = "양식 구성이 동일합니다.";
    } else {
      structureMessage = renderMissing(structure.missing);
    }

    let regulationMessage = "";
    if (!regulation.checked || !regulation.checked.length) {
      regulationMessage = "규정 비교 정보가 없어 세부 검사를 수행하지 못했습니다.";
    } else if (regulation.ok) {
      regulationMessage = "규정이 요구하는 항목이 확인되었습니다.";
    } else {
      regulationMessage = renderMissing(regulation.missing);
    }

    resultDetails.innerHTML = `
      <div class="detail-item ${structure.ok ? "pass" : "fail"}">
        <strong>양식 구조 비교</strong>
        <span>${structureMessage}</span>
        <small>일치율: ${coverage}%</small>
      </div>
      <div class="detail-item ${regulation.ok ? "pass" : "fail"}">
        <strong>규정 필수 문구/항목</strong>
        <span>${regulationMessage}</span>
      </div>
    `;
  }

  async function submitForm(event) {
    event.preventDefault();
    if (!currentDocType) {
      alert("문서 유형을 선택하세요.");
      return;
    }
    if (!currentFile) {
      alert("검토할 DOCX 파일을 첨부하세요.");
      return;
    }

    const formData = new FormData();
    formData.append("doc_type", currentDocType);
    formData.append("document", currentFile);

    setLoading(true);
    try {
      const response = await fetch("/api/v1/documents/inspect", {
        method: "POST",
        body: formData,
      });
      const payload = await response.json();
      if (!response.ok) {
        throw new Error(payload.detail || "검증에 실패했습니다.");
      }
      renderResult(payload);
    } catch (error) {
      resultPlaceholder.classList.add("hidden");
      resultCard.classList.remove("hidden");
      resultCard.classList.add("fail");
      statusPill.textContent = "ERROR";
      resultTitle.textContent = "검토 실패";
      resultSummary.textContent =
        error instanceof Error
          ? error.message
          : "알 수 없는 오류가 발생했습니다.";
      resultDetails.innerHTML = "";
    } finally {
      setLoading(false);
    }
  }

  docTypeSelect?.addEventListener("change", handleDocTypeChange);
  fileInput?.addEventListener("change", handleFileChange);
  form?.addEventListener("submit", submitForm);
  resetBtn?.addEventListener("click", resetAll);

  resetAll();
})();
