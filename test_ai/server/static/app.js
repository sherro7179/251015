document.addEventListener("DOMContentLoaded", () => {
  const form = document.getElementById("upload-form");
  const resultContainer = document.getElementById("result");
  const statusText = document.getElementById("status-text");
  const submitButton = document.getElementById("submit-btn");

  const renderList = (title, items, cssClass) => {
    if (!items || items.length === 0) {
      return "";
    }
    return `
      <div class="alert ${cssClass || ""}">
        <strong>${title}</strong>
        <ul class="missing-list">
          ${items.map((item) => `<li>${item}</li>`).join("")}
        </ul>
      </div>
    `;
  };

  form.addEventListener("submit", async (event) => {
    event.preventDefault();
    resultContainer.innerHTML = "";
    statusText.textContent = "분석 중입니다...";
    submitButton.disabled = true;

    const formData = new FormData(form);

    try {
      const response = await fetch("/api/validate", {
        method: "POST",
        body: formData,
      });

      const payload = await response.json();

      if (!response.ok) {
        throw new Error(payload.error || "분석 요청에 실패했습니다.");
      }

      const summaryHtml = payload.summary
        .map((line) => `<div class="pill">${line}</div>`)
        .join("");

      resultContainer.innerHTML = `
        <div class="result">
          <h3>분석 결과</h3>
          <p><strong>양식 유사도:</strong> ${payload.templateSimilarity}%</p>
          <div>${summaryHtml}</div>
          ${renderList("누락된 섹션", payload.missingSections, "alert")}
          ${renderList("규정 키워드 미포함", payload.missingKeywords, "alert")}
        </div>
      `;
      statusText.textContent = "분석이 완료되었습니다.";
    } catch (error) {
      resultContainer.innerHTML = `
        <div class="alert">
          <strong>오류:</strong> ${error.message}
        </div>
      `;
      statusText.textContent = "분석 중 오류가 발생했습니다.";
    } finally {
      submitButton.disabled = false;
    }
  });
});
