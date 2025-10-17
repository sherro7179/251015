from __future__ import annotations

import io
import re
from dataclasses import dataclass
from difflib import SequenceMatcher
from pathlib import Path
from typing import List

from docx import Document
from flask import Flask, jsonify, render_template, request


BASE_DIR = Path(__file__).parent
DATA_DIR = BASE_DIR / "data"


@dataclass(frozen=True)
class ReferenceMaterial:
    regulation_text: str
    template_text: str
    required_sections: List[str]
    required_keywords: List[str]


def _load_reference_material() -> ReferenceMaterial:
    regulation_text = (DATA_DIR / "regulation.md").read_text(encoding="utf-8")
    template_text = (DATA_DIR / "form_template.md").read_text(encoding="utf-8")

    required_sections = [
        line.lstrip("# ").strip()
        for line in template_text.splitlines()
        if line.strip().startswith("## ")
    ]

    # Capture bold keywords such as **지출 목적** from the regulation document.
    required_keywords = re.findall(r"\*\*(.+?)\*\*", regulation_text)

    return ReferenceMaterial(
        regulation_text=regulation_text,
        template_text=template_text,
        required_sections=required_sections,
        required_keywords=required_keywords,
    )


REFERENCE = _load_reference_material()

app = Flask(
    __name__,
    template_folder=str(BASE_DIR / "templates"),
    static_folder=str(BASE_DIR / "static"),
)

app.config["MAX_CONTENT_LENGTH"] = 8 * 1024 * 1024  # 8 MB uploads


def _extract_text_from_upload(filename: str, file_bytes: bytes) -> str:
    lowered = (filename or "").lower()
    if lowered.endswith(".docx"):
        document = Document(io.BytesIO(file_bytes))
        return "\n".join(paragraph.text for paragraph in document.paragraphs).strip()

    # Fall back to UTF-8, then CP949 for common Korean documents.
    try:
        return file_bytes.decode("utf-8").strip()
    except UnicodeDecodeError:
        return file_bytes.decode("cp949", errors="ignore").strip()


def _evaluate_document(document_text: str) -> dict:
    sm = SequenceMatcher(
        None,
        REFERENCE.template_text.replace("\r\n", "\n"),
        document_text.replace("\r\n", "\n"),
    )
    template_similarity = round(sm.ratio() * 100, 2)

    missing_sections = [
        section
        for section in REFERENCE.required_sections
        if section and section not in document_text
    ]

    missing_keywords = [
        keyword
        for keyword in REFERENCE.required_keywords
        if keyword and keyword not in document_text
    ]

    summary = []
    if template_similarity >= 85:
        summary.append("양식과의 유사도가 높습니다.")
    elif template_similarity >= 60:
        summary.append("양식과의 유사도가 보통입니다. 주요 항목을 재검토하세요.")
    else:
        summary.append("양식과의 유사도가 낮습니다. 필수 항목 누락 여부를 확인하세요.")

    if not missing_sections and not missing_keywords:
        summary.append("규정에서 요구하는 핵심 항목이 모두 포함되어 있습니다.")
    else:
        summary.append("누락된 항목을 보완한 후 결재를 진행하세요.")

    return {
        "templateSimilarity": template_similarity,
        "missingSections": missing_sections,
        "missingKeywords": missing_keywords,
        "summary": summary,
    }


@app.get("/")
def index():
    return render_template(
        "index.html",
        regulation_text=REFERENCE.regulation_text,
        template_text=REFERENCE.template_text,
    )


@app.post("/api/validate")
def validate_document():
    uploaded = request.files.get("document")
    if uploaded is None or uploaded.filename == "":
        return jsonify({"error": "파일을 선택해 주세요."}), 400

    file_bytes = uploaded.read()
    if not file_bytes:
        return jsonify({"error": "빈 파일은 분석할 수 없습니다."}), 400

    document_text = _extract_text_from_upload(uploaded.filename, file_bytes)
    if not document_text:
        return jsonify({"error": "문서에서 텍스트를 추출할 수 없습니다."}), 400

    result = _evaluate_document(document_text)
    return jsonify(result)


if __name__ == "__main__":
    # Development-time convenience: `python app.py` launches the server.
    app.run(host="0.0.0.0", port=8000, debug=True)
