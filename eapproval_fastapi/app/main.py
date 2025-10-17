from __future__ import annotations

import io
import zipfile
from pathlib import Path
import xml.etree.ElementTree as ET
import re
from typing import Any, Dict, List, Optional, Set

from fastapi import Depends, FastAPI, File, Form, HTTPException, Request, UploadFile
from fastapi.responses import HTMLResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates

from .config import get_settings
from .models import DocumentPayload, ReloadResponse, RulesetMetadata, ValidationResponse
from .rules import RuleEngine

DOC_TEMPLATES: dict[str, dict] = {
    "EXR": {
        "code": "EXR",
        "name": "Expense Request",
        "summary": "Pre-approval for marketing events and promotions.",
        "fields": [
            {"label": "Drafting Department", "value": "Marketing Team"},
            {"label": "Drafter", "value": "Lee Mark (Team Lead)"},
            {"label": "Event Schedule", "value": "2025-11-05 14:00 ~ 18:00"},
            {"label": "Key Items", "value": "Venue rental, catering, promotional materials"}
        ],
        "attachments": [
            "At least two quotations (three recommended)",
            "Event plan (required for on-site activities)"
        ],
        "tips": [
            "Break down the total budget (8,800,000 KRW VAT incl.) by item.",
            "Prepare an after-action report within 3 business days."
        ],
        "approval_flow": ["ROLE_LEAD", "ROLE_HEAD", "ROLE_FIN"],
        "default_risk_flags": ["event"],
        "sample_docx": "\uACB0\uC7AC\uC11C_\uC0D8\uD50C_\uC9C0\uCD9C\uD488\uC758_v2.docx"
    },
    "EXP": {
        "code": "EXP",
        "name": "Expense Report",
        "summary": "Post-travel or activity expense settlement.",
        "fields": [
            {"label": "Drafting Department", "value": "Sales Team"},
            {"label": "Drafter", "value": "Oh Sales (Assistant Manager)"},
            {"label": "Trip Purpose", "value": "Busan partner meetings"},
            {"label": "Key Costs", "value": "KTX, rental car, accommodation, meals"}
        ],
        "attachments": [
            "Invoices or receipts for each cost item",
            "Corporate card statement (if the card was used)"
        ],
        "tips": [
            "Missing receipts delay reimbursement."
        ],
        "approval_flow": ["ROLE_LEAD", "ROLE_FIN"],
        "default_risk_flags": [],
        "sample_docx": "\uCD94\uAC00\uC0D8\uD50C_\uC9C0\uCD9C\uACB0\uC758_v2.docx"
    },
    "PR": {
        "code": "PR",
        "name": "Purchase Request",
        "summary": "Request for new products or SaaS services.",
        "fields": [
            {"label": "Requesting Department", "value": "Development Division"},
            {"label": "Requester", "value": "Park Dev (Manager)"},
            {"label": "Primary Item", "value": "SaaS collaboration tool & SSO module"},
            {"label": "Subscription Term", "value": "12 months"}
        ],
        "attachments": [
            "Three quotations (exception narrative if fewer)",
            "Security review document for IT/SaaS",
            "Legal review document if personal data is processed"
        ],
        "tips": [
            "Confirm personal data processing and complete PIA if required."
        ],
        "approval_flow": ["ROLE_LEAD", "ROLE_HEAD", "ROLE_PUR", "ROLE_FIN"],
        "default_risk_flags": ["it_saas", "personal_data"],
        "sample_docx": "\uACB0\uC7AC\uC11C_\uC0D8\uD50C_\uAD6C\uB9E4\uC694\uCCAD_v2.docx"
    },
    "PO": {
        "code": "PO",
        "name": "Purchase Order",
        "summary": "Issue to supplier after purchase approval.",
        "fields": [
            {"label": "Ordering Department", "value": "Procurement Team"},
            {"label": "Item", "value": "New hire laptops (3 units)"},
            {"label": "Supplier", "value": "IT Zone"},
            {"label": "Delivery", "value": "Within 7 days of order"}
        ],
        "attachments": [
            "Contract or unit-price agreement",
            "Inspection report / tax invoice linkage"
        ],
        "tips": [
            "Double-check quantity and unit price before sending the PO."
        ],
        "approval_flow": ["ROLE_PUR", "ROLE_FIN"],
        "default_risk_flags": [],
        "sample_docx": "\uCD94\uAC00\uC0D8\uD50C_PO\uBC1C\uC8FC_v2.docx"
    },
    "OFF": {
        "code": "OFF",
        "name": "Official Letter",
        "summary": "Formal notice to partners or external organisations.",
        "fields": [
            {"label": "Subject", "value": "Quarterly security audit notice"},
            {"label": "Recipients", "value": "Partner security lead / Sales division"},
            {"label": "Issue Date", "value": "2025-10-19"},
            {"label": "Response Deadline", "value": "2025-11-02"}
        ],
        "attachments": [
            "Recipient list (required for bulk notices)"
        ],
        "tips": [
            "Align document and attachment names for traceability."
        ],
        "approval_flow": ["ROLE_LEAD", "ROLE_EXE"],
        "default_risk_flags": [],
        "sample_docx": "\uACB0\uC7AC\uC11C_\uC0D8\uD50C_\uACF5\uBB38\uBC1C\uC2E0_v2.docx"
    },
    "NDA": {
        "code": "NDA",
        "name": "NDA Approval",
        "summary": "Non-disclosure agreement with external partners.",
        "fields": [
            {"label": "Counterparty", "value": "O.O Solutions Co., Ltd."},
            {"label": "Purpose", "value": "AI joint PoC data exchange"},
            {"label": "Term", "value": "3 years from execution"},
            {"label": "Owner", "value": "Legal Counsel"}
        ],
        "attachments": [
            "Executed NDA",
            "DPA (required if personal data is exchanged)"
        ],
        "tips": [
            "Complete legal and security review before signing."
        ],
        "approval_flow": ["ROLE_LEAD", "ROLE_LGL", "ROLE_EXE"],
        "default_risk_flags": ["personal_data"],
        "sample_docx": "\uACB0\uC7AC\uC11C_\uC0D8\uD50C_NDA\uC2B9\uC778_v2.docx"
    },
    "LV": {
        "code": "LV",
        "name": "Leave Application",
        "summary": "Request for annual, sick, or family leave.",
        "fields": [
            {"label": "Applicant", "value": "Employee Jeong"},
            {"label": "Department", "value": "Platform Operations"},
            {"label": "Leave Type", "value": "Sick leave"},
            {"label": "Period", "value": "2025-10-20 ~ 2025-10-22"}
        ],
        "attachments": [
            "Medical certificate for sick leave",
            "Family event evidence for ceremonial leave"
        ],
        "tips": [
            "Nominate a stand-in approver to prevent workflow delays."
        ],
        "approval_flow": ["ROLE_LEAD"],
        "default_risk_flags": ["leave_sick"],
        "sample_docx": "\uACB0\uC7AC\uC11C_\uC0D8\uD50C_\uD734\uAC00\uC2E0\uCCAD_v2.docx",
        "regulation_keywords": [
            "\uCCAD\uBD80 (\uCCB4\uD06C\uB9AC\uC2A4\uD2B8)",
            "(\uBCD1\uAC00/\uACBD\uC870) \uC99D\uBE59 \uCCAD\uBD80"
        ]
    },
    "POL": {
        "code": "POL",
        "name": "Policy Draft / Revision",
        "summary": "Create or amend internal policies.",
        "fields": [
            {"label": "Drafting Department", "value": "Corporate Support"},
            {"label": "Drafter", "value": "Kim Guide"},
            {"label": "Effective Date", "value": "2025-10-31"},
            {"label": "Key Changes", "value": "Document number format, attachment checklist"}
        ],
        "attachments": [
            "Full policy text with comparison table",
            "Updated attachment checklist (if applicable)"
        ],
        "tips": [
            "Schedule training sessions before the effective date."
        ],
        "approval_flow": ["ROLE_LEAD", "ROLE_HEAD", "ROLE_FIN", "ROLE_EXE"],
        "default_risk_flags": [],
        "sample_docx": "\uACB0\uC7AC\uC11C_\uC0D8\uD50C_\uADDC\uC815\uC81C\uC815_v2.docx"
    }
}



app = FastAPI(
    title="Mock E-Approval Validator",
    version="0.1.0",
    description=(
        "Lightweight FastAPI service that validates e-approval documents "
        "against rules derived from the shared regulation package."
    ),
)



def _public_doc_templates() -> Dict[str, Dict[str, Any]]:
    view: Dict[str, Dict[str, Any]] = {}
    for code, data in DOC_TEMPLATES.items():
        view[code] = {
            "code": data["code"],
            "name": data["name"],
            "summary": data["summary"],
            "fields": data.get("fields", []),
            "attachments": data.get("attachments", []),
            "tips": data.get("tips", []),
            "approval_flow": data.get("approval_flow", []),
        }
    return view


PUBLIC_DOC_TEMPLATES = _public_doc_templates()


class TemplateInspectionError(Exception):
    """Raised when template inspection cannot be completed."""


class DocTemplateInspector:
    """Evaluate uploaded DOCX files against baseline templates."""

    def __init__(self, samples_dir: Path, templates: Dict[str, Dict[str, Any]]) -> None:
        self.samples_dir = samples_dir
        self.templates: Dict[str, Dict[str, Any]] = {}
        for code, meta in templates.items():
            sample_file = meta.get("sample_docx")
            sample_lines: List[str] = []
            if sample_file:
                sample_path = samples_dir / sample_file
                if sample_path.exists():
                    sample_lines = self._extract_lines(sample_path.read_bytes())
            structure_display = self._derive_structure_markers(sample_lines)
            if not structure_display:
                structure_display = meta.get("structure_keywords", [])
            regulation_display = self._derive_regulation_markers(sample_lines)
            if not regulation_display:
                regulation_display = meta.get("regulation_keywords", [])
            self.templates[code] = {
                "meta": meta,
                "structure_tokens": [DocTemplateInspector._normalize(line) for line in structure_display],
                "structure_display": structure_display,
                "regulation_tokens": [DocTemplateInspector._normalize(line) for line in regulation_display],
                "regulation_display": regulation_display,
            }

    def inspect(self, doc_type: str, payload: bytes) -> Dict[str, Any]:
        if doc_type not in self.templates:
            raise ValueError(f"Unsupported document type '{doc_type}'")
        try:
            lines = self._extract_lines(payload)
        except TemplateInspectionError:
            raise
        except Exception as exc:
            raise TemplateInspectionError("Failed to read DOCX content.") from exc

        if not lines:
            raise TemplateInspectionError("Document does not contain readable text.")

        metadata = self.templates[doc_type]
        structure_tokens: List[str] = metadata["structure_tokens"]
        structure_display = metadata["structure_display"]
        regulation_tokens: List[str] = metadata["regulation_tokens"]
        regulation_display = metadata["regulation_display"]

        normalized_lines: Set[str] = {
            self._normalize(line) for line in lines if self._normalize(line)
        }

        structure_missing = []
        for token, display in zip(structure_tokens, structure_display):
            if token and token not in normalized_lines:
                structure_missing.append(display)

        regulation_missing = []
        for token, display in zip(regulation_tokens, regulation_display):
            if token and token not in normalized_lines:
                regulation_missing.append(display)

        structure_ok = bool(structure_tokens) and not structure_missing
        regulation_ok = bool(regulation_tokens) and not regulation_missing

        return {
            "doc_type": doc_type,
            "structure": {
                "ok": structure_ok,
                "checked": structure_display,
                "missing": structure_missing,
                "coverage": (
                    0.0
                    if not structure_tokens
                    else round(
                        (len(structure_tokens) - len(structure_missing))
                        / len(structure_tokens),
                        2,
                    )
                ),
            },
            "regulation": {
                "ok": regulation_ok,
                "checked": regulation_display,
                "missing": regulation_missing,
            },
            "passed": structure_ok and regulation_ok,
        }

    @staticmethod
    def _extract_lines(blob: bytes) -> List[str]:
        try:
            with zipfile.ZipFile(io.BytesIO(blob)) as archive:
                xml = archive.read("word/document.xml")
        except Exception as exc:
            raise TemplateInspectionError("Invalid DOCX file.") from exc
        try:
            root = ET.fromstring(xml)
        except ET.ParseError as exc:
            raise TemplateInspectionError("Malformed DOCX XML.") from exc
        ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
        lines: List[str] = []
        for paragraph in root.findall(".//w:p", ns):
            texts = [
                node.text.strip()
                for node in paragraph.findall(".//w:t", ns)
                if node.text and node.text.strip()
            ]
            if texts:
                merged = " ".join(texts)
                cleaned = " ".join(merged.split())
                if cleaned:
                    lines.append(cleaned)
        return lines

    @staticmethod
    def _derive_structure_markers(lines: List[str]) -> List[str]:
        if not lines:
            return []
        keywords = [
            "\ubb38\uc11c\ubc88\ud638",
            "\uacb0\uc7ac\uc120",
            "\uccad\ubd80",
            "\uc608\uc0c1\uae08\uc561",
            "\uc694\uccad",
            "\ud569\uacc4",
            "\uc2b9\uc778",
            "\uc2e0\uccad",
        ]
        markers: List[str] = []
        seen: Set[str] = set()
        for line in lines:
            if line.startswith("결재서"):
                continue
            lower = DocTemplateInspector._normalize(line)
            if (
                any(keyword in line for keyword in keywords)
                or lower.endswith("v2")
            ):
                if lower not in seen:
                    markers.append(line)
                    seen.add(lower)
        if not markers:
            for line in lines[:8]:
                if line.startswith("결재서"):
                    continue
                lower = DocTemplateInspector._normalize(line)
                if lower not in seen:
                    markers.append(line)
                    seen.add(lower)
        return markers

    @staticmethod
    def _derive_regulation_markers(lines: List[str]) -> List[str]:
        markers: List[str] = []
        seen: Set[str] = set()
        for idx, line in enumerate(lines):
            lower = DocTemplateInspector._normalize(line)
            if "\uccad\ubd80" in lower:
                for follow in lines[idx + 1 :]:
                    candidate = follow.strip()
                    if not candidate:
                        continue
                    if candidate.startswith(("-", "•", "*")):
                        normalized = candidate.lstrip("-•* ").strip()
                        lowered = DocTemplateInspector._normalize(normalized)
                        if lowered and lowered not in seen:
                            markers.append(normalized)
                            seen.add(lowered)
                    else:
                        break
        return markers

    @staticmethod
    def _normalize(value: str) -> str:
        lowered = value.lower()
        return re.sub(r"[\s_\-]+", "", lowered)
settings = get_settings()
templates = Jinja2Templates(directory=str(settings.project_root / "templates"))
app.mount(
    "/static",
    StaticFiles(directory=str(settings.project_root / "static")),
    name="static",
)


# --------------------------------------------------------------------- lifespan
@app.on_event("startup")
def _load_rules() -> None:
    settings = get_settings()
    app.state.rule_engine = RuleEngine(settings.rules_path)
    app.state.template_inspector = DocTemplateInspector(settings.sample_docs_dir, DOC_TEMPLATES)


def get_engine() -> RuleEngine:
    engine = getattr(app.state, "rule_engine", None)
    if engine is None:
        settings = get_settings()
        engine = RuleEngine(settings.rules_path)
        app.state.rule_engine = engine
    return engine


def get_template_inspector() -> DocTemplateInspector:
    inspector = getattr(app.state, "template_inspector", None)
    if inspector is None:
        settings = get_settings()
        inspector = DocTemplateInspector(settings.sample_docs_dir, DOC_TEMPLATES)
        app.state.template_inspector = inspector
    return inspector


# --------------------------------------------------------------------- endpoints
@app.get("/health")
def health() -> dict[str, str]:
    return {"status": "ok"}


@app.get("/api/v1/rules", response_model=RulesetMetadata)
def rules_metadata(engine: RuleEngine = Depends(get_engine)) -> RulesetMetadata:
    return engine.metadata


@app.post("/api/v1/rules/reload", response_model=ReloadResponse)
def reload_rules(engine: RuleEngine = Depends(get_engine)) -> ReloadResponse:
    try:
        engine.load()
    except FileNotFoundError as exc:
        raise HTTPException(status_code=404, detail=str(exc)) from exc
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc
    return ReloadResponse(message="Rules reloaded", rules_version=engine.version)


@app.post("/api/v1/validate", response_model=ValidationResponse)
def validate_document(
    payload: DocumentPayload, engine: RuleEngine = Depends(get_engine)
) -> ValidationResponse:
    try:
        return engine.validate_document(payload)
    except Exception as exc:  # pragma: no cover - defensive
        raise HTTPException(status_code=500, detail=str(exc)) from exc


@app.post("/api/v1/documents/inspect")
async def inspect_document(
    doc_type: str = Form(...),
    document: UploadFile = File(...),
    inspector: DocTemplateInspector = Depends(get_template_inspector),
) -> Dict[str, Any]:
    if not document.filename:
        raise HTTPException(status_code=400, detail="Filename is required.")
    if not document.filename.lower().endswith(".docx"):
        raise HTTPException(status_code=400, detail="Only DOCX files are supported.")
    contents = await document.read()
    if not contents:
        raise HTTPException(status_code=400, detail="Uploaded file is empty.")
    try:
        inspection = inspector.inspect(doc_type, contents)
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc
    except TemplateInspectionError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc
    inspection["filename"] = document.filename
    return inspection


@app.get("/", response_class=HTMLResponse)
def landing(
    request: Request, engine: RuleEngine = Depends(get_engine)
) -> HTMLResponse:
    context = {
        "request": request,
        "doc_types": engine.doc_type_options,
        "roles": engine.role_options,
        "doc_templates": PUBLIC_DOC_TEMPLATES,
        "rules_version": engine.version,
    }
    return templates.TemplateResponse("index.html", context)
