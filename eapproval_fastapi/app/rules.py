from __future__ import annotations

import json
import re
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Tuple

from .models import (
    Attachment,
    DocumentPayload,
    RulesetMetadata,
    ValidationIssue,
    ValidationResponse,
)


ROLE_LABELS: Dict[str, str] = {
    "ROLE_LEAD": "Team Lead",
    "ROLE_HEAD": "Division Head",
    "ROLE_FIN": "Finance Approval",
    "ROLE_EXE": "Executive Approval",
    "ROLE_PUR": "Procurement Team",
    "ROLE_LGL": "Legal Review",
    "ROLE_SEC": "Security Review",
    "ROLE_CEO": "Chief Executive",
}

ROLE_ORDER: Dict[str, int] = {
    "ROLE_LEAD": 10,
    "ROLE_HEAD": 20,
    "ROLE_PUR": 25,
    "ROLE_FIN": 30,
    "ROLE_LGL": 40,
    "ROLE_SEC": 45,
    "ROLE_EXE": 50,
    "ROLE_CEO": 60,
}

ATTACHMENT_LABELS: Dict[str, str] = {
    "quote": "Quotation",
    "plan": "Event Plan",
    "receipt": "Invoice / Receipt",
    "card_statement": "Corporate Card Statement",
    "security_review": "Security Review Report",
    "legal_review": "Legal Review Report",
    "contract": "Contract",
    "inspection": "Inspection Report",
    "recipient_list": "Recipient List",
    "nda_original": "NDA Original",
    "dpa": "DPA (Data Processing Agreement)",
    "medical_certificate": "Medical Certificate",
    "family_event": "Family Event Evidence",
}

RISK_FLAG_LABELS: Dict[str, str] = {
    "personal_data": "Personal Data",
    "it_saas": "IT / SaaS Adoption",
    "event": "Event / Promotion",
    "leave_sick": "Sick Leave",
    "leave_family": "Family Event Leave",
}

@dataclass(frozen=True)
class ApprovalRule:
    doc_type: str
    min_amount: float
    max_amount: Optional[float]
    required_roles: List[str]
    allow_delegate: List[str]


@dataclass(frozen=True)
class AttachmentRequirement:
    doc_type: str
    type: str
    min_count: int
    required_risk_flags: List[str]
    note: str


@dataclass(frozen=True)
class RiskRule:
    risk_flag: str
    doc_types: List[str]
    required_roles: List[str]
    required_attachments: List[str]
    note: str


class RuleEngine:
    """In-memory ruleset with validation helpers."""

    def __init__(self, rules_path: Path) -> None:
        self.rules_path = rules_path
        self._raw: Dict[str, Any] = {}
        self._doc_no_pattern: re.Pattern[str] | None = None
        self._approval_rules: List[ApprovalRule] = []
        self._attachment_requirements: List[AttachmentRequirement] = []
        self._risk_rules: List[RiskRule] = []
        self._attachment_catalog: Dict[str, Dict[str, Any]] = {}
        self._risk_catalog: Dict[str, Dict[str, Any]] = {}
        self._role_catalog: Dict[str, str] = {}
        self.load()

    # --------------------------------------------------------------------- load
    def load(self) -> None:
        if not self.rules_path.exists():
            raise FileNotFoundError(f"Rules file not found: {self.rules_path}")
        with self.rules_path.open("r", encoding="utf-8") as handle:
            self._raw = json.load(handle)

        patterns = self._raw.get("patterns", {})
        doc_no_regex = patterns.get("doc_no")
        if not doc_no_regex:
            raise ValueError("Rules JSON missing doc_no pattern")
        self._doc_no_pattern = re.compile(doc_no_regex)

        self._approval_rules = [
            ApprovalRule(
                doc_type=entry["doc_type"],
                min_amount=float(entry.get("min_amount", 0)),
                max_amount=float(entry["max_amount"])
                if entry.get("max_amount") is not None
                else None,
                required_roles=entry.get("required_roles", []),
                allow_delegate=entry.get("allow_delegation_for", []),
            )
            for entry in self._raw.get("approval_requirements", [])
        ]

        self._attachment_requirements = []
        self._attachment_catalog = {}
        self._risk_catalog = {}
        for entry in self._raw.get("attachment_requirements", []):
            doc_type = entry["doc_type"]
            for cond in entry.get("conditions", []):
                self._attachment_requirements.append(
                    AttachmentRequirement(
                        doc_type=doc_type,
                        type=cond["type"],
                        min_count=int(cond.get("min_count", 1)),
                        required_risk_flags=cond.get("required_risk_flags", []),
                        note=cond.get("note", ""),
                    )
                )
                label = _label_from_mapping(cond["type"], ATTACHMENT_LABELS)
                attachment_entry = self._attachment_catalog.setdefault(
                    cond["type"], {"label": label, "notes": set()}
                )
                if cond.get("note"):
                    attachment_entry["notes"].add(cond["note"])
                for risk_flag in cond.get("required_risk_flags", []):
                    risk_entry = self._risk_catalog.setdefault(
                        risk_flag,
                        {
                            "label": _label_from_mapping(
                                risk_flag, RISK_FLAG_LABELS
                            ),
                            "notes": set(),
                        },
                    )
                    if cond.get("note"):
                        risk_entry["notes"].add(cond["note"])

        self._risk_rules = [
            RiskRule(
                risk_flag=item["risk_flag"],
                doc_types=item.get("doc_types", []),
                required_roles=item.get("required_roles", []),
                required_attachments=item.get("required_attachments", []),
                note=item.get("note", ""),
            )
            for item in self._raw.get("risk_requirements", [])
        ]
        for rule in self._risk_rules:
            risk_entry = self._risk_catalog.setdefault(
                rule.risk_flag,
                {
                    "label": _label_from_mapping(
                        rule.risk_flag, RISK_FLAG_LABELS
                    ),
                    "notes": set(),
                },
            )
            if rule.note:
                risk_entry["notes"].add(rule.note)
            for role in rule.required_roles:
                self._role_catalog.setdefault(
                    role, _label_from_mapping(role, ROLE_LABELS)
                )

        for approval_rule in self._approval_rules:
            for role in approval_rule.required_roles:
                self._role_catalog.setdefault(
                    role, _label_from_mapping(role, ROLE_LABELS)
                )

    # --------------------------------------------------------------- properties
    @property
    def version(self) -> str:
        return str(self._raw.get("version", "unknown"))

    @property
    def metadata(self) -> RulesetMetadata:
        updated_at = self._raw.get("updated_at", datetime.utcnow().isoformat())
        description = self._raw.get("description", "E-approval validation rules")
        stats = {
            "doc_types": sorted({rule.doc_type for rule in self._approval_rules}),
            "approval_rules": len(self._approval_rules),
            "attachment_rules": len(self._attachment_requirements),
            "risk_rules": len(self._risk_rules),
        }
        return RulesetMetadata(
            version=self.version,
            updated_at=updated_at,
            description=description,
            stats=stats,
        )

    # --------------------------------------------------------------- validation
    def validate_document(self, payload: DocumentPayload) -> ValidationResponse:
        issues: List[ValidationIssue] = []

        issues.append(self._validate_doc_number(payload))
        issues.append(self._validate_doc_type(payload))
        issues.extend(self._validate_approval_chain(payload))
        issues.extend(self._validate_attachments(payload))
        issues.extend(self._validate_risk_rules(payload))

        passed = all(issue.passed for issue in issues)
        return ValidationResponse(
            passed=passed,
            rules_version=self.version,
            issues=issues,
        )

    # ------------------------------------------------------------- rich helpers
    @property
    def doc_type_options(self) -> List[Dict[str, str]]:
        doc_types = self._raw.get("doc_types", [])
        if doc_types:
            return [
                {
                    "code": item["doc_type"],
                    "label": item.get("label", item["doc_type"]),
                }
                for item in doc_types
            ]
        doc_type_set = {rule.doc_type for rule in self._approval_rules}
        return [
            {"code": code, "label": code}
            for code in sorted(doc_type_set)
        ]

    @property
    def attachment_options(self) -> List[Dict[str, str]]:
        options: List[Dict[str, str]] = []
        for code, info in sorted(
            self._attachment_catalog.items(), key=lambda item: item[1]["label"]
        ):
            notes = sorted(info["notes"])
            options.append(
                {
                    "code": code,
                    "label": info["label"],
                    "note": " · ".join(notes),
                }
            )
        return options

    @property
    def risk_flag_options(self) -> List[Dict[str, str]]:
        options: List[Dict[str, str]] = []
        for code, info in sorted(
            self._risk_catalog.items(), key=lambda item: item[1]["label"]
        ):
            notes = sorted(info["notes"])
            options.append(
                {
                    "code": code,
                    "label": info["label"],
                    "note": " · ".join(notes),
                }
            )
        return options

    @property
    def role_options(self) -> List[Dict[str, str]]:
        items: List[Tuple[str, str]] = []
        for code, label in self._role_catalog.items():
            order = ROLE_ORDER.get(code, 999)
            items.append((code, label, order))
        sorted_items = sorted(items, key=lambda item: (item[2], item[0]))
        return [
            {"code": code, "label": label, "order": order}
            for code, label, order in sorted_items
        ]

    def _validate_doc_number(self, payload: DocumentPayload) -> ValidationIssue:
        assert self._doc_no_pattern is not None
        match = self._doc_no_pattern.match(payload.doc_no or "")
        return ValidationIssue(
            rule="doc_no_format",
            passed=bool(match),
            message=(
                "Document number matches required pattern"
                if match
                else "Document number does not match required pattern"
            ),
            details={"doc_no": payload.doc_no, "pattern": self._doc_no_pattern.pattern},
        )

    def _validate_doc_type(self, payload: DocumentPayload) -> ValidationIssue:
        known_doc_types = {
            entry["doc_type"] for entry in self._raw.get("doc_types", [])
        } or {rule.doc_type for rule in self._approval_rules}
        passed = payload.doc_type in known_doc_types
        return ValidationIssue(
            rule="doc_type_known",
            passed=passed,
            message=(
                "Document type is registered in ruleset"
                if passed
                else f"Unknown document type '{payload.doc_type}'"
            ),
            details={
                "doc_type": payload.doc_type,
                "allowed": sorted(known_doc_types),
            },
        )

    def _validate_approval_chain(
        self, payload: DocumentPayload
    ) -> Iterable[ValidationIssue]:
        amount = float(payload.amount_total or 0.0)
        roles_present = [member.role for member in payload.approval_chain]

        applicable_rules = [
            rule
            for rule in self._approval_rules
            if rule.doc_type == payload.doc_type
            and amount >= rule.min_amount
            and (
                rule.max_amount is None or amount <= rule.max_amount + 1e-6
            )  # small epsilon for floats
        ]

        if not applicable_rules:
            yield ValidationIssue(
                rule="approval_rules_missing",
                passed=False,
                message=(
                    "No approval rule found for document type and amount "
                    f"(doc_type={payload.doc_type}, amount={amount})"
                ),
                details={"amount": amount, "doc_type": payload.doc_type},
            )
            return

        for rule in applicable_rules:
            missing_roles = [
                role for role in rule.required_roles if role not in roles_present
            ]
            passed = not missing_roles
            yield ValidationIssue(
                rule=f"approval_roles::{rule.doc_type}::{rule.min_amount}-{rule.max_amount}",
                passed=passed,
                message=(
                    "Approval chain meets required roles"
                    if passed
                    else "Approval chain missing required roles"
                ),
                details={
                    "required_roles": rule.required_roles,
                    "present_roles": roles_present,
                    "missing_roles": missing_roles,
                    "allow_delegation_for": rule.allow_delegate,
                },
            )

    def _validate_attachments(
        self, payload: DocumentPayload
    ) -> Iterable[ValidationIssue]:
        attachments_by_type = _count_by_type(payload.attachments)
        for requirement in self._attachment_requirements:
            if requirement.doc_type != payload.doc_type:
                continue

            if requirement.required_risk_flags and not set(
                requirement.required_risk_flags
            ).issubset(set(payload.risk_flags)):
                continue

            provided = attachments_by_type.get(requirement.type, 0)
            passed = provided >= requirement.min_count
            yield ValidationIssue(
                rule=f"attachment::{requirement.doc_type}::{requirement.type}",
                passed=passed,
                message=(
                    f"Attachment '{requirement.type}' requirement satisfied"
                    if passed
                    else f"Attachment '{requirement.type}' requirement not met"
                ),
                details={
                    "required_min": requirement.min_count,
                    "provided": provided,
                    "note": requirement.note,
                    "risk_flags": payload.risk_flags,
                },
            )

    def _validate_risk_rules(
        self, payload: DocumentPayload
    ) -> Iterable[ValidationIssue]:
        roles_present = {member.role for member in payload.approval_chain}
        attachments_types = {attachment.type for attachment in payload.attachments}

        for rule in self._risk_rules:
            if rule.risk_flag not in payload.risk_flags:
                continue
            if rule.doc_types and payload.doc_type not in rule.doc_types:
                continue

            missing_roles = [role for role in rule.required_roles if role not in roles_present]
            missing_attachments = [
                required for required in rule.required_attachments if required not in attachments_types
            ]

            passed = not missing_roles and not missing_attachments
            yield ValidationIssue(
                rule=f"risk::{rule.risk_flag}",
                passed=passed,
                message=(
                    f"Risk flag '{rule.risk_flag}' requirements satisfied"
                    if passed
                    else f"Risk flag '{rule.risk_flag}' requirements not met"
                ),
                details={
                    "required_roles": rule.required_roles,
                    "missing_roles": missing_roles,
                    "required_attachments": rule.required_attachments,
                    "missing_attachments": missing_attachments,
                    "note": rule.note,
                },
            )


def _count_by_type(attachments: Iterable[Attachment]) -> Dict[str, int]:
    counts: Dict[str, int] = {}
    for item in attachments:
        counts[item.type] = counts.get(item.type, 0) + 1
    return counts


def _label_from_mapping(code: str, mapping: Dict[str, str]) -> str:
    if code in mapping:
        return mapping[code]
    return code.replace("_", " ").title()
