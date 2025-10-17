from __future__ import annotations

from typing import Any, Dict, List, Optional

from pydantic import BaseModel, Field


class ApprovalMember(BaseModel):
    role: str = Field(..., description="Role code such as ROLE_LEAD")
    user_id: Optional[str] = Field(
        default=None, description="Optional identifier of the approver"
    )


class Attachment(BaseModel):
    filename: str
    type: str = Field(..., description="Logical attachment type, e.g. 'quote'")


class DocumentPayload(BaseModel):
    doc_no: str = Field(..., description="Document number to validate")
    doc_type: str = Field(..., description="Document type code (e.g. EXR, PR)")
    title: Optional[str] = None
    amount_total: Optional[float] = Field(
        default=None, description="Numeric amount used for threshold rules"
    )
    risk_flags: List[str] = Field(
        default_factory=list,
        description="List of risk attributes such as 'personal_data', 'it_saas'",
    )
    approval_chain: List[ApprovalMember] = Field(
        default_factory=list, description="Ordered list of approval roles"
    )
    attachments: List[Attachment] = Field(
        default_factory=list, description="List of provided attachments"
    )


class ValidationIssue(BaseModel):
    rule: str = Field(..., description="Identifier of the validation rule")
    passed: bool = Field(..., description="Whether the rule passed")
    message: str = Field(..., description="Human readable summary")
    details: Dict[str, Any] = Field(
        default_factory=dict, description="Structured extra information"
    )


class ValidationResponse(BaseModel):
    passed: bool
    rules_version: str
    issues: List[ValidationIssue] = Field(default_factory=list)


class RulesetMetadata(BaseModel):
    version: str
    updated_at: str
    description: str
    stats: Dict[str, Any] = Field(default_factory=dict)


class ReloadResponse(BaseModel):
    message: str
    rules_version: str

