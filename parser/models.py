"""
parser/models.py — Schema 1: Chương Trình Đào Tạo (CTĐT)
"""
from dataclasses import dataclass, field
from typing import Optional
from datetime import datetime, timezone


@dataclass
class ProgramObjective:
    objective_code: str          # PO1, PO2 / general
    objective_type: str          # 'general' | 'specific'
    description_vi: str
    description_en: Optional[str] = None
    display_order: int = 1


@dataclass
class ProgramLearningOutcome:
    outcome_code: str            # PLO1.1, PLO2...
    outcome_group_code: str      # 1, 2, 3
    outcome_group_name_vi: str   # Kiến thức / Kỹ năng / Tự chủ
    outcome_group_name_en: Optional[str] = None
    outcome_subgroup_code: Optional[str] = None
    outcome_subgroup_name_vi: Optional[str] = None
    description_vi: str = ""
    description_en: Optional[str] = None
    level_text: Optional[str] = None
    display_order: int = 1


@dataclass
class POPLOMap:
    po_code: str
    plo_code: str
    mapping_value: str = "x"


@dataclass
class ProgramSection:
    section_code: str
    section_title_vi: str
    section_title_en: Optional[str] = None
    section_type: str = "other"
    content_vi: Optional[str] = None
    content_en: Optional[str] = None
    display_order: int = 1


@dataclass
class ProgramComponent:
    component_code: str          # 1, 1.1, 2.2.1
    component_name_vi: str
    component_name_en: Optional[str] = None
    component_type: str = "component_group"
    credits_required: Optional[float] = None
    credit_text: Optional[str] = None
    counts_toward_total_credits: bool = True
    note_text: Optional[str] = None
    parent_code: Optional[str] = None
    display_order: int = 1


@dataclass
class CourseSemesterAllocation:
    semester_no: int
    allocated_value: Optional[float] = None
    allocated_text: Optional[str] = None


@dataclass
class ProgramCourse:
    line_no: Optional[int]
    component_code: str
    course_code_snapshot: str
    course_name_vi_snapshot: str
    course_name_en_snapshot: Optional[str] = None
    credit_value: Optional[float] = None
    credit_text: Optional[str] = None
    credit_unit_type: str = "credit"
    is_required: bool = True
    counts_toward_total_credits: bool = True
    is_thesis: bool = False
    is_physical_education: bool = False
    is_defense_education: bool = False
    notes: Optional[str] = None
    prerequisite_code: Optional[str] = None
    teaching_language: Optional[str] = None
    semester_allocations: list = field(default_factory=list)
    display_order: int = 1


@dataclass
class CareerPath:
    description_vi: str
    description_en: Optional[str] = None
    display_order: int = 1


@dataclass
class GraduationRequirement:
    requirement_type: str
    requirement_text_vi: str
    requirement_text_en: Optional[str] = None
    is_mandatory: bool = True
    display_order: int = 1


@dataclass
class ProgramReferenceProgram:
    institution_name: str
    program_name_vi: Optional[str] = None
    program_name_en: Optional[str] = None
    summary_text: Optional[str] = None
    display_order: int = 1


# ── QA check ─────────────────────────────────────────────────────────────────

def _qa_check(record) -> dict:
    issues = []
    score = 100

    if not record.program_name_vi:
        issues.append("MISSING program_name_vi"); score -= 20
    if not record.major_code:
        issues.append("MISSING major_code"); score -= 10
    if record.total_credits is None:
        issues.append("MISSING total_credits"); score -= 10
    if not record.plos:
        issues.append("NO PLOs extracted"); score -= 15
    # Objectives (PO) là optional — nhiều CTĐT không có bảng PO riêng
    if not record.objectives:
        pass  # Không phạt, chỉ ghi chú trong _meta nếu cần
    if not record.components:
        issues.append("NO program structure/components"); score -= 10
    if not record.courses:
        issues.append("NO courses extracted"); score -= 15
    if not record.sections:
        issues.append("NO narrative sections"); score -= 5
    if not record.career_paths:
        issues.append("NO career paths"); score -= 5

    score = max(0, score)
    return {
        "is_ok": len(issues) == 0,
        "issues": issues,
        "completeness_score": score,
        "needs_review": score < 60,
    }


# ── Main record ───────────────────────────────────────────────────────────────

@dataclass
class CurriculumRecord:
    # ── training_programs ────────────────────────────────────────────────
    program_name_vi: str
    program_name_en: Optional[str]
    degree_name_vi: Optional[str]
    degree_name_en: Optional[str]
    level_name_vi: str
    level_name_en: Optional[str]

    # ── training_program_versions ────────────────────────────────────────
    version_label: str
    major_code: Optional[str]
    major_name_vi: Optional[str]
    major_name_en: Optional[str]
    training_type_vi: Optional[str]
    training_type_en: Optional[str]
    language_vi: Optional[str]
    language_en: Optional[str]
    duration_text_vi: Optional[str]
    duration_text_en: Optional[str]
    total_credits: Optional[float]
    applied_from_admission_year: Optional[int]
    adjustment_time_text: Optional[str]
    program_opening_decision_text: Optional[str]
    accreditation_text: Optional[str]
    degree_granting_unit_text: Optional[str]
    academic_management_unit_text: Optional[str]
    philosophy_text_vi: Optional[str]
    philosophy_text_en: Optional[str]
    issued_decision_text: Optional[str]  # "Ban hành theo QĐ..."

    # ── sub-tables ───────────────────────────────────────────────────────
    objectives: list = field(default_factory=list)
    plos: list = field(default_factory=list)
    po_plo_maps: list = field(default_factory=list)
    sections: list = field(default_factory=list)
    components: list = field(default_factory=list)
    courses: list = field(default_factory=list)
    career_paths: list = field(default_factory=list)
    graduation_requirements: list = field(default_factory=list)
    reference_programs: list = field(default_factory=list)

    # ── metadata ─────────────────────────────────────────────────────────
    source_file: str = ""
    extracted_at: str = field(
        default_factory=lambda: datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    )

    def to_dict(self) -> dict:
        qa = _qa_check(self)
        return {
            "training_program": {
                "program_name_vi": self.program_name_vi,
                "program_name_en": self.program_name_en,
                "degree_name_vi": self.degree_name_vi,
                "degree_name_en": self.degree_name_en,
                "level_name_vi": self.level_name_vi,
                "level_name_en": self.level_name_en,
            },
            "training_program_version": {
                "version_label": self.version_label,
                "major_code": self.major_code,
                "major_name_vi": self.major_name_vi,
                "major_name_en": self.major_name_en,
                "training_type_vi": self.training_type_vi,
                "training_type_en": self.training_type_en,
                "language_vi": self.language_vi,
                "language_en": self.language_en,
                "duration_text_vi": self.duration_text_vi,
                "duration_text_en": self.duration_text_en,
                "total_credits": self.total_credits,
                "applied_from_admission_year": self.applied_from_admission_year,
                "adjustment_time_text": self.adjustment_time_text,
                "program_opening_decision_text": self.program_opening_decision_text,
                "accreditation_text": self.accreditation_text,
                "degree_granting_unit_text": self.degree_granting_unit_text,
                "academic_management_unit_text": self.academic_management_unit_text,
                "philosophy_text_vi": self.philosophy_text_vi,
                "philosophy_text_en": self.philosophy_text_en,
                "issued_decision_text": self.issued_decision_text,
            },
            "objectives": [
                {
                    "objective_code": o.objective_code,
                    "objective_type": o.objective_type,
                    "description_vi": o.description_vi,
                    "description_en": o.description_en,
                    "display_order": o.display_order,
                }
                for o in self.objectives
            ],
            "plos": [
                {
                    "outcome_code": p.outcome_code,
                    "outcome_group_code": p.outcome_group_code,
                    "outcome_group_name_vi": p.outcome_group_name_vi,
                    "outcome_group_name_en": p.outcome_group_name_en,
                    "outcome_subgroup_code": p.outcome_subgroup_code,
                    "outcome_subgroup_name_vi": p.outcome_subgroup_name_vi,
                    "description_vi": p.description_vi,
                    "description_en": p.description_en,
                    "level_text": p.level_text,
                    "display_order": p.display_order,
                }
                for p in self.plos
            ],
            "po_plo_maps": [
                {
                    "po_code": m.po_code,
                    "plo_code": m.plo_code,
                    "mapping_value": m.mapping_value,
                }
                for m in self.po_plo_maps
            ],
            "sections": [
                {
                    "section_code": s.section_code,
                    "section_title_vi": s.section_title_vi,
                    "section_title_en": s.section_title_en,
                    "section_type": s.section_type,
                    "content_vi": s.content_vi,
                    "content_en": s.content_en,
                    "display_order": s.display_order,
                }
                for s in self.sections
            ],
            "components": [
                {
                    "component_code": c.component_code,
                    "component_name_vi": c.component_name_vi,
                    "component_name_en": c.component_name_en,
                    "component_type": c.component_type,
                    "credits_required": c.credits_required,
                    "credit_text": c.credit_text,
                    "counts_toward_total_credits": c.counts_toward_total_credits,
                    "note_text": c.note_text,
                    "parent_code": c.parent_code,
                    "display_order": c.display_order,
                }
                for c in self.components
            ],
            "courses": [
                {
                    "line_no": c.line_no,
                    "component_code": c.component_code,
                    "course_code_snapshot": c.course_code_snapshot,
                    "course_name_vi_snapshot": c.course_name_vi_snapshot,
                    "course_name_en_snapshot": c.course_name_en_snapshot,
                    "credit_value": c.credit_value,
                    "credit_text": c.credit_text,
                    "credit_unit_type": c.credit_unit_type,
                    "is_required": c.is_required,
                    "counts_toward_total_credits": c.counts_toward_total_credits,
                    "is_thesis": c.is_thesis,
                    "is_physical_education": c.is_physical_education,
                    "is_defense_education": c.is_defense_education,
                    "notes": c.notes,
                    "prerequisite_code": c.prerequisite_code,
                    "teaching_language": c.teaching_language,
                    "semester_allocations": [
                        {
                            "semester_no": a.semester_no,
                            "allocated_value": a.allocated_value,
                            "allocated_text": a.allocated_text,
                        }
                        for a in c.semester_allocations
                    ],
                    "display_order": c.display_order,
                }
                for c in self.courses
            ],
            "career_paths": [
                {
                    "description_vi": cp.description_vi,
                    "description_en": cp.description_en,
                    "display_order": cp.display_order,
                }
                for cp in self.career_paths
            ],
            "graduation_requirements": [
                {
                    "requirement_type": gr.requirement_type,
                    "requirement_text_vi": gr.requirement_text_vi,
                    "requirement_text_en": gr.requirement_text_en,
                    "is_mandatory": gr.is_mandatory,
                    "display_order": gr.display_order,
                }
                for gr in self.graduation_requirements
            ],
            "reference_programs": [
                {
                    "institution_name": rp.institution_name,
                    "program_name_vi": rp.program_name_vi,
                    "program_name_en": rp.program_name_en,
                    "summary_text": rp.summary_text,
                    "display_order": rp.display_order,
                }
                for rp in self.reference_programs
            ],
            "_meta": {
                "source_file": self.source_file,
                "extracted_at": self.extracted_at,
            },
            "_qa": qa,
        }
