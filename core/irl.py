"""
Compiles the Information Requirement List (Section 10 of the instructions):
every Finding across all Analysis sections flagged needs_irl=True becomes one
row, in the mandated 5-column format.
"""


def build_irl(sections):
    rows = []
    ref = 1
    for section_findings in sections.values():
        for finding in section_findings:
            if not finding.needs_irl:
                continue
            rows.append(
                {
                    "Reference Number": f"IRL-{ref:03d}",
                    "Financial Statement Area": finding.area,
                    "Observation": finding.observation,
                    "Reason for Concern": finding.reason_for_concern,
                    "Information / Clarification Required": finding.info_required,
                }
            )
            ref += 1
    return rows
