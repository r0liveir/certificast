# Certificast

Python library for generating individual PDF certificates from PPTX templates and
CSV or Python row data. LibreOffice converts populated slides to PDF. Scripts can
generate immediately or stage jobs for validation and publication together.
Dispatch is a separate future feature.

## Domain glossary

| Term | Meaning |
| --- | --- |
| Template | Source document containing named placeholders. |
| Row | Input record supplying values for one generated certificate. |
| Column mapping | Association between a template variable and an input column. |
| Shared variables | Literal fallback values reused across rows in a job. |
| Job | One template, input dataset, variable configuration, and output configuration. |
| Pipeline | In-memory collection of jobs, validated together and published as one run. |
| Staging | Registering a job without generating certificates. |
| Artifact | Generated certificate file. |
| Manifest | Editable CSV associating PDFs with source records and optional recipient emails. |
| Dispatch | Sending generated artifacts to recipients, separately from generation. |

## Sources of truth

- [Constitution](docs/CONSTITUTION.md)
- [Architectural decisions](docs/decisions/)
- [V1 specification](.scratch/certificast-v1/spec.md)
- [Implementation tickets](.scratch/certificast-v1/README.md)

The existing code is a partial prototype. Accepted v1 behavior is defined by the
specification and relevant ADRs, not assumed to be already implemented.
