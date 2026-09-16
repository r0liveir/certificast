[![PyPI - Version](https://img.shields.io/pypi/v/certificast)](https://pypi.org/project/certificast/)
[![PyPI - Python Version](https://img.shields.io/pypi/pyversions/certificast)](https://pypi.org/project/certificast/)

# certificast

Yet another certificate generator \o/

Generate numbered PDF certificates from a template certificate and a file.

Current scope: one PPTX file and one CSV file.

Requires Python 3.12+ and LibreOffice (`libreoffice` or `soffice`) on `PATH`.

## Why this?

Not sure 😕 but hope this helps someone!

## Generate certificates

Before generating, template variables must use uppercase names surrounded by double underscores:

```text
Certificate awarded to __NAME__ at __EVENT__
```

Example usage:

```python
import certificast

certificates = certificast.generate(
    input_template="certificate.pptx",
    input_file="people.csv",
    output_dir="output",
)

# print file names
for certificate in certificates:
    print(certificate)
```

You can also setup a `dict`

The output directory is created if necessary and must be empty. Each CSV row
produces `0001.pdf`, `0002.pdf`, and so on.

Use `columns_mapping` when CSV headers differ from template variables. Blank
cells fall back to `defaults` (Useful for one template variable shared on certificates):

```python
certificast.generate(
    "certificate.pptx",
    "people.csv",
    "output",
    columns_mapping={"NAME": "Full name"},
    defaults={"EVENT": "Annual Conference"},
    output_name="{NAME}_{EVENT}",
)
```

Custom names are prefixed with the row number and sanitized, producing names
such as `0001-Ana_Annual Conference.pdf`.

## Validate without generating

```python
rows = certificast.validate(
    "certificate.pptx",
    "people.csv",
    {"NAME": "Full name"},
    {"EVENT": "Annual Conference"},
)
```

Validation checks the template, CSV headers and rows, mappings, defaults, and
variable coverage. It raises `ValueError` for invalid input and returns the
resolved rows when successful.

## Defer multiple jobs

```python
pipeline = certificast.Pipeline()

pipeline.add(
    input_template="certificate.pptx",
    input_file="attendees.csv",
    output_dir="output/attendees",
)
pipeline.add(
    input_template="certificate.pptx",
    input_file="speakers.csv",
    output_dir="output/speakers",
    defaults={"EVENT": "Annual Conference"},
)

pipeline.validate()
certificates = pipeline.run()
```

Adding a job does not read or generate anything. `run()` validates every job
before starting generation, then returns all generated PDF paths. Jobs currently
publish independently; a later job failure can leave earlier output in place.

## Development

```sh
uv sync --all-groups
PYTHONPATH=src uv run pytest -q
uv run ruff check src tests
uv run mypy --strict src/certificast/main.py tests/test_main.py
```
