# certificast

A unified tool for bulk certificate generation and broadcast.

## First Supported Path

Certificast currently targets the narrowest useful workflow:

- PPTX certificate template
- Jinja-style variables such as `{{ person_name }}`
- CSV or Python row data
- explicit variable mapping
- one generated PPTX per row

## Python Usage

```python
import certificast

generated = certificast.generate(
    input_template="certificate-template.pptx",
    dict_variables={
        "person_name": "Name",
        "event_name": "Event",
        "email": "Email",
    },
    csv_file="recipients.csv",
    output_dir="certificates",
)

for certificate in generated:
    print(certificate.output_path)
```

By default, output files are named incrementally:

```txt
0001.pptx
0002.pptx
0003.pptx
```

Custom names can use CSV columns and `{index}`:

```python
certificast.generate(
    "certificate-template.pptx",
    {"person_name": "Name"},
    csv_file="recipients.csv",
    output_name="{index}-{Name}",
)
```

Missing template variables warn and fail generation.

## CLI Usage

```sh
certificast certificate-template.pptx recipients.csv \
  --variables '{"person_name":"Name","event_name":"Event"}' \
  --output-dir certificates
```
