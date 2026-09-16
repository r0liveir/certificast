# certificast

A Python library for generating PDF certificates from existing PPTX templates.
LibreOffice is required on `PATH` (`libreoffice` or `soffice`). Install the
library with `pip install -e .` and LibreOffice with your system package manager.

Generation supports a list of Python rows and a one-slide template. CSV
input, column mappings, explicit slide selection, naming, recipients, retention,
and pipelines will follow in the remaining v1 tickets. There is no v1 CLI.

```python
import certificast

try:
    result = certificast.generate(
        template="certificate-template.pptx",
        rows=[{"person_name": "Ana Silva"}],
        shared={"event_name": "Conference"},
        output_dir="certificates",
    )
except certificast.ValidationError as error:
    print(error.report)  # structured errors/warnings with job, row, variable context
else:
    print(result.certificates[0])  # absolute published path to certificates/0001.pdf
    print(result.manifest_path)
```

Templates use simple named placeholders such as `{{ person_name }}` or
`{{- event_name -}}`. Nonblank row values override shared literal fallbacks.
Missing keys, `None`, empty strings, and whitespace-only strings are blank;
zero and false are valid, and other nonblank text keeps its whitespace.
Ordinary text, groups, and tables are traversed. Replacements inside a run
preserve its formatting. Split placeholders use paragraph replacement and
return a warning about possible formatting loss. Existing text fitting settings
are preserved; output fidelity depends on LibreOffice and installed fonts.

All rows are validated before rendering. Each row produces a numbered PDF
(`0001.pdf`, `0002.pdf`, …); LibreOffice converts the populated PPTX files in
one invocation. Rows and template text are prepared once, and every expected
PDF must succeed before the complete batch is published. Empty rows return no
artifacts and a warning. Very large batches may exceed system command-length
limits; streaming and conversion chunking are deferred.

Successful output contains only the PDFs and a UTF-8 comma-separated
`manifest.csv`:

```csv
file,job,row,emails
0001.pdf,1,1,
```

Results expose `output_dir`, `manifest_path`, `certificates`, and `warnings`.
Artifact paths always refer to the published directory. Runtime failures raise
`GenerationError`; validation failures raise `ValidationError` before rendering.

The output parent must exist, and the destination must be new. Publication
currently requires Linux with libc `renameat2` and a filesystem supporting
`RENAME_NOREPLACE`. A uniquely owned temporary sibling holds all work; a
same-filesystem no-replace directory rename publishes the complete run. Even an
empty directory created after validation is preserved and blocks publication.
Unsupported kernels/filesystems fail safely rather than fall back to an unsafe
rename. This guarantees atomic visibility, rather than power-loss durability.
Handled failures remove only owned temporary work. If cleanup fails, the
exception includes the original error, cleanup error, and remaining location;
a process crash can also leave temporary work.

Run the public API checks with LibreOffice and Poppler's `pdfinfo` and
`pdftotext` installed:

```sh
PYTHONPATH=src python -m unittest discover -s tests -v
```

The real conversion check inspects PDF page count and extracted text. Other
checks exercise invalid input, converter failures, destination races, and
cleanup failures at the external process/filesystem boundaries. In environments
that restrict LibreOffice subprocesses, run the checks with appropriate process
permissions. For template fidelity, open a generated PDF and compare its layout,
fonts, and fitting against the source slide; PDF bytes are not a stable snapshot.
