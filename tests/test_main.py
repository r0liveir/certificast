import subprocess
from pathlib import Path
from unittest.mock import call, patch

import pytest
from pptx import Presentation

import certificast
from certificast.main import _ensure_empty_or_create


def make_template(path: Path) -> None:
    deck = Presentation()
    slide = deck.slides.add_slide(deck.slide_layouts[6])
    slide.shapes.add_textbox(0, 0, 1_000_000, 1_000_000).text = "__NAME__ | __EVENT__"
    deck.save(str(path))


def fake_libreoffice(
    command: list[str], **_: object
) -> subprocess.CompletedProcess[str]:
    inputs = [Path(argument) for argument in command if argument.endswith(".pptx")]
    assert [Presentation(str(path)).slides[0].shapes[0].text for path in inputs] == [
        "Ana | Conference",
        "Bruno | Workshop",
    ]
    for path in inputs:
        path.with_suffix(".pdf").write_bytes(b"%PDF-1.4\n")
    return subprocess.CompletedProcess(command, 0, "", "")


def test_generate(tmp_path: Path) -> None:
    template = tmp_path / "template.pptx"
    csv_file = tmp_path / "people.csv"
    output = tmp_path / "output"
    make_template(template)
    csv_file.write_text("Full name,EVENT\nAna,\nBruno,Workshop\n", encoding="utf-8")

    with (
        patch("certificast.main.shutil.which", return_value="libreoffice"),
        patch("certificast.main.subprocess.run", side_effect=fake_libreoffice),
    ):
        result = certificast.generate(
            str(template),
            str(csv_file),
            str(output),
            {"NAME": "Full name"},
            {"EVENT": "Conference"},
        )

    assert result == [output / "0001.pdf", output / "0002.pdf"]


def test_rejects_missing_value(tmp_path: Path) -> None:
    template = tmp_path / "template.pptx"
    csv_file = tmp_path / "people.csv"
    make_template(template)
    csv_file.write_text("NAME,EVENT\n,Workshop\n", encoding="utf-8")

    with pytest.raises(ValueError, match="missing value for 'NAME'"):
        certificast.validate(str(template), str(csv_file))


def test_rejects_bad_mapping_and_uncovered_variable(tmp_path: Path) -> None:
    template = tmp_path / "template.pptx"
    csv_file = tmp_path / "people.csv"
    make_template(template)
    csv_file.write_text("Full name\nAna\n", encoding="utf-8")

    with pytest.raises(ValueError, match="unknown variables"):
        certificast.validate(str(template), str(csv_file), {"WRONG": "Full name"})
    with pytest.raises(ValueError, match="missing required columns.*EVENT"):
        certificast.validate(str(template), str(csv_file), {"NAME": "Full name"})


def test_output_must_be_empty(tmp_path: Path) -> None:
    output = tmp_path / "output"
    assert _ensure_empty_or_create(output) == output
    (output / "existing.pdf").touch()
    with pytest.raises(FileExistsError):
        _ensure_empty_or_create(output)


def test_pipeline_defers_validates_and_runs_jobs() -> None:
    pipeline = certificast.Pipeline()
    with (
        patch("certificast.main.validate") as validate,
        patch(
            "certificast.main.generate",
            side_effect=[[Path("first/0001.pdf")], [Path("second/0001.pdf")]],
        ) as generate,
    ):
        pipeline.add(
            input_template="one.pptx", input_file="one.csv", output_dir="first"
        )
        pipeline.add(
            input_template="two.pptx", input_file="two.csv", output_dir="second"
        )
        validate.assert_not_called()
        generate.assert_not_called()

        assert pipeline.run() == [Path("first/0001.pdf"), Path("second/0001.pdf")]

    assert validate.call_args_list == [
        call("one.pptx", "one.csv", None, None),
        call("two.pptx", "two.csv", None, None),
    ]
    assert generate.call_count == 2
