import pytest
from pathlib import Path
from src.core.reporter import RunReporter


def test_reporter_generation(tmp_path):
    reporter = RunReporter("Factory Automation Run #42")
    
    reporter.add_result(
        part_number="P-001",
        name="Main Bracket",
        part_type="plate_with_holes",
        mass_kg=0.85,
        volume_cm3=110.0,
        drc_passed=True,
        exports=["pdf", "step", "dwg"],
    )
    reporter.add_result(
        part_number="P-002",
        name="Drive Shaft",
        part_type="cylinder",
        mass_kg=1.45,
        volume_cm3=180.0,
        drc_passed=False,
        drc_details="Wall thickness below tolerance",
        exports=["pdf"],
    )

    assert reporter.total_count == 2
    assert reporter.passed_count == 1
    assert reporter.failed_count == 1

    # Test Markdown generation
    md_file = tmp_path / "report.md"
    md_output = reporter.generate_markdown(md_file)
    assert md_file.exists()
    assert "Factory Automation Run #42" in md_output
    assert "P-001" in md_output
    assert "P-002" in md_output
    assert "Drive Shaft" in md_output

    # Test HTML generation
    html_file = tmp_path / "report.html"
    html_output = reporter.generate_html(html_file)
    assert html_file.exists()
    assert "<!DOCTYPE html>" in html_output
    assert "badge-pass" in html_output
    assert "badge-fail" in html_output
