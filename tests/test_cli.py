import pytest
from src.cli import parse_args


def test_cli_help(capsys):
    # Test that parsing with --help displays instructions and raises SystemExit
    with pytest.raises(SystemExit) as e:
        import sys
        from unittest.mock import patch

        with patch.object(sys, "argv", ["cli.py", "--help"]):
            parse_args()
    assert e.value.code == 0


def test_cli_arguments():
    import sys
    from unittest.mock import patch

    # Test sub-command parsing for generate-part
    with patch.object(
        sys,
        "argv",
        ["cli.py", "--config", "test.json", "generate-part", "--width", "5.0", "--export", "pdf"],
    ):
        args = parse_args()
        assert args.config == "test.json"
        assert args.command == "generate-part"
        assert args.width == 5.0
        assert args.export == "pdf"
        assert args.verbose is False

    # Test sub-command parsing for process-bom
    with patch.object(
        sys, "argv", ["cli.py", "--verbose", "process-bom", "--bom-file", "bom.csv", "--inspect"]
    ):
        args = parse_args()
        assert args.command == "process-bom"
        assert args.bom_file == "bom.csv"
        assert args.verbose is True
        assert args.inspect is True

    # Test sub-command parsing for create-assembly
    with patch.object(
        sys,
        "argv",
        [
            "cli.py",
            "create-assembly",
            "--assembly-number",
            "ASM-101",
            "--components",
            "part1.ipt,part2.ipt",
        ],
    ):
        args = parse_args()
        assert args.command == "create-assembly"
        assert args.assembly_number == "ASM-101"
        assert args.components == "part1.ipt,part2.ipt"

    # Test sub-command parsing for inspect
    with patch.object(sys, "argv", ["cli.py", "inspect", "--file", "part.ipt"]):
        args = parse_args()
        assert args.command == "inspect"
        assert args.file == "part.ipt"

    # Test sub-command parsing for export-batch
    with patch.object(
        sys, "argv", ["cli.py", "export-batch", "--file", "part.ipt", "--formats", "step,stl"]
    ):
        args = parse_args()
        assert args.command == "export-batch"
        assert args.file == "part.ipt"
        assert args.formats == "step,stl"

    # Test generate-part with new primitives and options
    with patch.object(
        sys,
        "argv",
        [
            "cli.py",
            "generate-part",
            "--type",
            "cylinder",
            "--diameter",
            "50",
            "--length",
            "120",
            "--unit",
            "mm",
        ],
    ):
        args = parse_args()
        assert args.command == "generate-part"
        assert args.type == "cylinder"
        assert args.diameter == 50.0
        assert args.length == 120.0
        assert args.unit == "mm"
