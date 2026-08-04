import pytest
from src.cli import parse_args

def test_cli_help(capsys):
    # Test that parsing with --help displays instructions and raises SystemExit
    with pytest.raises(SystemExit) as e:
        import sys
        from unittest.mock import patch
        with patch.object(sys, 'argv', ['cli.py', '--help']):
            parse_args()
    assert e.value.code == 0

def test_cli_arguments():
    import sys
    from unittest.mock import patch
    
    # Test sub-command parsing for generate-part
    with patch.object(
        sys, 'argv', 
        ['cli.py', '--config', 'test.json', 'generate-part', '--width', '5.0', '--export', 'pdf']
    ):
        args = parse_args()
        assert args.config == "test.json"
        assert args.command == "generate-part"
        assert args.width == 5.0
        assert args.export == "pdf"
        assert args.verbose is False

    # Test sub-command parsing for process-bom
    with patch.object(
        sys, 'argv', 
        ['cli.py', '--verbose', 'process-bom', '--bom-file', 'bom.csv']
    ):
        args = parse_args()
        assert args.command == "process-bom"
        assert args.bom_file == "bom.csv"
        assert args.verbose is True
