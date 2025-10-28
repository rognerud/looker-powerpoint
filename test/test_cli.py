from unittest.mock import patch
from looker_powerpoint.cli import Cli


def test_default_output_dir():
    """Test that the default output directory is 'output'."""
    with patch("os.getenv", return_value="dummy_value"):
        cli = Cli()
        args = cli.parser.parse_args()
        assert args.output_dir == "output"
