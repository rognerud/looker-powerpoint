import json
from unittest.mock import patch
from looker_powerpoint.cli import Cli


def test_default_output_dir():
    """Test that the default output directory is 'output'."""
    with patch("os.getenv", return_value="dummy_value"):
        cli = Cli()
        args = cli.parser.parse_args()
        assert args.output_dir == "output"


def _make_cli():
    with patch("os.getenv", return_value="dummy_value"):
        return Cli()


def _field(name):
    """Build a field dict with a field_group_variant derived from the short name."""
    return {"name": name, "field_group_variant": name.split(".")[-1]}


def _make_result(dimensions, measures, table_calculations, rows, custom_sorts=None, custom_pivots=None):
    """Build a json_bi-style result string for _make_df.

    Each entry in *dimensions*, *measures*, and *table_calculations* may be either a
    plain field name string or a dict with ``name``/``field_group_variant`` keys.
    Plain strings are automatically expanded with a ``field_group_variant`` equal to
    the portion of the name after the last dot (e.g. ``"view.date"`` → ``"date"``).
    """
    def _normalise(fields):
        return [_field(f) if isinstance(f, str) else f for f in fields]

    return json.dumps({
        "metadata": {
            "fields": {
                "dimensions": _normalise(dimensions),
                "measures": _normalise(measures),
                "table_calculations": _normalise(table_calculations or []),
            }
        },
        "rows": rows,
        "custom_sorts": custom_sorts or [],
        "custom_pivots": custom_pivots or [],
    })


class TestMakeDf:
    """Tests for Cli._make_df column ordering logic."""

    def test_no_pivots_column_order(self):
        """Dimensions come first, then measures, then table calcs — no pivots."""
        cli = _make_cli()
        result = _make_result(
            dimensions=["view.date", "view.name"],
            measures=["view.revenue", "view.cost"],
            table_calculations=["calc1"],
            rows=[
                {
                    "view.date.value": "2024-01-01",
                    "view.name.value": "A",
                    "view.revenue.value": 100,
                    "view.cost.value": 50,
                    "calc1.value": 2.0,
                }
            ],
        )
        df = cli._make_df(result)
        cols = list(df.columns)
        # All five columns must be present (renamed via field_group_variant)
        assert len(cols) == 5
        # Dimensions first
        dim_positions = [cols.index("date"), cols.index("name")]
        # Measures next
        measure_positions = [cols.index("revenue"), cols.index("cost")]
        # Calc last
        calc_position = cols.index("calc1")
        assert max(dim_positions) < min(measure_positions), "Dims must precede measures"
        assert max(measure_positions) < calc_position, "Measures must precede table calcs"

    def test_no_pivots_preserves_native_dimension_order(self):
        """Dimension order follows the metadata field order, not data column order."""
        cli = _make_cli()
        result = _make_result(
            dimensions=["view.z_dim", "view.a_dim"],
            measures=["view.measure"],
            table_calculations=[],
            rows=[
                {
                    # a_dim appears first in the data dict on purpose
                    "view.a_dim.value": "foo",
                    "view.z_dim.value": "bar",
                    "view.measure.value": 1,
                }
            ],
        )
        df = cli._make_df(result)
        cols = list(df.columns)
        # z_dim was listed first in metadata → must appear before a_dim
        assert cols.index("z_dim") < cols.index("a_dim")

    def test_pivoted_measures_ascending(self):
        """Pivot columns are kept in their Looker-native (ascending) appearance order."""
        cli = _make_cli()
        # Looker returns pivot values in month order: Jan, Feb, Mar
        result = _make_result(
            dimensions=["view.region"],
            measures=["view.revenue"],
            table_calculations=[],
            rows=[
                {
                    "view.region.value": "North",
                    "view.revenue|FIELD|2024-01.value": 10,
                    "view.revenue|FIELD|2024-02.value": 20,
                    "view.revenue|FIELD|2024-03.value": 30,
                }
            ],
            custom_pivots=["view.month"],
            custom_sorts=["view.month asc 0"],
        )
        df = cli._make_df(result)
        pivot_cols = [c for c in df.columns if "revenue" in c]
        # The pivot order should match Looker's data order: Jan → Feb → Mar
        assert pivot_cols == [
            "view.revenue|FIELD|2024-01.value",
            "view.revenue|FIELD|2024-02.value",
            "view.revenue|FIELD|2024-03.value",
        ]

    def test_pivoted_measures_descending(self):
        """When the primary pivot is sorted desc, pivot column order is reversed."""
        cli = _make_cli()
        # Looker returns pivot values in month order: Jan, Feb, Mar
        # but sorts descending → we expect Mar, Feb, Jan
        result = _make_result(
            dimensions=["view.region"],
            measures=["view.revenue"],
            table_calculations=[],
            rows=[
                {
                    "view.region.value": "North",
                    "view.revenue|FIELD|2024-01.value": 10,
                    "view.revenue|FIELD|2024-02.value": 20,
                    "view.revenue|FIELD|2024-03.value": 30,
                }
            ],
            custom_pivots=["view.month"],
            custom_sorts=["view.month desc 0"],
        )
        df = cli._make_df(result)
        pivot_cols = [c for c in df.columns if "revenue" in c]
        # Descending → reversed from appearance order
        assert pivot_cols == [
            "view.revenue|FIELD|2024-03.value",
            "view.revenue|FIELD|2024-02.value",
            "view.revenue|FIELD|2024-01.value",
        ]

    def test_numeric_pivot_values_use_appearance_order_not_lexicographic(self):
        """
        Numeric-like pivot values ("2", "10") must be ordered by their appearance in
        Looker data, not lexicographically (which would give "10" before "2").
        """
        cli = _make_cli()
        # Looker returns data in numeric order: week 2, week 10
        result = _make_result(
            dimensions=["view.region"],
            measures=["view.revenue"],
            table_calculations=[],
            rows=[
                {
                    "view.region.value": "East",
                    "view.revenue|FIELD|2.value": 200,
                    "view.revenue|FIELD|10.value": 1000,
                }
            ],
            custom_pivots=["view.week"],
            custom_sorts=["view.week asc 0"],
        )
        df = cli._make_df(result)
        pivot_cols = [c for c in df.columns if "revenue" in c]
        # Should follow Looker's data order: 2 then 10
        # (lexicographic sort would give "10" before "2", which is wrong)
        assert pivot_cols == [
            "view.revenue|FIELD|2.value",
            "view.revenue|FIELD|10.value",
        ]

    def test_pivot_sort_only_reverses_on_exact_field_match(self):
        """
        A sort field that is a *substring* of the pivot field name must NOT trigger
        pivot_descending — only an exact token match should.
        """
        cli = _make_cli()
        # The sort field "view.month" is a substring of "view.month_long",
        # but only "view.month_long" is the actual pivot.
        result = _make_result(
            dimensions=["view.region"],
            measures=["view.revenue"],
            table_calculations=[],
            rows=[
                {
                    "view.region.value": "West",
                    "view.revenue|FIELD|Jan.value": 1,
                    "view.revenue|FIELD|Feb.value": 2,
                }
            ],
            custom_pivots=["view.month_long"],
            # "view.month" desc should NOT affect ordering of "view.month_long" pivot
            custom_sorts=["view.month desc 0"],
        )
        df = cli._make_df(result)
        pivot_cols = [c for c in df.columns if "revenue" in c]
        # No pivot_descending should be triggered → appearance order preserved
        assert pivot_cols == [
            "view.revenue|FIELD|Jan.value",
            "view.revenue|FIELD|Feb.value",
        ]
