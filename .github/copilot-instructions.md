# Looker PowerPoint CLI - Copilot Instructions

## Architecture Overview

This CLI tool integrates Looker dashboard data with PowerPoint presentations by embedding YAML metadata in shape alternative text and replacing shapes with live data.

### Core Data Flow
1. **Shape Discovery**: `looker_powerpoint/tools/find_alt_text.py` extracts YAML metadata from PowerPoint shape alternative text using XPath queries on PowerPoint XML
2. **Validation**: Pydantic models in `models.py` validate and transform shape metadata into `LookerShape` and `LookerReference` objects
3. **Data Fetching**: `LookerClient` asynchronously fetches data from Looker API using the looker-sdk
4. **Shape Replacement**: CLI updates PowerPoint shapes in-place based on shape type (PICTURE, TABLE, TEXT_BOX, CHART, AUTO_SHAPE)

### Key Components

- **`cli.py`**: Main CLI orchestrator with Rich logging and argparse. Entry point is `lppt` command
- **`models.py`**: Pydantic models with shape-type-specific validation logic in `push_down_relevant_data()` 
- **`looker.py`**: Async Looker API client wrapper around looker-sdk
- **`tools/find_alt_text.py`**: PowerPoint XML parsing for metadata extraction

## Critical Patterns

### Shape Metadata Format
Shapes store integration config as YAML in alternative text:
```yaml
id: "123"
id_type: "look"  # or "meta"
result_format: "json_bi"  # or "jpg", "png"
label: "column_name"  # for text extraction
filter: "dimension_name"  # for filtering
```

### Shape Type Handling
- **PICTURE**: Auto-converts to image format, sets image dimensions from shape size
- **TABLE**: Enables formatting, renames columns using Looker metadata field mappings
- **TEXT_BOX/AUTO_SHAPE**: Extracts single values using `label` field or falls back to full data string
- **CHART**: Parses JSON data into CategoryChartData for chart replacement

### Error Handling Convention
Failed shapes get red circle overlays (via `_mark_failure()`) unless `--hide-errors` flag is used. Enables visual debugging in presentations.

### XML Manipulation Pattern
All PowerPoint shape updates use direct XML manipulation via lxml:
```python
# Standard XPath patterns for shape properties
".//p:nvSpPr/p:cNvPr"      # Shape properties
".//p:nvPicPr/p:cNvPr"     # Picture properties  
".//p:nvGraphicFramePr/p:cNvPr"  # Table/chart properties
```

## Development Workflows

### Environment Setup
```bash
uv sync  # Install all dependencies including dev tools
```

### Key Commands
- `uv run lppt -f file.pptx` - Process PowerPoint file
- `uv run lppt --help` - See all CLI options
- `cd docs && make html` - Build Sphinx documentation (uses uv via Makefile)

### Required Environment Variables
```bash
LOOKERSDK_BASE_URL=https://your-looker.com
LOOKERSDK_CLIENT_ID=your_client_id  
LOOKERSDK_CLIENT_SECRET=your_secret
```

### Testing Pattern
Debug JSON files are automatically written during processing for inspection: `debug_{shape_id}.json`

## Project Conventions

- **uv package manager**: All dependency management and script execution via uv
- **Pydantic v2**: All data validation with field validators and model validators
- **Async/await**: Looker API calls are batched asynchronously via `asyncio.gather()`
- **Rich logging**: Structured console output with progress indicators
- **Shape identification**: Uses `shape_id` for targeting, `shape_number` for iteration
- **In-place updates**: Default behavior modifies original files unless `--self` flag prevents it

## Integration Points

- **Looker SDK**: Direct API integration via `looker_sdk.init40()`
- **python-pptx**: PowerPoint file manipulation and shape access
- **pandas**: Data transformation for tables and charts
- **lxml**: XML parsing for alternative text metadata
- **Rich**: Console UI and argument parsing with `RichHelpFormatter`

## Meta Shapes
Special shapes with `meta: true` are removed from output presentations (cleanup markers for development).