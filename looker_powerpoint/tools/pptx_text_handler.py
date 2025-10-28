import logging
import re
from pptx import Presentation
from pptx.dml.color import RGBColor
from jinja2 import Environment, BaseLoader
import pandas as pd

# ---------- Emoji removal helper ----------
# Regex to match emoji and a broad set of pictographs/symbols.
_EMOJI_REGEX = re.compile(
    "["
    "\U0001f600-\U0001f64f"  # emoticons
    "\U0001f300-\U0001f5ff"  # symbols & pictographs
    "\U0001f680-\U0001f6ff"  # transport & map symbols
    "\U0001f1e0-\U0001f1ff"  # flags (iOS)
    "\U00002702-\U000027b0"  # dingbats
    "\U000024c2-\U0001f251"
    "\U0001f900-\U0001f9ff"  # Supplemental Symbols and Pictographs
    "\U0001fa70-\U0001faff"  # Symbols and Pictographs Extended-A
    "\U00002600-\U000026ff"  # Misc symbols
    "]+",
    flags=re.UNICODE,
)


def remove_emojis_from_string(s):
    if not isinstance(s, str):
        return s
    return _EMOJI_REGEX.sub("", s)


def sanitize_value(v):
    # Keep numbers, None, etc. For strings, remove emojis.
    if v is None:
        return v
    # pandas NA / numpy.nan: return as-is (Jinja will render as 'nan' if cast)
    try:
        import numpy as _np

        if v is _np.nan:
            return v
    except Exception:
        pass
    if isinstance(v, str):
        return remove_emojis_from_string(v)
    return v


_WS_RE = re.compile(r"\s+")


def sanitize_header_name(h):
    """
    Remove emojis, strip, and replace inner whitespace with underscores.
    """
    if h is None:
        return h
    # convert to str in case it's not already
    s = str(h)
    # remove emojis and trailing/leading spaces handled by that function
    s = remove_emojis_from_string(s)
    # collapse internal whitespace to single underscore
    s = _WS_RE.sub("_", s)
    # strip leading/trailing underscores produced by replacement of leading/trailing spaces
    s = s.strip("_")
    return s


def sanitize_dataframe_headers(df):
    """
    Return a new DataFrame with sanitized column headers:
    - emojis removed
    - internal whitespace replaced with underscores
    - leading/trailing underscores trimmed
    """
    # build rename mapping
    rename_map = {col: sanitize_header_name(col) for col in df.columns}
    # return a copy with renamed columns
    return df.rename(columns=rename_map)


# ---------- Marker encoding for colored segments ----------
_START = "\u0002"
_SEP = "\u0003"
_END = "\u0004"
MARKER_RE = re.compile(
    re.escape(_START)
    + r"(#[0-9A-Fa-f]{6})"
    + re.escape(_SEP)
    + r"(.*?)"
    + re.escape(_END),
    re.DOTALL,
)


def encode_colored_text(text, hex_color):
    return f"{_START}{hex_color}{_SEP}{text}{_END}"


def decode_marked_segments(rendered_text):
    segments = []
    pos = 0
    for m in MARKER_RE.finditer(rendered_text):
        if m.start() > pos:
            segments.append((rendered_text[pos : m.start()], None))
        hex_color = m.group(1)
        txt = m.group(2)
        segments.append((txt, hex_color))
        pos = m.end()
    if pos < len(rendered_text):
        segments.append((rendered_text[pos:], None))
    return segments


# ---------- Formatting copy helper ----------
def copy_run_format(src_run, dest_run):
    try:
        dest_run.font.bold = src_run.font.bold
        dest_run.font.italic = src_run.font.italic
        dest_run.font.underline = src_run.font.underline
        dest_run.font.strike = src_run.font.strike
        if src_run.font.name:
            dest_run.font.name = src_run.font.name
        if src_run.font.size:
            dest_run.font.size = src_run.font.size
        # Try to copy RGB color if available
        try:
            col = src_run.font.color
            if (
                getattr(col, "type", None) is not None
                and getattr(col, "rgb", None) is not None
            ):
                rgb = col.rgb
                try:
                    dest_run.font.color.rgb = RGBColor(rgb[0], rgb[1], rgb[2])
                except Exception:
                    try:
                        dest_run.font.color.rgb = rgb
                    except Exception:
                        pass
        except Exception:
            pass
    except Exception:
        pass


# ---------- Robust colorize_positive filter ----------
def colorize_positive(
    value, positive_hex="#008000", negative_hex="#C00000", zero_hex="#000000"
):
    """
    Try to parse 'value' robustly; return marker-wrapped text for coloring.
    """

    def try_parse_number(v):
        if v is None:
            return None
        if isinstance(v, (int, float)):
            try:
                return float(v)
            except Exception:
                return None
        # pandas NA / numpy.nan
        try:
            import numpy as _np

            if v is _np.nan:
                return None
        except Exception:
            pass
        if isinstance(v, str):
            s = v.strip()
            if s == "":
                return None
            # Remove emojis before parsing (safety)
            s = remove_emojis_from_string(s)
            s = s.replace(",", "").replace(" ", "")
            s = re.sub(r"^[^\d\-\+\.]+", "", s)
            s = re.sub(r"[^\d\.eE\-\+]$", "", s)
            try:
                return float(s)
            except Exception:
                m = re.match(r"^\(([\d\.,\-]+)\)$", v.strip())
                if m:
                    inner = m.group(1).replace(",", "")
                    try:
                        return -float(inner)
                    except Exception:
                        return None
                return None
        try:
            return float(v)
        except Exception:
            return None

    num = try_parse_number(value)
    if num is None:
        hex_color = zero_hex
    elif num > 0:
        hex_color = positive_hex
    elif num < 0:
        hex_color = negative_hex
    else:
        hex_color = zero_hex

    text = "" if value is None else str(value)
    # ensure the output text also has emojis removed (defensive)
    text = remove_emojis_from_string(text)
    return encode_colored_text(text, hex_color)


# ---------- Jinja env ----------
def make_jinja_env():
    env = Environment(loader=BaseLoader(), autoescape=False)
    env.filters["colorize_positive"] = colorize_positive
    return env


def render_text_with_jinja(text, context, env=None):
    if env is None:
        env = make_jinja_env()
    template = env.from_string(text)
    return template.render(**(context or {}))


# ---------- Extract original runs and text ----------
def extract_text_and_run_meta(text_frame):
    parts = []
    run_meta = []
    for p in text_frame.paragraphs:
        for r in p.runs:
            run_text = r.text or ""
            parts.append(run_text)
            run_meta.append({"text": run_text, "run_obj": r})
        parts.append("\n")
        run_meta.append({"text": "\n", "run_obj": None})
    if parts and parts[-1] == "\n":
        parts.pop()
        run_meta.pop()
    full_text = "".join(parts)
    return full_text, run_meta


# ---------- Reinsert rendered text, applying color markers ----------
def reinsert_rendered_text(text_frame, rendered_text, original_run_meta):
    segments = decode_marked_segments(rendered_text)
    text_frame.text = ""
    orig_runs = [rm["run_obj"] for rm in original_run_meta]
    if not orig_runs:
        p = (
            text_frame.paragraphs[0]
            if text_frame.paragraphs
            else text_frame.add_paragraph()
        )
        for seg_text, seg_color in segments:
            lines = seg_text.split("\n")
            for i_line, line_part in enumerate(lines):
                if i_line > 0:
                    p = text_frame.add_paragraph()
                if line_part == "":
                    continue
                r = p.add_run()
                r.text = line_part
                if seg_color:
                    try:
                        rgb = seg_color.lstrip("#")
                        r.font.color.rgb = RGBColor(
                            int(rgb[0:2], 16), int(rgb[2:4], 16), int(rgb[4:6], 16)
                        )
                    except Exception:
                        pass
        return

    cur_para = (
        text_frame.paragraphs[0]
        if text_frame.paragraphs
        else text_frame.add_paragraph()
    )
    for p in text_frame.paragraphs:
        p.clear()

    template_idx = 0
    for seg_text, seg_color in segments:
        lines = seg_text.split("\n")
        for i_l, lp in enumerate(lines):
            if i_l > 0:
                cur_para = text_frame.add_paragraph()
                try:
                    cur_para.alignment = text_frame.paragraphs[0].alignment
                except Exception:
                    pass
            if lp == "":
                continue
            r = cur_para.add_run()
            r.text = lp
            attempts = 0
            while attempts < len(orig_runs) and orig_runs[template_idx] is None:
                template_idx = (template_idx + 1) % len(orig_runs)
                attempts += 1
            src_run = orig_runs[template_idx]
            template_idx = (template_idx + 1) % len(orig_runs)
            if src_run is not None:
                try:
                    copy_run_format(src_run, r)
                except Exception:
                    pass
            if seg_color:
                try:
                    rgb = seg_color.lstrip("#")
                    r.font.color.rgb = RGBColor(
                        int(rgb[0:2], 16), int(rgb[2:4], 16), int(rgb[4:6], 16)
                    )
                except Exception:
                    pass


# ---------- High-level processor ----------
def process_text_field(shape, text_to_insert, df, env=None):
    jinja_tag_re = re.compile(r"({{.*?}}|{%.+?%})", re.DOTALL)
    text_frame = shape.text_frame
    full_text, run_meta = extract_text_and_run_meta(text_frame)

    if not jinja_tag_re.search(full_text):
        logging.debug("No Jinja tags found in shape; applying fallback if different.")
        if full_text != (text_to_insert or ""):
            shape.text = text_to_insert or ""
        return

    df_sanitized = sanitize_dataframe_headers(df)

    # 2) convert to list-of-dicts and also sanitize row string values if needed
    rows = df_sanitized.to_dict(orient="records")

    # Sanitize string values in each row
    # Build sanitized rows and pass to Jinja as 'rows'
    context = {"rows": rows}

    # try:
    rendered = render_text_with_jinja(full_text, context, env=env)
    # except Exception as e:
    # raise RuntimeError(f"Jinja rendering failed: {e}")

    reinsert_rendered_text(text_frame, rendered, run_meta)
