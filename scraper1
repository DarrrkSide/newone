#!/usr/bin/env python3
"""
BCA/PRT Scoresheet Scanner
===========================
Reads scanned (handwritten) NAVPERS 6110/10 (BCA) and NAVPERS 6110/11 (PRT)
scoresheets and extracts the data fields into a single Excel workbook.

Because the sheets are filled out by hand, this uses a vision-capable AI
model (not plain OCR) to read each page, since handwriting recognition
needs judgment/context that traditional OCR engines get wrong.

Two provider options - pick whichever you have access to:

  --provider gemini    FREE. Uses Google's Gemini API, which has a
                        no-cost tier (no credit card required, just a
                        Google account + API key from
                        https://aistudio.google.com/apikey). This is the
                        default. Free-tier requests are rate limited
                        (a handful of pages per minute), so large batches
                        will take a while - the script paces requests
                        automatically. Note: Google's free tier terms
                        allow them to use your inputs to improve their
                        models, so don't use it on sensitive/PII-heavy
                        scans if that's a concern for you - use --provider
                        anthropic instead in that case (paid, not used
                        for training).

  --provider anthropic  Paid, pay-as-you-go. Uses Claude's vision API.
                        Needs an ANTHROPIC_API_KEY from
                        https://console.anthropic.com. Not free, but
                        usually just a few dollars for a large batch of
                        scans, and Anthropic does not train on API data.

SETUP
-----
1. Install dependencies:
     pip install pdf2image openpyxl pillow
     pip install google-genai        # if using --provider gemini (default)
     pip install anthropic           # if using --provider anthropic

   pdf2image also requires the `poppler` utilities to be installed on your
   system (provides `pdftoppm`):
     - Mac:     brew install poppler
     - Ubuntu:  sudo apt-get install poppler-utils
     - Windows: https://github.com/oschwartz10612/poppler-windows

2. Set an API key as an environment variable:
     export GEMINI_API_KEY=...       # free key from https://aistudio.google.com/apikey
   or
     export ANTHROPIC_API_KEY=sk-ant-...   # from https://console.anthropic.com

USAGE
-----
Put all your scanned files (PDF, JPG, or PNG) in one folder, then run:

     python scan_scoresheets.py --input ./scans --output results.xlsx
     python scan_scoresheets.py --input ./scans --output results.xlsx --provider anthropic
     python scan_scoresheets.py --input ./scans --output results.xlsx --passes 3

--passes N (default 1) re-reads each page N times instead of once. Any
field that comes back empty on the first read but gets filled in on a
later attempt is highlighted YELLOW in the output spreadsheet, so you
know exactly which cells were "hard to read" and deserve a manual check
against the original scan. Each extra pass is a full extra API call per
page (more time on the free Gemini tier, a bit more cost on Anthropic).

Each PDF page or image is treated as one scanned form page. A single BCA
form is one person. A single PRT form is a roster and can hold several
people. The script matches BCA and PRT records for the same person by
name and merges them into one row per person in the output spreadsheet.

Notes / limitations
--------------------
- If a name is spelled slightly differently between a person's BCA sheet
  and their row on the PRT roster, they may not auto-match. Check the
  "Needs Review" column in the output.
- The model is instructed to leave a field blank rather than guess when
  a handwritten value is illegible, and to flag anything it is unsure about
  in the Needs Review notes. Always spot check the output against the
  original scans before using it officially.
"""

import argparse
import base64
import io
import json
import os
import re
import sys
import time
from dataclasses import dataclass, field
from pathlib import Path

try:
    from pdf2image import convert_from_path
except ImportError:
    sys.exit("Missing dependency. Run: pip install pdf2image")

try:
    from PIL import Image
except ImportError:
    sys.exit("Missing dependency. Run: pip install pillow")

from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter

ANTHROPIC_MODEL = "claude-sonnet-4-6"
GEMINI_MODEL = "gemini-2.5-flash"  # free-tier eligible as of mid-2026; change if Google renames/retires it
DPI = 300  # scan resolution when rasterizing PDF pages; higher = more legible handwriting, slower
MAX_IMAGE_DIM = 1568  # downscale longest edge to this before sending to the API
GEMINI_FREE_TIER_DELAY_SEC = 4.5  # pace requests to stay under free-tier RPM caps

# -----------------------------------------------------------------------
# Output columns, in the order they'll appear in the spreadsheet.
# These map directly to the rows in expected_values.xlsx.
# -----------------------------------------------------------------------
OUTPUT_COLUMNS = [
    "UIC",
    "Rate/Rank",
    "Name (Last, First)",
    "DoD ID",
    "Sex",
    "Date of Birth",
    "BCA Date",
    "Height (Rounded)",
    "Weight (Rounded)",
    "Waist (Average)",
    "PRT Date",
    "Cardio Modality",
    "Number of Push Ups",
    "Plank Time (M:SS)",
    "Calories (if bike chosen)",
    "Cardio Time (MM:SS)",
    "Needs Review",
    "Source File(s)",
]

# -----------------------------------------------------------------------
# The extraction prompt. Field formats are taken directly from the
# expected_values.xlsx reference the form owner supplied, so the model
# knows what a plausible reading looks like for each field.
# -----------------------------------------------------------------------
EXTRACTION_PROMPT = """You are transcribing a single scanned page from a U.S. Navy \
Body Composition Assessment (BCA, NAVPERS 6110/10) or Physical Readiness Test \
(PRT, NAVPERS 6110/11) scoresheet. All entries you need to read are HANDWRITTEN. \
Printed form labels are NOT the data - ignore them except as context for where \
handwritten values were entered.

First, decide which kind of page this is:
- "BCA": the Body Composition Assessment sheet. It holds data for ONE person \
(height, weight, waist, WHtR, body fat%, DoD ID, DOB, etc).
- "PRT": the Physical Readiness Test sheet. It is a ROSTER TABLE that can hold \
MULTIPLE people (one per row): push ups, plank time, cardio modality/time, etc.
- "OTHER": the page has no relevant handwritten data (e.g. blank, or an \
unrelated page).

Then extract the handwritten values using these format expectations to help \
you interpret unclear handwriting (a value outside the plausible range is a \
strong signal you may be misreading it - double check before committing to it):

- UIC: 5-digit number
- Rate/Rank: alphanumeric Navy rate/rank, e.g. "E1"/"E-1" through "E8"/"E-8", \
"O1"/"O-1" through "O6"/"O-6", "SR", "SA", "SN", "PO1", "PO2", "PO3", "CPO", \
"YN1"/"YN2"/"YN3" (rating prefix + number), "YNC" (rating prefix + "C"), \
"YNSC" (rating prefix + "SC"), "ENS", "LTJG", "LT", "LCDR", "CDR", "CAPT". The \
letters before the number/suffix are the rating abbreviation and vary by person.
- Name: "Last, First, M.I." format, alphabetic, may include hyphens
- DoD ID: 10-digit number
- Sex: "M" or "F"
- Date of Birth / BCA Date / PRT Date: dates, various formats \
(DD/MM/YY, DD/MM/YYYY, DD Month YY, DD Month YYYY - month may be abbreviated). \
Transcribe the date as written; do not reformat it.
- Height (Rounded): 50-80, in 0.5 increments (e.g. 56.5)
- Weight (Rounded): 97-325, whole number
- Waist (Average): 20-50, may have decimals to the thousandths (e.g. 43.125)
- Cardio Modality: Run, Row, Bike, Swim, Tread/Treadmill
- Number of Push Ups: 1-100, whole number
- Plank Time: M:SS format, between :01 and 3:30
- Calories (only present if Bike was the chosen cardio modality): 50-300, whole number
- Cardio Time: MM:SS format, between :01 and 20:00 (not applicable if Bike was chosen; use Calories instead)

Respond with ONLY a single JSON object (no markdown fences, no commentary), \
exactly matching one of these shapes:

If form_type is "BCA":
{
  "form_type": "BCA",
  "uic": "" ,
  "rate_rank": "",
  "name": "",
  "dod_id": "",
  "sex": "",
  "dob": "",
  "bca_date": "",
  "height_rounded": "",
  "weight_rounded": "",
  "waist_average": "",
  "prt_cardio_choice": "",
  "notes": ""
}

If form_type is "PRT":
{
  "form_type": "PRT",
  "prt_date": "",
  "uic": "",
  "entries": [
    {
      "name": "",
      "sex": "",
      "rank_rate": "",
      "cardio_modality": "",
      "push_ups": "",
      "plank_time": "",
      "cardio_time_or_calories": "",
      "notes": ""
    }
  ]
}

If form_type is "OTHER":
{ "form_type": "OTHER" }

Rules:
- If a field is illegible, blank, or not present on the page, use an empty \
string "" for it - never guess or invent a value.
- Use "notes" to flag anything uncertain, e.g. "waist reading could be 43.1 \
or 45.1 - handwriting ambiguous" or "name partially obscured". Leave "notes" \
as "" if nothing needs flagging.
- For a PRT roster, include one entry per row that has any handwritten data. \
Skip empty rows.
- Transcribe names exactly as written (Last, First, M.I.) so they can be \
matched against the same person's BCA sheet.
"""


def encode_image(img: Image.Image) -> str:
    if img.mode != "RGB":
        img = img.convert("RGB")
    w, h = img.size
    longest = max(w, h)
    if longest > MAX_IMAGE_DIM:
        scale = MAX_IMAGE_DIM / longest
        img = img.resize((int(w * scale), int(h * scale)), Image.LANCZOS)
    buf = io.BytesIO()
    img.save(buf, format="JPEG", quality=92)
    return base64.standard_b64encode(buf.getvalue()).decode("utf-8")


def pages_from_file(path: Path, poppler_path: str = None):
    """Yield (page_label, PIL.Image) for each page/image in a file."""
    suffix = path.suffix.lower()
    if suffix == ".pdf":
        images = convert_from_path(str(path), dpi=DPI, poppler_path=poppler_path)
        for i, img in enumerate(images, start=1):
            yield f"{path.name} (page {i})", img
    elif suffix in (".png", ".jpg", ".jpeg", ".tif", ".tiff", ".bmp"):
        yield path.name, Image.open(path)
    else:
        return


def _clean_json_text(text: str) -> dict:
    text = re.sub(r"^```(json)?|```$", "", text.strip(), flags=re.MULTILINE).strip()
    return json.loads(text)


def call_anthropic(client, img: Image.Image, source_label: str) -> dict:
    b64 = encode_image(img)
    try:
        resp = client.messages.create(
            model=ANTHROPIC_MODEL,
            max_tokens=2000,
            messages=[{
                "role": "user",
                "content": [
                    {"type": "image", "source": {"type": "base64", "media_type": "image/jpeg", "data": b64}},
                    {"type": "text", "text": EXTRACTION_PROMPT},
                ],
            }],
        )
    except Exception as e:
        print(f"  ! API error on {source_label}: {e}", file=sys.stderr)
        return {"form_type": "ERROR", "notes": str(e)}

    text = "".join(block.text for block in resp.content if block.type == "text").strip()
    try:
        return _clean_json_text(text)
    except json.JSONDecodeError:
        print(f"  ! Could not parse model response for {source_label}:\n{text[:500]}", file=sys.stderr)
        return {"form_type": "ERROR", "notes": "unparsable model response"}


def call_gemini(client, img: Image.Image, source_label: str) -> dict:
    from google.genai import types

    if img.mode != "RGB":
        img = img.convert("RGB")
    w, h = img.size
    longest = max(w, h)
    if longest > MAX_IMAGE_DIM:
        scale = MAX_IMAGE_DIM / longest
        img = img.resize((int(w * scale), int(h * scale)), Image.LANCZOS)
    buf = io.BytesIO()
    img.save(buf, format="JPEG", quality=92)

    try:
        resp = client.models.generate_content(
            model=GEMINI_MODEL,
            contents=[
                EXTRACTION_PROMPT,
                types.Part.from_bytes(data=buf.getvalue(), mime_type="image/jpeg"),
            ],
        )
    except Exception as e:
        print(f"  ! API error on {source_label}: {e}", file=sys.stderr)
        return {"form_type": "ERROR", "notes": str(e)}
    finally:
        time.sleep(GEMINI_FREE_TIER_DELAY_SEC)

    text = (resp.text or "").strip()
    try:
        return _clean_json_text(text)
    except json.JSONDecodeError:
        print(f"  ! Could not parse model response for {source_label}:\n{text[:500]}", file=sys.stderr)
        return {"form_type": "ERROR", "notes": "unparsable model response"}


def norm_name(name: str) -> str:
    """Loose key for matching the same person across BCA and PRT sheets."""
    return re.sub(r"[^a-z]", "", name.lower())


@dataclass
class Person:
    uic: str = ""
    rate_rank: str = ""
    name: str = ""
    dod_id: str = ""
    sex: str = ""
    dob: str = ""
    bca_date: str = ""
    height_rounded: str = ""
    weight_rounded: str = ""
    waist_average: str = ""
    prt_date: str = ""
    cardio_modality: str = ""
    push_ups: str = ""
    plank_time: str = ""
    calories: str = ""
    cardio_time: str = ""
    notes: list = field(default_factory=list)
    sources: list = field(default_factory=list)
    retry_highlight: set = field(default_factory=set)  # output column names filled only on a retry pass

    def _set_if_empty(self, attr: str, value: str, output_col: str, is_retry: bool):
        if value and not getattr(self, attr):
            setattr(self, attr, value)
            if is_retry:
                self.retry_highlight.add(output_col)

    def merge_bca(self, d: dict, retry_fields: set, source: str):
        self._set_if_empty("uic", d.get("uic", ""), "UIC", "uic" in retry_fields)
        self._set_if_empty("rate_rank", d.get("rate_rank", ""), "Rate/Rank", "rate_rank" in retry_fields)
        self._set_if_empty("name", d.get("name", ""), "Name (Last, First)", "name" in retry_fields)
        self._set_if_empty("dod_id", d.get("dod_id", ""), "DoD ID", "dod_id" in retry_fields)
        self._set_if_empty("sex", d.get("sex", ""), "Sex", "sex" in retry_fields)
        self._set_if_empty("dob", d.get("dob", ""), "Date of Birth", "dob" in retry_fields)
        self._set_if_empty("bca_date", d.get("bca_date", ""), "BCA Date", "bca_date" in retry_fields)
        self._set_if_empty("height_rounded", d.get("height_rounded", ""), "Height (Rounded)", "height_rounded" in retry_fields)
        self._set_if_empty("weight_rounded", d.get("weight_rounded", ""), "Weight (Rounded)", "weight_rounded" in retry_fields)
        self._set_if_empty("waist_average", d.get("waist_average", ""), "Waist (Average)", "waist_average" in retry_fields)
        self._set_if_empty("cardio_modality", d.get("prt_cardio_choice", ""), "Cardio Modality", "prt_cardio_choice" in retry_fields)
        if d.get("notes"):
            self.notes.append(f"[BCA] {d['notes']}")
        self.sources.append(source)

    def merge_prt(self, d: dict, retry_fields: set, prt_date: str, uic: str,
                  prt_date_is_retry: bool, uic_is_retry: bool, source: str):
        self._set_if_empty("uic", uic, "UIC", uic_is_retry)
        self._set_if_empty("rate_rank", d.get("rank_rate", ""), "Rate/Rank", "rank_rate" in retry_fields)
        self._set_if_empty("name", d.get("name", ""), "Name (Last, First)", "name" in retry_fields)
        self._set_if_empty("sex", d.get("sex", ""), "Sex", "sex" in retry_fields)
        self._set_if_empty("prt_date", prt_date, "PRT Date", prt_date_is_retry)
        self._set_if_empty("cardio_modality", d.get("cardio_modality", ""), "Cardio Modality", "cardio_modality" in retry_fields)
        self._set_if_empty("push_ups", d.get("push_ups", ""), "Number of Push Ups", "push_ups" in retry_fields)
        self._set_if_empty("plank_time", d.get("plank_time", ""), "Plank Time (M:SS)", "plank_time" in retry_fields)
        cardio_val = d.get("cardio_time_or_calories", "")
        if self.cardio_modality.strip().lower().startswith("bike"):
            self._set_if_empty("calories", cardio_val, "Calories (if bike chosen)", "cardio_time_or_calories" in retry_fields)
        else:
            self._set_if_empty("cardio_time", cardio_val, "Cardio Time (MM:SS)", "cardio_time_or_calories" in retry_fields)
        if d.get("notes"):
            self.notes.append(f"[PRT] {d['notes']}")
        self.sources.append(source)


# -----------------------------------------------------------------------
# Multi-pass merging: run the model over the same page N times and take
# the first non-empty answer for each field, tracking which fields were
# ONLY filled on a retry (pass 2+) so those cells can be highlighted.
# -----------------------------------------------------------------------
BCA_FIELDS = ["uic", "rate_rank", "name", "dod_id", "sex", "dob", "bca_date",
              "height_rounded", "weight_rounded", "waist_average", "prt_cardio_choice"]
PRT_ENTRY_FIELDS = ["name", "sex", "rank_rate", "cardio_modality", "push_ups",
                     "plank_time", "cardio_time_or_calories"]


def merge_page_passes_bca(pass_results: list) -> tuple:
    merged = {}
    retry_fields = set()
    notes = []
    for i, r in enumerate(pass_results):
        if r.get("form_type") != "BCA":
            continue
        for f in BCA_FIELDS:
            val = (r.get(f) or "").strip()
            if val and not merged.get(f):
                merged[f] = val
                if i > 0:
                    retry_fields.add(f)
        if r.get("notes"):
            notes.append(r["notes"])
    merged["notes"] = "; ".join(notes)
    return merged, retry_fields


def merge_page_passes_prt(pass_results: list):
    prt_date, uic = "", ""
    prt_date_retry, uic_retry = False, False
    entries: dict = {}  # norm_name -> {"fields": {}, "retry_fields": set(), "notes": [], "name": str}

    for i, r in enumerate(pass_results):
        if r.get("form_type") != "PRT":
            continue
        pd = (r.get("prt_date") or "").strip()
        if pd and not prt_date:
            prt_date = pd
            prt_date_retry = i > 0
        u = (r.get("uic") or "").strip()
        if u and not uic:
            uic = u
            uic_retry = i > 0

        for e in r.get("entries", []):
            key = norm_name(e.get("name", ""))
            if not key:
                continue
            slot = entries.setdefault(key, {"fields": {}, "retry_fields": set(), "notes": [], "name": ""})
            for f in PRT_ENTRY_FIELDS:
                val = (e.get(f) or "").strip()
                if val and not slot["fields"].get(f):
                    slot["fields"][f] = val
                    if i > 0:
                        slot["retry_fields"].add(f)
            if not slot["name"]:
                slot["name"] = e.get("name", "")
            if e.get("notes"):
                slot["notes"].append(e["notes"])

    return prt_date, uic, prt_date_retry, uic_retry, entries


def process_folder(input_dir: Path, client, provider: str, poppler_path: str = None, passes: int = 1):
    people: list = []  # list of Person objects, one per page
    files = sorted(
        p for p in input_dir.iterdir()
        if p.suffix.lower() in (".pdf", ".png", ".jpg", ".jpeg", ".tif", ".tiff", ".bmp")
    )
    if not files:
        sys.exit(f"No PDF/image files found in {input_dir}")

    for path in files:
        for label, img in pages_from_file(path, poppler_path=poppler_path):
            # For each page/image, run passes
            pass_results = []
            for p in range(passes):
                tag = f"{label} (pass {p + 1}/{passes})" if passes > 1 else label
                print(f"Reading {tag} ...")
                pass_results.append(call_model(client, img, tag))
            
            form_type = next(
                (r.get("form_type") for r in pass_results if r.get("form_type") in ("BCA", "PRT")), None
            )

            if form_type == "BCA":
                merged, retry_fields = merge_page_passes_bca(pass_results)
                person = Person()
                person.merge_bca(merged, retry_fields, label)
                people.append(person)

            elif form_type == "PRT":
                prt_date, uic, prt_date_retry, uic_retry, entries = merge_page_passes_prt(pass_results)
                for key, slot in entries.items():
                    if not key:
                        continue
                    person = Person()
                    entry_dict = dict(slot["fields"])
                    entry_dict["name"] = slot["name"]
                    entry_dict["notes"] = "; ".join(slot["notes"])
                    person.merge_prt(entry_dict, slot["retry_fields"], prt_date, uic, prt_date_retry, uic_retry, label)
                    people.append(person)

            else:
                errs = [r.get("notes") for r in pass_results if r.get("form_type") == "ERROR"]
                if errs:
                    print(f"  ! skipped {label}: {errs[0]}", file=sys.stderr)
            # OTHER -> ignore silently

    return people


def write_workbook(people, output_path: Path):
    wb = Workbook()
    ws = wb.active
    ws.title = "BCA-PRT Results"

    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill("solid", start_color="1F4E78")
    for col, title in enumerate(OUTPUT_COLUMNS, start=1):
        cell = ws.cell(row=1, column=col, value=title)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    review_fill = PatternFill("solid", start_color="FFF2CC")
    retry_fill = PatternFill("solid", start_color="FFFF00")

    for r, person in enumerate(sorted(people, key=lambda p: p.name), start=2):
        row_vals = [
            person.uic,
            person.rate_rank,
            person.name,
            person.dod_id,
            person.sex,
            person.dob,
            person.bca_date,
            person.height_rounded,
            person.weight_rounded,
            person.waist_average,
            person.prt_date,
            person.cardio_modality,
            person.push_ups,
            person.plank_time,
            person.calories,
            person.cardio_time,
            " | ".join(person.notes),
            "; ".join(dict.fromkeys(person.sources)),
        ]
        for c, val in enumerate(row_vals, start=1):
            cell = ws.cell(row=r, column=c, value=val)
            col_name = OUTPUT_COLUMNS[c - 1]
            if col_name in person.retry_highlight:
                cell.fill = retry_fill
            elif c == len(row_vals) - 1 and val:  # Needs Review column
                cell.fill = review_fill

    widths = [8, 10, 22, 12, 6, 12, 12, 10, 10, 12, 12, 12, 10, 12, 12, 12, 30, 30]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A2"

    wb.save(output_path)
    print(f"\nSaved {len(people)} record(s) to {output_path}")


def build_client(provider: str):
    if provider == "gemini":
        try:
            from google import genai
        except ImportError:
            sys.exit("Missing dependency. Run: pip install google-genai")
        api_key = os.environ.get("GEMINI_API_KEY") or os.environ.get("GOOGLE_API_KEY")
        if not api_key:
            sys.exit(
                "Set the GEMINI_API_KEY environment variable first.\n"
                "Get a free key (no credit card) at https://aistudio.google.com/apikey"
            )
        return genai.Client(api_key=api_key)
    else:
        try:
            import anthropic
        except ImportError:
            sys.exit("Missing dependency. Run: pip install anthropic")
        api_key = os.environ.get("ANTHROPIC_API_KEY")
        if not api_key:
            sys.exit(
                "Set the ANTHROPIC_API_KEY environment variable first.\n"
                "Get a key at https://console.anthropic.com"
            )
        return anthropic.Anthropic(api_key=api_key)


def main():
    ap = argparse.ArgumentParser(description="Extract BCA/PRT scoresheet data into Excel.")
    ap.add_argument("--input", required=True, help="Folder of scanned PDFs/images")
    ap.add_argument("--output", default="bca_prt_results.xlsx", help="Output .xlsx path")
    ap.add_argument(
        "--provider", choices=["gemini", "anthropic"], default="gemini",
        help="Which AI vision API to use. 'gemini' is free (default). 'anthropic' is paid.",
    )
    ap.add_argument(
        "--poppler-path", default=None,
        help=(
            "Folder containing pdftoppm.exe/pdftoppm (Poppler's 'bin' or 'Library\\bin' "
            "folder). Only needed if Poppler isn't on your system PATH - e.g. on locked-down "
            "Windows machines where you extracted a portable Poppler zip instead of installing "
            "it. Example: --poppler-path \"C:\\Users\\you\\Downloads\\poppler-24.08.0\\Library\\bin\""
        ),
    )
    ap.add_argument(
        "--passes", type=int, default=1,
        help=(
            "How many times to read each page (default 1). Values >1 make the model "
            "re-attempt any field it couldn't read on the first pass, and cells that "
            "only got filled in on a later pass are highlighted YELLOW in the output "
            "so you know to double check them. Each extra pass is a full extra API "
            "call per page, so a value of 3 roughly triples time/cost."
        ),
    )
    args = ap.parse_args()

    input_dir = Path(args.input)
    if not input_dir.is_dir():
        sys.exit(f"Input folder not found: {input_dir}")

    client = build_client(args.provider)
    people = process_folder(input_dir, client, args.provider, poppler_path=args.poppler_path, passes=args.passes)
    write_workbook(people, Path(args.output))


if __name__ == "__main__":
    main()
