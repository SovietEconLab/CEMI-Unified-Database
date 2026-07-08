https://doi.org/10.5281/zenodo.21254567
[README.txt](https://github.com/user-attachments/files/29787752/README.txt)
CEMI Career Database — Reproduction Guide

This guide accompanies the Data Descriptor A unified database of the Soviet Central
Economic Mathematical Institute, 1961–1987. It lets anyone reproduce the resource end-to-end: (A) re-run the
OCR + translation of the original ARAN archival scans, and (B) rebuild, from three
frozen Excel inputs, the unified SQLite database cemi_career.db and its
self-contained static interface cemi_career_ui.html.

The complete workflow is the flow chart in Figure 1 of the accompanying Data Descriptor.
One bipartite archival source (ARAN Fond 1959, Opis 1) feeds two collection /
processing pipelines — Track A (2 stages, A1–A2) and Track B (7 stages,
B1–B7) — into three intermediate Excel files, which one deterministic builder
(cemi_career_db.py, 5 phases) absorbs into a single SQLite file; a second script
(cemi_career_ui.py) turns that database into a browser interface.

The two halves are independent. Re-running the OCR + translation (Part A) is
optional — its outputs are frozen in the deposit. Part B (DB + UI) rebuilds from
the three frozen Excel inputs alone and needs no API key.

------------------------------------------------------------------------

0. What is in the deposit

The deposit is organised into numbered stage folders (01–05) plus a set of
root-level files (the codebooks, glossaries, and this README).
The builder scripts, the three Excel inputs, and the prebuilt artefacts
(cemi_career.db, cemi_career_ui.html) live inside folder 05., not at the
repository root — see Part B §B.1 for the working-directory rules this implies.

    CEMI-Unified-Database/
    │
    ├─ README.md · README.txt · README.docx        # this guide (three formats)
    │
    ├─ CEMI_Codebook_v1_EN.pdf                                  # variable codebook
    ├─ CEMI_Institution_Classification_Codebook_v1_EN.pdf           # institution 4-sector codebook
    │
    ├─ institution glossary (Nolting 4-sector classified, all sheets, classified).xlsx  # 4-sector coding
    ├─ institution glossary (Nolting 4-sector classified, all sheets, unified).xlsx
    ├─ glossary(0530).xlsx
    ├─ ARAN F 1959 Op 1 CEMI original document lists.xlsx
    │
    ├─ 01. Raw Materials(Sample)/        # sample ARAN PDF scans (1959 1 *.pdf) + copies of the Part-A scripts
    ├─ 02. OCR and Translate Code/       # Part-A code
    │  ├─ cemi_processor.py              #   Stage B1 — OCR      (PDF → Russian Excel)
    │  ├─ translate_cemi.py              #   Stage B4 — translation (Russian → English Excel)
    │  ├─ requirements.txt              #   Part-A dependencies (see §1)
    │  ├─ README_ocr.md
    │  └─ README_translate.md
    ├─ 03. OCR Results/                  # 25 Stage-B1 Russian OCR workbooks (frozen)
    ├─ 04. English Translation Results/   # 25 Stage-B4 English-translation workbooks (frozen)
    │                                    #   (+ 3 auxiliary glossary workbooks; the builder also accepts the legacy spelling "Tranlation")
    └─ 05. Unified Data and Database Construction Code/   # Part-B code + the three frozen Excel inputs + prebuilt artefacts
       ├─ cemi_career_db.py             #   deterministic DB builder
       ├─ cemi_career_ui.py            #   static-interface generator
       ├─ data.xlsx                     #   Track B input — personnel roster (12,211 × 27)
       ├─ cemi_translation_db.xlsx      #   Track A input — annual cross-tabulations + glossary + sheets-index
       ├─ research_field_subfield.xlsx  #   Track A input — 22 yearly research-subfield snapshots
       ├─ cemi_career.db                #   BUILT artefact — SQLite, LEAN build
       │                                #     (41 tables · 55,711 rows · ~5.4 MB · sheets_index = 181)
       └─ cemi_career_ui.html           #   BUILT artefact — self-contained interface (~7.5 MB)

  Two build tiers. The prebuilt cemi_career.db shipped inside folder 05. is the
  lean build (55,711 rows · 181 sheets_index rows · 636 classified institutions),
  produced by running the builder inside folder 05. with no extra files co-located.
  Reproducing the full-fidelity build documented in the codebook (56,931 rows · 757
  sheets_index rows · 652 classified institutions) requires placing the institution
  glossary and the two provenance folders next to data.xlsx first — see §B.1.

------------------------------------------------------------------------

1. Software requirements

-   Python 3.11 (reference; 3.8+ also works). No external database server is needed —
    SQLite is built into Python.

Part A — OCR + translation (only if you re-run it):

    pip install anthropic openpyxl PyMuPDF Pillow
    export ANTHROPIC_API_KEY="..."      # needed for Part A only

The pinned Part-A dependencies are listed in 02. OCR and Translate Code/requirements.txt
(there is currently no requirements.txt at the repository root). Without the environment
variable, cemi_processor.py and translate_cemi.py prompt for the key interactively.

Part B — building the DB and UI (always required):

    pip install pandas openpyxl
    # sqlite3 is part of the Python standard library

-   cemi_career_db.py requires pandas + openpyxl (pandas reads the .xlsx inputs
    through openpyxl).
-   cemi_career_ui.py requires no third-party packages at all — it uses only the
    Python standard library (sqlite3, json).

A virtual environment of ≈ 80 MB is sufficient for the full build.

------------------------------------------------------------------------

2. The pipeline at a glance (Figure 1)

Bipartite source → ARAN Fond 1959, Opis 1 (annual delo, 1961–1987).

Track A · demographic / aggregate — transcribed by hand at ARAN; no OCR
(the regular tabular otchet template makes vision-API OCR unnecessary):

  ---------------------------------------------------------------------------------------------------------------
  Stage                   What it does                   Output
  ----------------------- ------------------------------ --------------------------------------------------------
  A1                      On-site manual transcription   two thematic workbooks

  A2                      Thematic separation            cemi_translation_db.xlsx, research_field_subfield.xlsx
  ---------------------------------------------------------------------------------------------------------------

Track B · personnel career / OCR pipeline — seven stages:

  --------------------------------------------------------------------------------------------------------------------------
  Stage                   What it does                                    Tool
  ----------------------- ----------------------------------------------- --------------------------------------------------
  B1                      Per-page OCR of the spiski scans                cemi_processor.py (Claude vision)

  B2                      Per-year Russian Excel emission                 —

  B3                      Manual correction · RU                          (spreadsheet)

  B4                      Translation RU → EN                             translate_cemi.py (Claude API)

  B5                      Manual correction · EN                          (spreadsheet)

  B6                      Consolidation into one workbook                 → data.xlsx (12,211 × 27)

  B7                      Final correction pass on the unified workbook   (pre-revision baseline kept as a row-level diff)
  --------------------------------------------------------------------------------------------------------------------------

Convergence: cemi_career_db.py (5 phases) → cemi_career.db →
cemi_career_ui.py → cemi_career_ui.html.

------------------------------------------------------------------------

Part A — Reproducing the OCR + English Translation (optional)

  The 25 Stage-B1 Russian workbooks and 25 Stage-B4 English workbooks are frozen in the
  deposit (03. OCR Results/, 04. English Translation Results/). Running this part
  overwrites them. Skip to Part B if you only want to rebuild cemi_career.db.
  Track A carries no OCR — the two Track-A workbooks are produced by hand
  transcription (stages A1–A2).

A.1 What the "1959 1 *.pdf" files look like

    1959 1 <delo>.pdf            e.g.  1959 1 111.pdf
    1959 1 <delo>(<year>).pdf    e.g.  1959 1 76(1965).pdf

1959 is the fond number and 1 the opis number in ARAN’s catalogue; the third
token is the delo; the optional (year) suffix is the reporting year. Stage B1 reads
the delo and year from the filename.

A.2 Stage B1 — OCR with cemi_processor.py

    cd "02. OCR and Translate Code"
    python cemi_processor.py "../01. Raw Materials(Sample)"

The script lists every PDF containing "1959", you select one per invocation, it reads
ANTHROPIC_API_KEY, rasterises each page (default 150 DPI), sends it to the Claude
vision API for per-page Russian OCR, then structures the document into a multi-sheet
workbook written next to the input PDF as cemi_1959_1_<delo>_<year>_full.xlsx.
Per-page OCR is cached under .cemi_cache/<pdf_stem>/ so a failed run resumes without
re-billing.

Useful flags: --model (Anthropic model id), --dpi (default 150), --force (ignore
cache). After Stage B1 you have 25 packet workbooks in 03. OCR Results/. Stage B3
(manual Russian-side correction) is then done by hand in Excel.

A.3 Stage B4 — Russian → English translation with translate_cemi.py

    cd "03. OCR Results"
    python "../02. OCR and Translate Code/translate_cemi.py" --all

  ------------------------------------------------------------------------------------
  Command                                    Behaviour
  ------------------------------------------ -----------------------------------------
  python translate_cemi.py                   Interactive — prompts for the API key

  python translate_cemi.py --all             Translate every file in --dir

  python translate_cemi.py FILE.xlsx         Translate a single file

  python translate_cemi.py --dir <path>      Translate files in a specific directory

  python translate_cemi.py --model <id>      Override the default model

  python translate_cemi.py --batch-size 80   Override the cell batch size
  ------------------------------------------------------------------------------------

It applies BGN/PCGN transliteration, fixes Soviet academic ranks to canonical English,
preserves Roman-numeral months, and keeps Russian historical terms alongside their
English equivalents. Output is written with an _EN.xlsx suffix; move the files into
04. English Translation Results/. Stage B5 (manual English-side correction) follows
by hand.

A.4 Stages B6 + B7 — Consolidation into data.xlsx

The 25 per-year *_EN.xlsx workbooks are consolidated into a single
data.xlsx (12,211 rows × 27 columns, 1961–1987) — stage B6 — followed by one
stage B7 final correction pass on the unified workbook (within-year promotion order,
dual-appointment splits, 1961/1962 institute-phase tagging). The pre-B7 baseline is kept
verbatim (the 25 04. English Translation Results/ workbooks) so every Stage-B7 edit is
recoverable as a row-level diff.

After Stage B7 you have the three canonical Excel inputs that Part B consumes, all
inside folder 05.: data.xlsx, cemi_translation_db.xlsx, research_field_subfield.xlsx.

------------------------------------------------------------------------

Part B — Building cemi_career.db and cemi_career_ui.html

  Deterministic and idempotent: the same inputs always produce a bit-identical SQLite
  file (SHA-256 verified). No Anthropic API key is needed.

B.1 ⚠️ Execution condition — the working directory and co-located inputs

The Part-B scripts and their three Excel inputs live in
05. Unified Data and Database Construction Code/, not at the repository root.
Running python cemi_career_db.py from the repo root fails (file not found,
exit 2). You must cd into folder 05. first (or pass explicit -- paths, §B.2).

cemi_career_db.py resolves its three Excel inputs by filename in the current working
directory (or by explicit -- paths). In addition, it derives a base directory from
the location of data.xlsx and, relative to that base directory, it looks for:

-   the institution glossary (institution glossary*.xlsx, auto-discovered beside data.xlsx);
-   an OCR Results/ folder (Track-B provenance);
-   an English Translation Results/ folder (the legacy spelling English Tranlation Results/
    is also accepted).

Whatever it finds beside data.xlsx determines which of the two build tiers you get.

Tier 1 — Lean build (matches the prebuilt cemi_career.db in folder 05.)

    cd "05. Unified Data and Database Construction Code"
    python cemi_career_db.py

With only the three Excel inputs present, glossary enrichment and OCR/translation
provenance are skipped. Result: 55,711 rows · 181 sheets_index rows · 636 classified
institutions — identical to the prebuilt cemi_career.db shipped in folder 05.

Tier 2 — Full-fidelity build (the counts documented in the codebook: 56,931 rows)

To reproduce the full deposit, co-locate the glossary and the two provenance folders
next to data.xlsx (i.e. inside folder 05.) before building. From the repository
root:

macOS / Linux

    cd "05. Unified Data and Database Construction Code"
    cp "../institution glossary (Nolting 4-sector classified, all sheets, classified).xlsx" .
    cp -R "../03. OCR Results" "OCR Results"
    cp -R "../04. English Translation Results" "English Translation Results"
    python cemi_career_db.py

Windows (PowerShell) — symlinks/junctions avoid copying:

    cd "05. Unified Data and Database Construction Code"
    Copy-Item "..\institution glossary (Nolting 4-sector classified, all sheets, classified).xlsx" .
    New-Item -ItemType Junction -Name "OCR Results" -Target "..\03. OCR Results"
    New-Item -ItemType Junction -Name "English Translation Results" -Target "..\04. English Translation Results"
    python cemi_career_db.py

Result: 56,931 rows · 757 sheets_index rows · 652 classified institutions — the
figures reported in the Data Descriptor and in the Verification block below.

Summary of the two tiers:

  ------------------------------------------------------------------------------------------------------------------------------------------------------
  Build                                                      Total rows   sheets_index   classified institutions   person      person_year_observation
  ---------------------------------------------------------- ------------ -------------- ------------------------- ----------- -------------------------
  Lean (prebuilt DB / folder-05-only)                        55,711       181            636                       1,954       12,211

  Full-fidelity (glossary + provenance folders co-located)   56,931       757            652                       1,954       12,211
  ------------------------------------------------------------------------------------------------------------------------------------------------------

The core career counts (1,954 persons, 12,211 person-year observations, 41 tables)
are identical in both tiers; only the provenance and institution-classification
enrichment differ.

B.2 Build the database — one command (from folder 05.)

    python cemi_career_db.py

That single command runs the five sequential phases (sharing one in-memory Vocab
cache so dimension IDs agree across the event and aggregate sides):

1.  Phase A — schema bootstrap (41 tables + 20 indexes).
2.  Phase B — personnel loading from data.xlsx (identity hash → person_year_observation
    → up to 9 source event-fact tables); institution glossary enrichment (if the glossary
    is co-located per §B.1).
3.  Phase B.5 — derived artefacts (cemi_position_span, person_degree,
    start_of_work.is_initial_join).
4.  Phase C — aggregate loading from cemi_translation_db.xlsx (9 demo_* tables +
    glossary + sheets_index provenance).
5.  Phase D — research subfields from research_field_subfield.xlsx
    (demo_subfield with schema_year_signature).

Re-running on the same inputs overwrites the file and reproduces it bit-for-bit.

Optional flags

  Flag            Default                        Description
  --------------- ------------------------------ -----------------------------------------------
  --personnel     data.xlsx                      personnel workbook (Track B)
  --translation   cemi_translation_db.xlsx       aggregate workbook (Track A)
  --research      research_field_subfield.xlsx   subfield workbook (Phase D)
  --db            cemi_career.db                 output SQLite path
  --reset         (no-op)                        the builder always overwrites the existing DB

Example with explicit paths (the glossary and the OCR Results/ +
English Translation Results/ folders must sit next to --personnel for the
full-fidelity build):

    python cemi_career_db.py \
       --personnel   ./inputs/data.xlsx \
       --translation ./inputs/cemi_translation_db.xlsx \
       --research    ./inputs/research_field_subfield.xlsx \
       --db          ./outputs/cemi_career.db

B.3 Generate the interface — one command

    python cemi_career_ui.py

Reads cemi_career.db and writes a single self-contained cemi_career_ui.html
(~7.5 MB) that opens in any browser without a server. It exposes the corpus on twelve
tabs — Overview, Career search, Institutions, Positions, Personnel growth,
Degrees & titles, Nationality, Age distribution, Research fields, Party & trainees,
Provenance, and the bilingual Glossary. A Provenance tab exposes the full sheets_index of transcribed source sheets, and a
per-record “Source Material” field cites the originating ARAN document
(fond/opis/delo). This script uses only the
Python standard library — no pandas/openpyxl needed.

Optional flags

  Flag    Default               Description
  ------- --------------------- -----------------------
  --db    cemi_career.db        input SQLite database
  --out   cemi_career_ui.html   output HTML file

Build time (DB + UI) is under five minutes on a commodity laptop (no GPU).

------------------------------------------------------------------------

Verification

The invariant counts below hold for both build tiers; the tier-dependent counts
(total rows, sheets_index) are annotated with their lean / full-fidelity values.

    python - <<'PY'
    import sqlite3
    con = sqlite3.connect("cemi_career.db")
    tabs = [r[0] for r in con.execute(
        "SELECT name FROM sqlite_master WHERE type='table' AND name NOT LIKE 'sqlite_%'")]
    print("tables:", len(tabs))                                   # expect 41
    print("total rows:", sum(con.execute(f'SELECT COUNT(*) FROM \"{t}\"').fetchone()[0] for t in tabs))  # lean 55,711 · full 56,931
    print("persons:", con.execute("SELECT COUNT(*) FROM person").fetchone()[0])                 # 1,954
    print("person-year obs:", con.execute("SELECT COUNT(*) FROM person_year_observation").fetchone()[0]) # 12,211
    print("sheets_index:", con.execute("SELECT COUNT(*) FROM sheets_index").fetchone()[0])      # lean 181 · full 757
    print("integrity:", con.execute("PRAGMA integrity_check").fetchone()[0])                    # ok
    print("fk violations:", con.execute("PRAGMA foreign_key_check").fetchall())                 # []
    PY

If the sqlite3 command-line tool is installed you can also run
sqlite3 cemi_career.db "PRAGMA integrity_check; PRAGMA foreign_key_check;" — both
should return ok / no rows.

SHA-256. The lean and full-fidelity builds are different files with different
hashes; each is bit-identical across repeated runs on the same inputs. When comparing
shasum -a 256 cemi_career.db against a published Zenodo checksum, confirm which build
tier that release corresponds to (the prebuilt DB shipped in folder 05. is the lean build).

------------------------------------------------------------------------

License and citation

-   Data: CC-BY 4.0 (the SQLite file and all Excel inputs).
-   Code: MIT License (everything in this repository).

All code is hosted at https://github.com/SovietEconLab/CEMI-Unified-Database, with
each tagged release mirrored to Zenodo under the DOI
https://doi.org/10.5281/zenodo.21254567.
When citing the dataset, cite both the Data Descriptor and the version-specific
Zenodo DOI of the release you ran.

Dataset citation:

  Kim, D., Hwang, C. & Kim, S. A unified database of the Soviet Central Economic
  Mathematical Institute, 1961-1987. Zenodo
  https://doi.org/10.5281/zenodo.21254567 (2026).
