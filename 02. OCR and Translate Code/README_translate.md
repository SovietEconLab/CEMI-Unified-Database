# CEMI Translation Script — Usage Guide

`translate_cemi.py` translates CEMI archival Excel files from Russian to
English using the Anthropic Claude API, applying the same translation
conventions used to produce the existing `EN/` reference files.

## Quick start

1. **Install dependencies** (one time):

   ```bash
   pip install openpyxl anthropic
   ```

2. **Open a terminal** and navigate to the folder containing the
   `cemi_1959_1*.xlsx` files (or any folder under it):

   ```bash
   cd "DB 구축 작업 논문/CEMI OCR"
   ```

3. **Run the script**:

   ```bash
   python3 translate_cemi.py
   ```

   The script will:
   - List all `cemi_1959_1*.xlsx` files it finds
   - Show which already have English versions in `EN/`
   - Process each file one at a time, prompting for an API key per file

4. **Enter your Anthropic API key** when prompted. The key is hidden
   while you type. Paste your `sk-ant-...` key and press Enter.

## Common options

| Command | Description |
| --- | --- |
| `python3 translate_cemi.py` | Process all files, prompt for key per file |
| `python3 translate_cemi.py --all` | Prompt for the key only once (then process all) |
| `python3 translate_cemi.py FILE.xlsx` | Translate a single file |
| `python3 translate_cemi.py --dir PATH` | Scan a different directory |
| `python3 translate_cemi.py --model claude-sonnet-4-6` | Use a different model |
| `python3 translate_cemi.py --batch-size 50` | Smaller batches (slower but safer) |

## Skip the per-file prompt

Set the API key once in your shell session:

```bash
export ANTHROPIC_API_KEY=sk-ant-xxx...
python3 translate_cemi.py
```

The script will detect the variable and offer to reuse it for each file.

## Output

Translated files are saved to a new `EN/` subfolder next to the originals,
with the suffix `_EN.xlsx`. For example:

```
CEMI OCR/
├── cemi_1959_1_111_1966_full.xlsx           # original
├── cemi_1959_1_750_1979_full.xlsx           # original
├── EN/
│   ├── cemi_1959_1_111_1966_full_EN.xlsx    # translated
│   └── cemi_1959_1_750_1979_full_EN.xlsx    # translated
└── translate_cemi.py
```

If a translated file already exists, the script asks before overwriting.

## Translation conventions

The script instructs Claude to apply the conventions established for the
existing `EN/` reference files:

- **Names**: BGN/PCGN transliteration (НЕМЧИНОВ → NEMCHINOV)
- **Academic ranks**: Standard English forms (д.э.н. → Doctor of Economic Sciences)
- **Soviet terms**: Romanized historical terms preserved (soiskatel,
  aspirantura, sovmestitel, stazher)
- **Dates**: Roman-numeral months kept (`30.I.1965`)
- **Excel structure**: Sheets, formulas, cell coordinates all preserved

## Cost estimate (rough)

For a typical file with ~1,500 Cyrillic cells: about 20 API calls,
~30,000–50,000 tokens. With Claude Opus 4.6 (Nov 2026 pricing), that's
roughly $1–2 per file. Smaller files (e.g., 200 cells) are about $0.20.

The 750_1979 packet (3,565 cells, the largest) will be the most
expensive — budget about $3.

## Troubleshooting

| Problem | Fix |
| --- | --- |
| `Missing 'anthropic'` | `pip install anthropic` |
| `Missing 'openpyxl'` | `pip install openpyxl` |
| API errors / rate limits | The script auto-retries 3× with backoff; if it still fails, the offline transliteration fallback is used for that batch |
| Korean folder name issues | The script handles macOS NFD/NFC automatically |
| Unicode garbage in output | Verify your terminal supports UTF-8 |

## How it works

1. Scans the target directory for `cemi_1959_1*.xlsx`.
2. For each file:
   - Copies it to `EN/<name>_EN.xlsx`.
   - Walks every cell, finds Cyrillic content, batches into ~80-cell groups.
   - Sends each batch to Claude with the system prompt below.
   - Writes Claude's translations back into the cells.
   - Reports remaining Cyrillic cells (should be 0 if translation succeeded).
3. Saves the workbook with formulas, sheet names, and structure intact.

The system prompt is embedded in the script (search for
`TRANSLATION_SYSTEM_PROMPT`) and codifies all conventions: BGN/PCGN
transliteration table, fixed mappings for ranks/institutions/dates, output
format requirements (strict JSON), and edge-case handling (struck-through
text, handwritten markers, mixed Russian/English content).
