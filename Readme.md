PPT Inconsistency Checker
One-command pipeline that extracts text (and image OCR) from PowerPoint slides and finds factual/logical inconsistencies across slides using rule-based checks and Gemini 2.5 Flash for semantic comparisons.

Features
Phase 1: robust .pptx extraction (text, tables, notes) + image extraction + OCR via pytesseract

Phase 2: rule-based checks for numeric mismatches, percentage-sum checks, timeline/year heuristics

Phase 2: mandatory Gemini 2.5 Flash call for semantic / textual contradiction detection

Merges rule-based + LLM results, deduplicates, prints human-readable summary and JSON report

No intermediate JSON files required (in-memory pipeline)

CLI — runs in terminal, no UI

Installation:
pip install -r requirements.txt

Run pipeline:
python pipeline_full.py \
  --pptx path/to/ppt.pptx \
  --imgdir ./slide_images \
  --output final_report.json \
  --api_key (your gemini key here)


Working:
Extract: python-pptx to pull text, tables and embedded images. Images are saved to --imgdir. OCR  extracts text from images.

Detect (rules): number/token extraction, normalization, context-similarity based numeric matching . Percentage-sum and year heuristics.

Detect (LLM): Gemini 2.5 Flash compares slide summaries for contradictions. Script forces a JSON-only return and  parses from non-JSON outputs.

Merge: rule + LLM issues combined, deduplicated, reported.

Security:
No hardcoded Api key in code, pass via CLI.

Accuracy: rule-based numeric checks catch exact numeric mismatches. Gemini handles  textual contradictions. 

Clarity: output includes slides_involved, issue_type, and excerpts in description. JSON is machine-parsable.

Scalability: script is single-process; for large decks, add batching, clustering, and async model calls.

Thoughtfulness: context-similarity reduces false positives; percent-sum checks catch internal arithmetic issues, fallback parsing for Gemini ensures robust behavior.