# Slide Content Inconsistency Detection – Design Notes

## Overview
This tool processes a PowerPoint presentation, extracts slide text, and uses Google Gemini to detect inconsistencies in tone, style, and messaging.

---

## Design Decisions

### 1. Input Format
- **Reason:** We directly parse `.pptx` files using `python-pptx` to avoid intermediate JSON files  and reduce I/O.
- **Alternative considered:** Save to JSON first for inspection — rejected for production.

### 2. LLM Choice
- **Gemini Pro** selected for:
  - Strong reasoning capability.
  - Support for structured JSON output.
  - Ability to handle multi-slide context.

### 3. JSON-Only Output Enforcement
- Gemini occasionally returns extra explanations alongside JSON.

### 4. Error Handling
- Graceful handling of:
  - Empty slides.
  - Slides with only images (no text).
  - Gemini API timeouts.

### 5. Output Format
- Final results are stored in `output/results.json` with:
  - `slide_number`
  - `issue`
  - `suggestion`

---

## Prompt Text Used

### Primary Prompt
```text
You are an AI that checks PowerPoint slides for inconsistencies in tone, style, and message. 
Input: JSON list of slides with their text. 
Output: JSON array where each element has:
- slide_number (integer)
- issue (string)
- suggestion (string)

Output ONLY valid JSON. No explanations.
