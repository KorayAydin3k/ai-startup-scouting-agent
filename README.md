# AI Startup Scouting Prototype

## Overview
This project is a lightweight prototype designed to automate the discovery of startup accelerators and their portfolio companies, and to generate concise, AI-assisted value propositions for each startup.

The goal of the project is not large-scale scraping, but to demonstrate a clean, reliable and production-oriented approach to data ingestion and AI-assisted reasoning.


## Architecture
- **Google Sheets** as a simple, transparent datastore  
- **Google Apps Script** as the orchestration layer  
- **LLM API** for controlled natural language generation  

This architecture was intentionally chosen to maximize clarity, reproducibility and ease of review.


## Data Model

### Accelerators Sheet
| Column | Description |
|------|------------|
| website | Normalized accelerator URL (primary key) |
| name | Accelerator name |
| country | Country or region |
| source_url | Source of discovery |
| last_checked_at | ISO 8601 timestamp |

### Startups Sheet
| Column | Description |
|------|------------|
| website | Normalized startup URL (primary key) |
| name | Startup name |
| country | Country |
| accelerator_website | Related accelerator (foreign key) |
| source_url | Portfolio page URL |
| last_checked_at | ISO 8601 timestamp |
| value_proposition | AI-generated description |
| vp_generated_at | Generation timestamp |
| vp_model | Model used |


## Key Design Principles

### Idempotency
All ingestion steps are designed to be safely re-runnable.  
URL normalization and deduplication ensure that running the pipeline multiple times does not create duplicate records.

### Data Quality
URLs are normalized before storage, timestamps are always recorded in ISO format, and each record keeps track of its source.

### Conservative AI Usage
The language model is used only after data collection is complete, and is explicitly instructed to:
- Use only information present on the startup website
- Avoid invented features or marketing language
- Return a single concise sentence

If insufficient information is available, the output remains intentionally generic.


## How It Works

1. **Accelerator ingestion**  
   A small, curated seed list of accelerators is inserted into the `accelerators` sheet.

2. **Startup discovery**  
   For each accelerator, common portfolio paths (e.g. `/portfolio`, `/companies`) are scanned to identify external startup websites.

3. **AI value proposition generation**  
   Visible website text is extracted and passed to the language model to generate a short value proposition.

Each step can be executed independently through a custom menu in Google Sheets.



## Setup Instructions

1. Create a Google Sheet with two tabs: `accelerators` and `startups`
2. Open **Extensions → Apps Script** and paste the content of `Code.gs`
3. (Optional) Set the `OPENAI_API_KEY` in Apps Script **Script Properties**
4. Reload the spreadsheet to enable the custom menu



## Assumptions
- Portfolio pages are publicly accessible and mostly static HTML
- Startup websites provide enough textual content for basic summarization
- The system is intended for small to medium batch sizes


## Limitations
- JavaScript-rendered websites are intentionally skipped
- Startup discovery relies on heuristic link extraction
- No language detection or translation is applied



## Future Improvements
- Headless browser support for JS-heavy websites
- Smarter startup classification and tagging
- Batch processing and rate-limit awareness
- Vector-based semantic search across startups



## Disclaimer
This project is a technical case study and not intended for production use.
