# GitHub Copilot Instructions for Iboga Lab Billing Documents

This repository contains tools and templates for drafting and generating official business documents (quotes/devis and invoices/factures) for **Iboga Lab** (https://www.ibogalab.tech/).

When helping the user write or modify these files, follow these guidelines:

1. **Document Formats**:
   - The primary source document is always a Markdown (`.md`) file with YAML front-matter metadata.
   - Output files are generated in HTML, PDF, and DOCX formats.
2. **Branding & Visuals**:
   - Use the brand colors: Forest Green `#103824` (primary) and Sage Green `#669D69` (accent).
   - Fonts used are Montserrat and Poppins.
3. **Automatic Compilation**:
   - Every time a Markdown file is edited, it must be compiled using the Python synchronizer script:
     ```bash
     python skills/ibogalab-invoice-generator/scripts/sync.py path/to/document.md
     ```
   - Suggest running the script in watch mode during edits:
     ```bash
     python skills/ibogalab-invoice-generator/scripts/sync.py --watch path/to/document.md
     ```
4. **Numbering State**:
   - Keep numbers as `number: auto` when creating new files; the script increments the counter (saved in `~/.ibogalab-invoice-settings.json`) and overwrites `auto` in the Markdown.
