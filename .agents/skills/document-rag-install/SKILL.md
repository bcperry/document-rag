---
name: document-rag-install
description: "Install and validate Document RAG. Use when setting up the PDF and PowerPoint semantic search project, resolving dependencies, configuring GitHub Models authentication, or checking PowerPoint rendering prerequisites. Triggers: install document rag, set up document search, configure PDF search, configure PowerPoint search."
argument-hint: "Optional repository path"
---

# Install Document RAG

Set up this repository for local PDF and PowerPoint ingestion and semantic search.

## Procedure

1. Locate the repository root containing `pyproject.toml` and `document_rag.py`.
2. Verify required commands without displaying credentials.

   Native Windows PowerShell:

   ```powershell
   Get-Command uv
   Get-Command gh
   gh auth status
   ```

   POSIX shell or WSL:

   ```bash
   command -v uv
   command -v gh
   gh auth status
   ```

3. Install the locked Python environment using `uv` only:

   ```bash
   uv sync
   ```

4. Verify Windows PowerPoint automation before ingesting `.pptx` files with vision.

   Native Windows PowerShell:

   ```powershell
   Get-Command uv
   Get-Command gh
   $app = New-Object -ComObject PowerPoint.Application
   $app.Version
   $app.Quit()
   ```

   WSL:

   ```bash
   command -v powershell.exe
   command -v wslpath
   powershell.exe -NoProfile -NonInteractive -Command '$app = New-Object -ComObject PowerPoint.Application; Write-Output $app.Version; $app.Quit()'
   ```

5. Validate the installation:

   ```bash
   uv run pytest -q
   uv run python document_rag.py --help
   ```

6. Report which prerequisites passed and any blocked document formats.

## Rules

- Never use `pip`; this project uses `uv` exclusively.
- Never print or persist the GitHub token. Authentication is read through `gh`.
- Do not create a database during installation unless the user also provides a corpus.
- PowerPoint visual ingestion requires installed Windows PowerPoint, either on native Windows or through WSL interoperability.
- Native Windows must not require `wslpath`; WSL requires both `powershell.exe` and `wslpath`.
- Scanned PDFs require an OCR text layer before their text can be searched.
- Legacy `.ppt` files must be converted to `.pptx`.
