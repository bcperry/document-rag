# Agent Guide

This repository provides a local semantic index for PDF and PowerPoint corpora. Agents should use the workflows in `.agents/skills/` when the host supports agent skills:

- `.agents/skills/document-rag-install/SKILL.md` for installation and environment validation.
- `.agents/skills/document-rag/SKILL.md` for ingestion, maintenance, search, and source-grounded answers.

## Project Rules

- Use `uv` for all Python environment, dependency, and command execution. Do not use `pip`.
- Run commands from the repository root.
- Never display or persist GitHub credentials. Use the authenticated GitHub CLI.
- Keep separate document corpora in separate SQLite databases under `.rag/`.
- Do not alter PowerPoint chunking: one slide is one atomic chunk.
- Include useful visual descriptions in retrieval, but exclude basic or text-only visual descriptions from embeddings.
- Preserve numeric year/month metadata used to sort reporting periods.
- Do not modify generated databases or virtual-environment files directly.
- Support PowerPoint rendering on both native Windows and WSL; native Windows must not depend on `wslpath`.
- Treat a zero-embedding ingest as non-searchable. Report skipped reasons and use vision-enabled ingestion for image-only slides.
- Use `RAG_EMBED_PROVIDER=local` with a dedicated database for large corpora that exceed GitHub Models embedding quotas.
- Preserve the visual-description cache so interrupted ingestion can resume without repeating completed model calls.

## Common Commands

```bash
uv sync
uv run pytest -q
uv run python document_rag.py --help
uv run python document_rag.py ingest "<document-directory>" --prune
uv run python document_rag.py search "<query>" --limit 8
uv run python document_rag.py stats
```

Pass `--db .rag/<corpus>.db` before the subcommand to select an isolated corpus.

## Answering Questions

Use retrieval results as evidence. Cite PowerPoint claims as `[Document Name, slide N]` and PDF claims as `[Document Name, page N]`. When results conflict, identify the relevant document and reporting period. Say when the indexed evidence is insufficient rather than supplying unsupported details.

## Validation

After code changes, run the narrowest relevant test followed by the full test suite when practical. Keep changes focused and preserve the existing CLI contract unless the task explicitly requires changing it.
