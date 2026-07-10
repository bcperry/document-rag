---
name: document-rag
description: "Ingest, maintain, and query PDF and PowerPoint corpora with Document RAG. Use for semantic document search, searching slides or screenshots, indexing a document directory, answering questions from presentations, querying a RAG database, or citing PDF pages and PowerPoint slides. Triggers: index documents, ingest PDFs, ingest PowerPoints, search documents, query corpus, ask the slides, find in presentations."
argument-hint: "Corpus directory, database path, or natural-language query"
---

# Use Document RAG

Operate a local semantic index over PDF and PowerPoint files. Run all commands from the repository root using `uv`.

## Select A Corpus

- Use `.rag/index.db` when the user does not specify a database.
- Use a dedicated `--db .rag/<corpus>.db` when corpora must remain isolated.
- Keep using the same database for later queries against that corpus.
- For a Windows path in WSL, translate `C:\...` to `/mnt/c/...` and preserve spaces with shell quoting.

## Ingest

Index a directory recursively:

```bash
uv run python document_rag.py --db .rag/<corpus>.db ingest "<document-directory>" --prune
```

Ingestion behavior:

- Supports `.pdf` and `.pptx`.
- Stores document name plus numeric year/month parsed from filenames when present.
- Keeps each PowerPoint slide as exactly one chunk.
- Splits only long PDF pages.
- Renders pages and slides, then describes each image with one VLM request.
- Runs up to five vision requests concurrently by default.
- Stores descriptions for basic/text-only slides but excludes them from embeddings.
- Skips unchanged files by content hash.

Use `--no-vision` only when the user explicitly accepts text-only retrieval. Image-only slides without extractable text cannot produce embeddings in this mode and are reported as skipped. Adjust request pressure with `--vision-concurrency <n>`.

After ingestion, run:

```bash
uv run python document_rag.py --db .rag/<corpus>.db stats
```

Report document count, chunk count, embedding count, skipped files and their reasons, and errors. Do not describe an index with zero embeddings as a successful end-to-end retrieval test.

## Query

Run semantic search with the user's natural-language question:

```bash
uv run python document_rag.py --db .rag/<corpus>.db search "<query>" --limit 8
```

For broad or ambiguous questions, run two or three focused query variants and merge the evidence. Prefer results with lower vector distance, but evaluate the text rather than treating distance as confidence.

## Answer From Results

1. Read the returned extracted text and any useful visual description.
2. Synthesize only claims supported by retrieved results.
3. Cite every material claim using the result metadata:
   - PowerPoint: `[Document Name, slide N]`
   - PDF: `[Document Name, page N]`
4. Distinguish conflicting documents or reporting periods using numeric year/month metadata.
5. State when retrieval is insufficient instead of filling gaps from assumptions.
6. Do not expose raw database internals unless the user asks for diagnostics.

## Maintenance

Refresh a corpus after files change:

```bash
uv run python document_rag.py --db .rag/<corpus>.db ingest "<document-directory>" --prune
```

The command reprocesses changed files, skips unchanged files, and removes records for deleted files.
