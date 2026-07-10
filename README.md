# Document RAG

A focused local semantic-search index for collections of PDF and PowerPoint files. Documents are recursively discovered, extracted and rendered by page or slide, described by a small vision model, embedded with GitHub Models, and stored in SQLite with `sqlite-vec`. Each PowerPoint slide is one atomic chunk; long PDF pages use overlapping chunks.

## Requirements

- Python 3.10+
- [`uv`](https://docs.astral.sh/uv/)
- Authenticated GitHub CLI: `gh auth login`
- Windows PowerPoint for `.pptx` visual rendering, either from native Windows or through WSL interoperability

`gh auth token` authenticates requests to GitHub Models. Text uses `text-embedding-3-small`; slide and page images use the low-tier `openai/gpt-4.1-nano` vision model.

## Setup

```bash
uv sync
```

## Index Documents

Point `ingest` at a directory. It recursively finds `.pdf` and `.pptx` files:

```bash
uv run python document_rag.py ingest ./documents
```

Re-running the command skips files whose SHA-256 hash has not changed. Use `--prune` to also remove indexed records for files no longer present:

```bash
uv run python document_rag.py ingest ./documents --prune
```

By default, every PowerPoint slide and PDF page is rendered to a temporary JPEG. Each vision request carries exactly one image, and up to five requests run concurrently. Basic title, intro, divider, agenda, and text-only slides are flagged as not visually useful using both model classification and deterministic PowerPoint structure checks. Their descriptions are stored for inspection but excluded from embeddings; extracted slide text remains searchable. Descriptions for visually useful screenshots, photographs, charts, diagrams, and status layouts are included in embeddings. Adjust concurrency with `--vision-concurrency`, or use `--no-vision` for a text-only run.

PowerPoint rendering works when Python runs directly on Windows and when it runs in WSL with `powershell.exe` and `wslpath` available. Native Windows paths are passed directly to PowerPoint; WSL paths are translated before invoking PowerPoint COM.

`--no-vision` can only index text that `python-pptx` extracts. An image-only deck or slide with no useful extracted text has no content to embed and is reported as skipped with `no searchable text or useful visual description`. Run normal vision-enabled ingestion for screenshots, photographs, and other visual-only content.

GitHub Models free usage is rate limited. One-image requests are more reliable with the small vision model, while concurrency reduces time spent waiting on individual responses.

Scanned PDFs need OCR before ingestion; `pypdf` only extracts an existing text layer. Legacy `.ppt` files must be converted to `.pptx`.

## Search

```bash
uv run python document_rag.py search "What are the program admission requirements?"
uv run python document_rag.py search "faculty research priorities" --limit 12
```

Each result includes the document name, parsed year/month when present in the filename, source-relative path, file type, page or slide, chunk text, and vector distance. Lower distance means a closer semantic match.

## Index Options

The default database is `.rag/index.db`. Override it per corpus with `--db` or globally with `RAG_DB_PATH`:

```bash
uv run python document_rag.py --db .rag/catalog-a.db ingest ./catalog-a
uv run python document_rag.py --db .rag/catalog-a.db search "degree requirements"
```

PDF page chunk sizing is configurable during ingestion. These options do not split PowerPoint slides:

```bash
uv run python document_rag.py ingest ./documents --max-chars 2400 --overlap-chars 300
```

Use one database per document root. Relative source paths are the stable document identifiers within that corpus.

## Maintenance

```bash
uv run python document_rag.py stats
uv run pytest
```

## Agent Skills

Portable workspace skills and root-level agent instructions are provided for AI agents:

- `AGENTS.md` provides repository-wide operating rules and common commands.
- `.agents/skills/document-rag-install/SKILL.md` installs and validates prerequisites.
- `.agents/skills/document-rag/SKILL.md` ingests, maintains, queries, and cites document corpora.

Agents that support the open agent-skills directory convention can discover these workflows from `.agents/skills/`. Other agents can read `AGENTS.md` directly. The operational skill instructs agents to keep corpora in separate databases, use source-relative page or slide citations, and synthesize answers only from retrieved evidence.
