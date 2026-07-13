# Document RAG

A focused local semantic-search index for collections of PDF and PowerPoint files. Documents are recursively discovered, extracted and rendered by page or slide, described by a small vision model, embedded locally or with GitHub Models, and stored in SQLite with `sqlite-vec`. Each PowerPoint slide is one atomic chunk; long PDF pages use overlapping chunks.

## Requirements

- Python 3.10+
- [`uv`](https://docs.astral.sh/uv/)
- Authenticated GitHub CLI: `gh auth login`
- Windows PowerPoint for `.pptx` visual rendering, either from native Windows or through WSL interoperability

`gh auth token` authenticates requests to GitHub Models. By default, text uses `text-embedding-3-small`; slide and page images use the low-tier `openai/gpt-4.1-nano` vision model. Set `RAG_EMBED_PROVIDER=local` to use the CPU-local `BAAI/bge-small-en-v1.5` embedding model instead.

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

By default, PowerPoint slides and PDF pages are rendered to temporary JPEGs. Deterministic checks skip vision requests for cover, agenda, divider, and PowerPoint slides without substantive visual structures. Each remaining vision request carries exactly one image, and up to five requests run concurrently. Completed descriptions are cached under `.rag/vision-cache`, so interrupted or rate-limited ingestion resumes without repeating finished image requests. Descriptions for useful screenshots, photographs, charts, diagrams, and status layouts are included in embeddings. Adjust concurrency with `--vision-concurrency`, or use `--no-vision` for a text-only run.

PowerPoint rendering works when Python runs directly on Windows and when it runs in WSL with `powershell.exe` and `wslpath` available. Native Windows paths are passed directly to PowerPoint; WSL paths are translated before invoking PowerPoint COM.

`--no-vision` can only index text that `python-pptx` extracts. An image-only deck or slide with no useful extracted text has no content to embed and is reported as skipped with `no searchable text or useful visual description`. Run normal vision-enabled ingestion for screenshots, photographs, and other visual-only content.

GitHub Models free usage is rate limited. Requests fail clearly when a daily quota returns a long `Retry-After` value instead of sleeping for hours. For large corpora, use local embeddings so only vision requests consume GitHub Models quota:

```powershell
$env:RAG_EMBED_PROVIDER = "local"
uv run python document_rag.py --db .rag/catalog-local.db ingest ./documents --prune
uv run python document_rag.py --db .rag/catalog-local.db search "degree requirements"
```

Use a separate database for each embedding provider because the vector dimensions differ.

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
