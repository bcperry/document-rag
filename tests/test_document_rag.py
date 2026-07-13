from pathlib import Path
import urllib.error

import pytest
import document_rag as rag


def test_chunk_text_preserves_size_limit():
    text = " ".join(f"word-{index}" for index in range(200))

    chunks = rag.chunk_text(text, max_chars=200, overlap_chars=30)

    assert len(chunks) > 1
    assert all(len(chunk) <= 200 for chunk in chunks)


def test_discover_documents_is_recursive_and_filters_extensions(tmp_path: Path):
    nested = tmp_path / "nested"
    nested.mkdir()
    (tmp_path / "guide.PDF").touch()
    (nested / "deck.pptx").touch()
    (nested / "notes.txt").touch()

    documents = rag.discover_documents(tmp_path)

    assert documents == [tmp_path / "guide.PDF", nested / "deck.pptx"]


def test_discover_documents_can_limit_search_to_top_level(tmp_path: Path):
    nested = tmp_path / "nested"
    nested.mkdir()
    top_level = tmp_path / "guide.pdf"
    top_level.touch()
    (nested / "deck.pptx").touch()

    documents = rag.discover_documents(tmp_path, recursive=False)

    assert documents == [top_level]


def test_windows_path_does_not_call_wslpath_on_native_windows(tmp_path: Path, monkeypatch):
    source = tmp_path / "Presentation.pptx"
    expected = str(source.resolve())
    monkeypatch.setattr(rag.sys, "platform", "win32")
    monkeypatch.setattr(
        rag.subprocess,
        "run",
        lambda *args, **kwargs: (_ for _ in ()).throw(AssertionError("wslpath called")),
    )

    assert rag.windows_path(source) == expected


def test_rendered_slide_images_deduplicates_windows_case_globs(tmp_path: Path, monkeypatch):
    monkeypatch.setattr(rag.sys, "platform", "win32")

    def fake_glob(self, pattern):
        if pattern == "*.JPG":
            return [self / "Slide1.JPG"]
        if pattern == "*.jpg":
            return [self / "slide1.jpg"]
        return []

    monkeypatch.setattr(rag.Path, "glob", fake_glob)

    assert rag.rendered_slide_images(tmp_path) == [tmp_path / "slide1.jpg"]


def test_parse_document_date_normalizes_filename_variants():
    assert rag.parse_document_date("ARCYBER Microsoft Unified CSDR APR 2026") == (2026, 4)
    assert rag.parse_document_date("Army CHRA Microsoft Unified CSDR July 2026") == (2026, 7)
    assert rag.parse_document_date("Army USACE TAD Microsoft Unified CSDR JAN2026") == (2026, 1)
    assert rag.parse_document_date("Review 2025-09") == (2025, 9)
    assert rag.parse_document_date("Roadmap 2027") == (2027, None)
    assert rag.parse_document_date("Undated presentation") == (None, None)


def test_basic_slide_detection_forces_cover_and_agenda_only():
    assert rag.is_basic_slide("slide 1", "Quarterly Review") is True
    assert rag.is_basic_slide("page 1", "Quarterly Review") is True
    assert rag.is_basic_slide("slide 4", "Agenda\nIntroductions\nNext steps") is True
    assert rag.is_basic_slide("slide 15", "Modernization project status table") is False


def test_text_only_powerpoint_visual_is_forced_basic(tmp_path: Path, monkeypatch):
    monkeypatch.setattr(
        rag,
        "extract_powerpoint",
        lambda path: [("slide 2", "Text-only project update")],
    )
    monkeypatch.setattr(rag, "powerpoint_substantive_visuals", lambda path: [False])
    observed_mask = []

    def add_descriptions(path, sections, concurrency, describe_mask):
        observed_mask.extend(describe_mask)
        return [
            {
                "location": "slide 2",
                "text": "Text-only project update",
                "visual_description": None,
                "visual_useful": False,
            }
        ]

    monkeypatch.setattr(rag, "add_visual_descriptions", add_descriptions)

    sections = rag.extract_sections(tmp_path / "deck.pptx")

    assert sections[0]["visual_useful"] is False
    assert observed_mask == [False]


def test_ingest_indexes_and_skips_unchanged_document(tmp_path: Path, monkeypatch):
    document_root = tmp_path / "documents"
    document_root.mkdir()
    source = document_root / "sample APR 2026.pdf"
    source.write_bytes(b"fake pdf")
    db_path = tmp_path / "index.db"

    monkeypatch.setattr(
        rag,
        "extract_sections",
        lambda path, describe_images, vision_concurrency: [
            {
                "location": "page 1",
                "text": "searchable text",
                "visual_description": "A text-only page.",
                "visual_useful": False,
            }
        ],
    )
    monkeypatch.setattr(
        rag,
        "get_embeddings",
        lambda texts: [[float(index % 7) for index in range(rag.EMBED_DIM)] for _ in texts],
    )

    first = rag.ingest(document_root, db_path)
    second = rag.ingest(document_root, db_path)

    assert first["indexed"] == 1
    assert first["documents"][0]["name"] == "sample APR 2026"
    assert first["documents"][0]["year"] == 2026
    assert first["documents"][0]["month"] == 4
    assert first["errors"] == []
    assert second["unchanged"] == 1
    assert rag.stats(db_path)["chunks"] == 1

    results = rag.search("find this", db_path=db_path)

    assert results[0]["name"] == "sample APR 2026"
    assert results[0]["year"] == 2026
    assert results[0]["month"] == 4
    assert results[0]["path"] == "sample APR 2026.pdf"
    assert results[0]["location"] == "page 1"
    assert results[0]["text"] == "searchable text"


def test_ingest_skips_powerpoint_without_embeddable_content(tmp_path: Path, monkeypatch):
    document_root = tmp_path / "documents"
    document_root.mkdir()
    (document_root / "Presentation.pptx").write_bytes(b"fake pptx")
    db_path = tmp_path / "index.db"
    extract_calls = 0

    def extract_empty_sections(path, describe_images, vision_concurrency):
        nonlocal extract_calls
        extract_calls += 1
        return [
            {
                "location": "slide 1",
                "text": "",
                "visual_description": None,
                "visual_useful": False,
            }
        ]

    monkeypatch.setattr(
        rag,
        "extract_sections",
        extract_empty_sections,
    )
    monkeypatch.setattr(
        rag,
        "get_embeddings",
        lambda texts: (_ for _ in ()).throw(AssertionError("embedding requested")),
    )

    result = rag.ingest(document_root, db_path, describe_images=False)
    second = rag.ingest(document_root, db_path, describe_images=False)

    assert result["indexed"] == 0
    assert result["skipped"] == 1
    assert result["documents"][0]["reason"] == (
        "no searchable text or useful visual description"
    )
    assert second["unchanged"] == 1
    assert extract_calls == 1
    assert rag.stats(db_path)["documents"] == 1
    assert rag.stats(db_path)["embeddings"] == 0


def test_changed_document_with_no_embeddable_content_removes_stale_chunks(
    tmp_path: Path, monkeypatch
):
    document_root = tmp_path / "documents"
    document_root.mkdir()
    source = document_root / "sample.pdf"
    source.write_bytes(b"searchable")
    db_path = tmp_path / "index.db"
    sections = [
        {
            "location": "page 1",
            "text": "searchable text",
            "visual_description": None,
            "visual_useful": False,
        }
    ]
    monkeypatch.setattr(
        rag,
        "extract_sections",
        lambda path, describe_images, vision_concurrency: sections,
    )
    monkeypatch.setattr(
        rag,
        "get_embeddings",
        lambda texts: [[float(index % 7) for index in range(rag.EMBED_DIM)] for _ in texts],
    )

    first = rag.ingest(document_root, db_path, describe_images=False)
    source.write_bytes(b"empty")
    sections.clear()
    second = rag.ingest(document_root, db_path, describe_images=False)

    assert first["indexed"] == 1
    assert second["skipped"] == 1
    assert rag.stats(db_path)["documents"] == 1
    assert rag.stats(db_path)["chunks"] == 0
    assert rag.stats(db_path)["embeddings"] == 0


def test_visual_descriptions_are_added_to_section_text(tmp_path: Path, monkeypatch):
    images = [tmp_path / "slide-1.jpg", tmp_path / "slide-2.jpg"]
    for image in images:
        image.touch()
    monkeypatch.setattr(rag, "render_sections", lambda path, output_dir: images)
    monkeypatch.setattr(
        rag,
        "describe_image_cached",
        lambda path, text: {"description": f"Description {path.stem}", "useful": True},
    )

    sections = rag.add_visual_descriptions(
        tmp_path / "deck.pptx",
        [("slide 1", "First title"), ("slide 2", "Second title")],
        concurrency=1,
        describe_mask=[True, True],
    )

    assert sections == [
        {
            "location": "slide 1",
            "text": "First title",
            "visual_description": "Description slide-1",
            "visual_useful": False,
        },
        {
            "location": "slide 2",
            "text": "Second title",
            "visual_description": "Description slide-2",
            "visual_useful": True,
        },
    ]
def test_visual_descriptions_preserve_order_with_concurrency(tmp_path: Path, monkeypatch):
    images = [tmp_path / f"slide-{index}.jpg" for index in range(1, 7)]
    monkeypatch.setattr(rag, "render_sections", lambda path, output_dir: images)
    monkeypatch.setattr(
        rag,
        "describe_image_cached",
        lambda path, text: {"description": f"Description {path.stem}", "useful": True},
    )

    sections = [(f"slide {index}", f"Title {index}") for index in range(1, 7)]
    described = rag.add_visual_descriptions(
        tmp_path / "deck.pptx",
        sections,
        concurrency=3,
        describe_mask=[True] * len(sections),
    )

    assert [item["visual_description"].rsplit(" ", 1)[-1] for item in described] == [
        f"slide-{index}" for index in range(1, 7)
    ]


def test_visual_description_cache_reuses_completed_request(tmp_path: Path, monkeypatch):
    image = tmp_path / "slide.jpg"
    image.write_bytes(b"image bytes")
    cache_dir = tmp_path / "cache"
    calls = 0

    def describe(path, text):
        nonlocal calls
        calls += 1
        return {"description": "Cached description", "useful": True}

    monkeypatch.setattr(rag, "VISION_CACHE_DIR", cache_dir)
    monkeypatch.setattr(rag, "describe_image", describe)

    first = rag.describe_image_cached(image, "Status chart")
    second = rag.describe_image_cached(image, "Status chart")

    assert first == second
    assert calls == 1
    assert len(list(cache_dir.glob("*.json"))) == 1


def test_local_embeddings_use_configured_embedder(monkeypatch):
    class Vector:
        def tolist(self):
            return [0.1, 0.2, 0.3]

    class Embedder:
        def embed(self, texts, batch_size):
            assert texts == ["one", "two"]
            assert batch_size == rag.EMBED_BATCH_SIZE
            return [Vector(), Vector()]

    monkeypatch.setattr(rag, "EMBED_PROVIDER", "local")
    monkeypatch.setattr(rag, "_local_embedder", Embedder())

    assert rag.get_embeddings(["one", "two"]) == [
        [0.1, 0.2, 0.3],
        [0.1, 0.2, 0.3],
    ]


def test_github_embedding_daily_quota_fails_without_long_sleep(monkeypatch):
    error = urllib.error.HTTPError(
        rag.EMBED_URL,
        429,
        "rate limited",
        {"Retry-After": "62925"},
        None,
    )
    monkeypatch.setattr(rag, "EMBED_PROVIDER", "github")
    monkeypatch.setattr(rag, "get_github_token", lambda: "token")
    monkeypatch.setattr(
        rag.urllib.request,
        "urlopen",
        lambda *args, **kwargs: (_ for _ in ()).throw(error),
    )

    with pytest.raises(RuntimeError, match="quota exhausted.*62925 seconds"):
        rag.get_embeddings(["query"])


def test_basic_visual_description_is_stored_but_not_embedded(tmp_path: Path, monkeypatch):
    monkeypatch.setattr(
        rag,
        "extract_sections",
        lambda path, describe_images, vision_concurrency: [
            {
                "location": "slide 1",
                "text": "Welcome to the review",
                "visual_description": "A basic title slide with a logo.",
                "visual_useful": False,
            }
        ],
    )

    chunks = rag.make_chunks(
        tmp_path / "deck.pptx", tmp_path, 2400, 300, True, 1
    )

    assert chunks[0]["visual_description"] == "A basic title slide with a logo."
    assert chunks[0]["visual_useful"] is False
    assert chunks[0]["embedding_text"] == "Welcome to the review"


def test_powerpoint_slide_is_one_chunk_regardless_of_length(tmp_path: Path, monkeypatch):
    long_text = "A" * 5000
    monkeypatch.setattr(
        rag,
        "extract_sections",
        lambda path, describe_images, vision_concurrency: [
            {
                "location": "slide 9",
                "text": long_text,
                "visual_description": "A useful status diagram.",
                "visual_useful": True,
            }
        ],
    )

    chunks = rag.make_chunks(
        tmp_path / "deck.pptx", tmp_path, 2400, 300, True, 1
    )

    assert len(chunks) == 1
    assert chunks[0]["location"] == "slide 9"
    assert chunks[0]["text"] == long_text
    assert chunks[0]["visual_description"] == "A useful status diagram."
