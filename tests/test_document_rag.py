from pathlib import Path

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
    assert rag.is_basic_slide("slide 4", "Agenda\nIntroductions\nNext steps") is True
    assert rag.is_basic_slide("slide 15", "Modernization project status table") is False


def test_text_only_powerpoint_visual_is_forced_basic(tmp_path: Path, monkeypatch):
    monkeypatch.setattr(
        rag,
        "extract_powerpoint",
        lambda path: [("slide 2", "Text-only project update")],
    )
    monkeypatch.setattr(rag, "powerpoint_substantive_visuals", lambda path: [False])
    monkeypatch.setattr(
        rag,
        "add_visual_descriptions",
        lambda path, sections, concurrency: [
            {
                "location": "slide 2",
                "text": "Text-only project update",
                "visual_description": "The slide contains several bullet points.",
                "visual_useful": True,
            }
        ],
    )

    sections = rag.extract_sections(tmp_path / "deck.pptx")

    assert sections[0]["visual_useful"] is False


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
    monkeypatch.setattr(
        rag,
        "extract_sections",
        lambda path, describe_images, vision_concurrency: [
            {
                "location": "slide 1",
                "text": "",
                "visual_description": None,
                "visual_useful": False,
            }
        ],
    )
    monkeypatch.setattr(
        rag,
        "get_embeddings",
        lambda texts: (_ for _ in ()).throw(AssertionError("embedding requested")),
    )

    result = rag.ingest(document_root, db_path, describe_images=False)

    assert result["indexed"] == 0
    assert result["skipped"] == 1
    assert result["documents"][0]["reason"] == (
        "no searchable text or useful visual description"
    )
    assert rag.stats(db_path)["embeddings"] == 0


def test_visual_descriptions_are_added_to_section_text(tmp_path: Path, monkeypatch):
    images = [tmp_path / "slide-1.jpg", tmp_path / "slide-2.jpg"]
    for image in images:
        image.touch()
    monkeypatch.setattr(rag, "render_sections", lambda path, output_dir: images)
    monkeypatch.setattr(
        rag,
        "describe_image",
        lambda path, text: {"description": f"Description {path.stem}", "useful": True},
    )

    sections = rag.add_visual_descriptions(
        tmp_path / "deck.pptx",
        [("slide 1", "First title"), ("slide 2", "Second title")],
        concurrency=1,
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
        "describe_image",
        lambda path, text: {"description": f"Description {path.stem}", "useful": True},
    )

    sections = [(f"slide {index}", f"Title {index}") for index in range(1, 7)]
    described = rag.add_visual_descriptions(
        tmp_path / "deck.pptx", sections, concurrency=3
    )

    assert [item["visual_description"].rsplit(" ", 1)[-1] for item in described] == [
        f"slide-{index}" for index in range(1, 7)
    ]


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
