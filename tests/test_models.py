import pytest

from models import normalize_presentation


def _deck(*slides):
    return {"title": "Test", "slides": list(slides)}


def test_requires_dict():
    with pytest.raises(ValueError):
        normalize_presentation(["not", "a", "dict"])


def test_requires_at_least_one_slide():
    with pytest.raises(ValueError):
        normalize_presentation(_deck())


def test_unknown_layout_downgrades_to_text():
    data = normalize_presentation(_deck({"type": "content", "layout": "made_up", "title": "X"}))
    assert data["slides"][0]["layout"] == "text"


def test_section_without_layout_becomes_divider():
    data = normalize_presentation(_deck({"type": "section", "title": "Раздел"}))
    assert data["slides"][0]["layout"] == "divider"


def test_section_with_quote_layout_is_preserved():
    data = normalize_presentation(
        _deck({"type": "section", "layout": "quote", "title": "Цитата", "subtitle": "Автор"})
    )
    assert data["slides"][0]["layout"] == "quote"
    assert data["slides"][0]["subtitle"] == "Автор"


def test_chart_with_insufficient_data_downgrades_to_text():
    data = normalize_presentation(
        _deck({
            "type": "content", "layout": "chart", "title": "X",
            "chart": {"type": "bar", "categories": ["Только одна"], "series": [{"name": "Y", "values": [1]}]},
        })
    )
    assert data["slides"][0]["layout"] == "text"
    assert "chart" not in data["slides"][0]


def test_valid_chart_is_kept():
    data = normalize_presentation(
        _deck({
            "type": "content", "layout": "chart", "title": "X",
            "chart": {
                "type": "bar",
                "categories": ["2023", "2024"],
                "series": [{"name": "Выручка", "values": [1, 2]}],
            },
        })
    )
    assert data["slides"][0]["layout"] == "chart"
    assert data["slides"][0]["chart"]["categories"] == ["2023", "2024"]


def test_comparison_missing_a_side_downgrades_to_cards():
    data = normalize_presentation(
        _deck({
            "type": "content", "layout": "comparison", "title": "X",
            "comparison": {"left": {"title": "Только левая"}},
        })
    )
    assert data["slides"][0]["layout"] == "cards"


def test_team_needs_at_least_two_members():
    data = normalize_presentation(
        _deck({"type": "content", "layout": "team", "title": "X", "team": [{"name": "Один"}]})
    )
    assert data["slides"][0]["layout"] == "text"


def test_big_photo_without_image_query_downgrades_to_text():
    data = normalize_presentation(_deck({"type": "content", "layout": "big_photo", "title": "X"}))
    assert data["slides"][0]["layout"] == "text"


@pytest.mark.parametrize(
    "visual_format,expected_max_bullets",
    [("image_only", 1), ("image_heavy", 2), ("balanced", 4)],
)
def test_visual_format_clamps_bullet_count(visual_format, expected_max_bullets):
    data = normalize_presentation(
        _deck({
            "type": "content", "layout": "text", "title": "X",
            "bullets": ["Раз", "Два", "Три", "Четыре", "Пять"],
        }),
        visual_format=visual_format,
    )
    assert len(data["slides"][0]["bullets"]) == expected_max_bullets


def test_body_field_is_kept_and_cleaned():
    data = normalize_presentation(
        _deck({
            "type": "content", "layout": "text", "title": "X",
            "body": "Стоит отметить, что это связный абзац с водой в начале.",
        })
    )
    assert data["slides"][0]["body"] == "Это связный абзац с водой в начале."


def test_non_dict_slide_items_are_skipped():
    data = normalize_presentation(_deck("not a dict", {"type": "content", "title": "Valid"}))
    assert len(data["slides"]) == 1
    assert data["slides"][0]["title"] == "Valid"
