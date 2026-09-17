from text_cleanup import clean_text


def test_removes_filler_phrase_at_start():
    assert clean_text("Стоит отметить, что рынок растёт.") == "Рынок растёт."


def test_strips_markdown_bold():
    assert clean_text("**Жирный** текст") == "Жирный текст"


def test_strips_leading_list_marker():
    assert clean_text("- пункт списка") == "Пункт списка"
    assert clean_text("1. пронумерованный пункт") == "Пронумерованный пункт"


def test_collapses_repeated_punctuation_and_whitespace():
    assert clean_text("Слишком много!!!  пробелов") == "Слишком много! пробелов"


def test_single_idea_keeps_only_first_sentence_when_long():
    long_text = (
        "Первое предложение специально сделано настолько длинным, чтобы гарантированно "
        "превысить порог в сто сорок символов и запустить логику обрезки. "
        "Второе предложение должно быть отрезано полностью."
    )
    result = clean_text(long_text)
    assert result.startswith("Первое предложение")
    assert "Второе предложение" not in result


def test_single_idea_false_keeps_multiple_sentences():
    long_text = (
        "Первое предложение достаточно длинное само по себе, чтобы превысить порог. "
        "Второе предложение должно остаться, так как это связный абзац."
    )
    result = clean_text(long_text, single_idea=False)
    assert "Второе предложение" in result


def test_empty_and_none_like_input():
    assert clean_text("") == ""
