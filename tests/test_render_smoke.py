"""Дымовой тест рендера: использует тот же синтетический набор слайдов, что
и scripts/render_check.py, чтобы одну и ту же проверку можно было запускать
и вручную (`python scripts/render_check.py`), и в CI через pytest."""

import asyncio
import sys
from pathlib import Path

import pytest


def _cleanup(*paths: str) -> None:
    for p in paths:
        Path(p).unlink(missing_ok=True)

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "scripts"))

import render_check  # noqa: E402
from models import normalize_presentation  # noqa: E402
from pptx_builder import build_presentation  # noqa: E402


def test_render_all_layouts_without_background(tmp_path):
    async def run():
        data = normalize_presentation(render_check.RAW, visual_format="balanced")
        path = await build_presentation(data, user_id=1, theme_preset="warm")
        return path

    path = asyncio.run(run())
    issues = render_check._check_textbox_overflow(path)
    _cleanup(path)
    assert not issues, f"Text boxes overflow the slide bounds: {issues}"


def test_render_with_custom_backgrounds(tmp_path):
    from PIL import Image

    light_bg = tmp_path / "light.jpg"
    dark_bg = tmp_path / "dark.jpg"
    Image.new("RGB", (800, 500), (235, 230, 220)).save(light_bg)
    Image.new("RGB", (800, 500), (15, 18, 25)).save(dark_bg)

    async def run():
        data = normalize_presentation(render_check.RAW, visual_format="balanced")
        light_path = await build_presentation(data, user_id=2, theme_preset="corporate", background_path=light_bg)
        dark_path = await build_presentation(data, user_id=3, theme_preset="dark", background_path=dark_bg)
        return light_path, dark_path

    light_path, dark_path = asyncio.run(run())
    issues = render_check._check_textbox_overflow(light_path) + render_check._check_textbox_overflow(dark_path)
    _cleanup(light_path, dark_path)
    assert not issues


@pytest.mark.parametrize("visual_format", ["image_only", "image_heavy", "balanced"])
def test_render_each_visual_format(visual_format):
    async def run():
        data = normalize_presentation(render_check.RAW, visual_format=visual_format)
        return await build_presentation(data, user_id=4, theme_preset="warm")

    path = asyncio.run(run())
    issues = render_check._check_textbox_overflow(path)
    _cleanup(path)
    assert not issues
