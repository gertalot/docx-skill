"""tests/test_brand.py"""
import json
from pathlib import Path

import pytest
from docx_builder.brand import Brand


class TestBrand:
    def test_create_brand(self):
        brand = Brand(
            primary="#1C2541",
            accent="#6FFFE9",
            body="#3E465E",
            font_family="Helvetica",
        )
        assert brand.primary == "#1C2541"
        assert brand.font_family == "Helvetica"

    def test_resolve_weight(self):
        brand = Brand(
            primary="#000000",
            accent="#000000",
            body="#000000",
            font_family="Helvetica",
        )
        regular = brand.resolve_weight(400)
        assert isinstance(regular, str)
        assert "Helvetica" in regular

    def test_save_and_load(self, tmp_path):
        original = Brand(
            primary="#1C2541",
            accent="#6FFFE9",
            body="#3E465E",
            font_family="Helvetica",
            emphasis_weight=500,
            heading_weight=700,
        )
        path = tmp_path / "brand.json"
        original.save(path)
        loaded = Brand.load(path)
        assert loaded.primary == original.primary
        assert loaded.accent == original.accent
        assert loaded.font_family == original.font_family
        assert loaded.emphasis_weight == original.emphasis_weight

    def test_save_creates_valid_json(self, tmp_path):
        brand = Brand(
            primary="#000", accent="#fff", body="#333", font_family="Helvetica"
        )
        path = tmp_path / "brand.json"
        brand.save(path)
        data = json.loads(path.read_text())
        assert "primary" in data
        assert "font_family" in data

    def test_rgb_conversion(self):
        brand = Brand(
            primary="#1C2541", accent="#6FFFE9", body="#3E465E",
            font_family="Helvetica",
        )
        rgb = brand.primary_rgb
        assert rgb.red == 0x1C
        assert rgb.green == 0x25
        assert rgb.blue == 0x41

    def test_resolve_weight_missing_font_falls_back(self):
        brand = Brand(
            primary="#000", accent="#000", body="#000",
            font_family="Definitely Not A Real Font",
        )
        result = brand.resolve_weight(400)
        assert result == "Definitely Not A Real Font"

    def test_rgb_to_docx(self):
        from docx.shared import RGBColor as DocxRGBColor
        brand = Brand(
            primary="#1C2541", accent="#6FFFE9", body="#3E465E",
            font_family="Helvetica",
        )
        docx_rgb = brand.primary_rgb.to_docx()
        assert isinstance(docx_rgb, DocxRGBColor)

    def test_logo_path_none_by_default(self):
        brand = Brand(
            primary="#000", accent="#fff", body="#333", font_family="Helvetica"
        )
        assert brand.logo_path is None
