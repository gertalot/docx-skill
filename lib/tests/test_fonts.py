"""tests/test_fonts.py"""
import pytest

from docx_builder.fonts import FontDiscovery, FontFamily


@pytest.fixture(autouse=True)
def _clear_font_cache():
    """Clear FontDiscovery cache before each test to avoid state leakage."""
    FontDiscovery.clear_cache()
    yield
    FontDiscovery.clear_cache()


class TestFontFamily:
    def test_closest_weight_exact_match(self):
        family = FontFamily(
            name="Test Sans",
            weights={
                400: "Test Sans",
                500: "Test Sans Medium",
                700: "Test Sans Bold",
            },
        )
        assert family.closest(500) == "Test Sans Medium"

    def test_closest_weight_rounds_to_nearest(self):
        family = FontFamily(
            name="Test Sans",
            weights={400: "Test Sans", 700: "Test Sans Bold"},
        )
        assert family.closest(500) == "Test Sans"
        assert family.closest(600) == "Test Sans Bold"

    def test_closest_weight_single_variant(self):
        family = FontFamily(name="Mono", weights={400: "Mono Regular"})
        assert family.closest(700) == "Mono Regular"

    def test_weight_names(self):
        family = FontFamily(
            name="Test Sans",
            weights={
                400: "Test Sans",
                500: "Test Sans Medium",
                700: "Test Sans Bold",
            },
        )
        assert family.weight_name(400) == "Regular"
        assert family.weight_name(500) == "Medium"
        assert family.weight_name(700) == "Bold"


class TestFontDiscovery:
    def test_list_families_returns_nonempty(self):
        families = FontDiscovery.list_families()
        assert len(families) > 0

    def test_get_family_known_font(self):
        family = FontDiscovery.get_family("Helvetica")
        assert family is not None
        assert family.name == "Helvetica"
        assert 400 in family.weights

    def test_get_family_unknown_returns_none(self):
        family = FontDiscovery.get_family("Definitely Not A Real Font")
        assert family is None

    def test_is_available_system_font(self):
        assert FontDiscovery.is_available("Helvetica")

    def test_is_available_missing_font(self):
        assert not FontDiscovery.is_available("Definitely Not A Real Font")
