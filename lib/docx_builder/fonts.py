"""Font discovery: scan system fonts, parse weights, query families.

macOS only. Scans /System/Library/Fonts, /Library/Fonts, ~/Library/Fonts.
"""
from __future__ import annotations

import logging
import sys
from dataclasses import dataclass, field
from pathlib import Path

from fontTools.ttLib import TTFont, TTCollection

if sys.platform != "darwin":
    raise NotImplementedError("Font discovery only supports macOS")

log = logging.getLogger(__name__)

FONT_DIRS = [
    Path("/System/Library/Fonts"),
    Path("/System/Library/Fonts/Supplemental"),
    Path("/Library/Fonts"),
    Path.home() / "Library" / "Fonts",
]

WEIGHT_NAMES = {
    100: "Thin",
    200: "Ultra Light",
    300: "Light",
    400: "Regular",
    500: "Medium",
    600: "Demi Bold",
    700: "Bold",
    800: "Extra Bold",
    900: "Black",
}

_WEIGHT_BUCKETS = [100, 200, 300, 400, 500, 600, 700, 800, 900]


def _weight_class_to_bucket(weight_class: int) -> int:
    """Round an OS/2 usWeightClass value to the nearest standard bucket."""
    return min(_WEIGHT_BUCKETS, key=lambda b: abs(b - weight_class))


def _extract_font_info(font: TTFont) -> tuple[str, int, str] | None:
    """Extract (family_name, weight, full_name) from a TTFont.

    Returns None if essential metadata is missing or font is italic.
    """
    name_table = font.get("name")
    os2_table = font.get("OS/2")
    if not name_table or not os2_table:
        return None

    def get_name(name_id: int) -> str | None:
        for plat in (3, 1):
            record = name_table.getName(
                name_id,
                plat,
                1 if plat == 3 else 0,
                0x409 if plat == 3 else 0,
            )
            if record:
                return str(record)
        return None

    family = get_name(1)
    full_name = get_name(4)
    if not family or not full_name:
        return None

    if os2_table.fsSelection & 0x01:  # ITALIC flag
        return None

    weight = _weight_class_to_bucket(os2_table.usWeightClass)
    return family, weight, full_name


@dataclass
class FontFamily:
    """A font family with its available weight variants."""

    name: str
    weights: dict[int, str] = field(default_factory=dict)

    def closest(self, target: int) -> str:
        """Return the font name for the weight closest to target."""
        if not self.weights:
            raise ValueError(f"FontFamily {self.name!r} has no weight variants")
        best = min(self.weights, key=lambda w: abs(w - target))
        return self.weights[best]

    def weight_name(self, weight: int) -> str:
        """Return the human-readable name for a weight value."""
        return WEIGHT_NAMES.get(weight, f"Weight {weight}")

    def __repr__(self) -> str:
        variants = ", ".join(
            f"{self.weight_name(w)} ({w})" for w in sorted(self.weights)
        )
        return f"FontFamily({self.name!r}, [{variants}])"


class FontDiscovery:
    """Scan system fonts and query families."""

    _cache: dict[str, FontFamily] | None = None

    @classmethod
    def _scan(cls) -> dict[str, FontFamily]:
        """Scan font directories and build the family cache."""
        if cls._cache is not None:
            return cls._cache

        families: dict[str, FontFamily] = {}

        for font_dir in FONT_DIRS:
            if not font_dir.exists():
                continue
            for path in font_dir.iterdir():
                suffix = path.suffix.lower()
                if suffix not in (".ttf", ".otf", ".ttc"):
                    continue
                try:
                    if suffix == ".ttc":
                        with TTCollection(str(path)) as collection:
                            for font in collection.fonts:
                                info = _extract_font_info(font)
                                if info is None:
                                    continue
                                family_name, weight, full_name = info
                                if family_name not in families:
                                    families[family_name] = FontFamily(
                                        name=family_name,
                                    )
                                families[family_name].weights[weight] = full_name
                    else:
                        with TTFont(str(path)) as font:
                            info = _extract_font_info(font)
                            if info is not None:
                                family_name, weight, full_name = info
                                if family_name not in families:
                                    families[family_name] = FontFamily(
                                        name=family_name,
                                    )
                                families[family_name].weights[weight] = full_name
                except Exception:
                    log.debug("Skipping unreadable font: %s", path)
                    continue

        cls._cache = families
        return families

    @classmethod
    def list_families(cls) -> list[FontFamily]:
        """List all font families on the system."""
        return sorted(cls._scan().values(), key=lambda f: f.name)

    @classmethod
    def get_family(cls, name: str) -> FontFamily | None:
        """Get a font family by name. Case-sensitive."""
        return cls._scan().get(name)

    @classmethod
    def is_available(cls, font_name: str) -> bool:
        """Check if a font family is installed."""
        return cls.get_family(font_name) is not None

    @classmethod
    def clear_cache(cls) -> None:
        """Clear the font cache (useful for testing)."""
        cls._cache = None
