"""Brand configuration: colours, fonts, and logo for document generation.

Stores hex colour strings, resolves font weights via FontDiscovery,
and persists to/from JSON for reuse across sessions.
"""
from __future__ import annotations

import json
from dataclasses import dataclass, field
from io import BytesIO
from pathlib import Path

from docx_builder.fonts import FontDiscovery


@dataclass(frozen=True)
class RGBColor:
    """An RGB colour with named channel access."""

    red: int
    green: int
    blue: int


def _hex_to_rgb(hex_str: str) -> RGBColor:
    """Convert a hex colour string (with or without #) to RGBColor."""
    h = hex_str.lstrip("#")
    if len(h) == 3:
        h = h[0] * 2 + h[1] * 2 + h[2] * 2
    if len(h) != 6:
        raise ValueError(f"Invalid hex colour: {hex_str!r}")
    return RGBColor(
        red=int(h[0:2], 16),
        green=int(h[2:4], 16),
        blue=int(h[4:6], 16),
    )


def _strip_hash(hex_str: str) -> str:
    """Strip the leading # from a hex colour string."""
    return hex_str.lstrip("#")


@dataclass
class Brand:
    """Brand configuration with colour, font, and logo settings.

    Attributes:
        primary: Primary colour as hex string (e.g. "#1C2541").
        accent: Accent colour as hex string.
        body: Body text colour as hex string.
        light_bg: Light background colour as hex string.
        border: Border colour as hex string.
        font_family: Font family name (must be installed on the system).
        emphasis_weight: Font weight for emphasis text (default 500).
        heading_weight: Font weight for headings (default 700).
        logo_path: Optional path to logo file (PNG or SVG).
    """

    primary: str
    accent: str
    body: str
    light_bg: str = "#F7F9FC"
    border: str = "#E2E8F0"
    font_family: str = "Helvetica"
    emphasis_weight: int = 500
    heading_weight: int = 700
    logo_path: Path | None = None

    _weight_cache: dict[int, str] = field(
        default_factory=dict, init=False, repr=False, compare=False,
    )

    # --- Font resolution ---

    def resolve_weight(self, weight: int) -> str:
        """Resolve a weight number to an installed font face name.

        Uses FontDiscovery to find the closest matching weight in the
        configured font family. Results are cached per Brand instance.
        """
        if weight in self._weight_cache:
            return self._weight_cache[weight]
        family = FontDiscovery.get_family(self.font_family)
        if family is None:
            raise ValueError(
                f"Font family {self.font_family!r} not found on this system"
            )
        face_name = family.closest(weight)
        self._weight_cache[weight] = face_name
        return face_name

    @property
    def font_regular(self) -> str:
        """Font face name for regular (400) weight."""
        return self.resolve_weight(400)

    @property
    def font_emphasis(self) -> str:
        """Font face name for emphasis weight."""
        return self.resolve_weight(self.emphasis_weight)

    @property
    def font_heading(self) -> str:
        """Font face name for heading weight."""
        return self.resolve_weight(self.heading_weight)

    # --- Colour properties ---

    @property
    def primary_rgb(self) -> RGBColor:
        """Primary colour as RGBColor."""
        return _hex_to_rgb(self.primary)

    @property
    def accent_rgb(self) -> RGBColor:
        """Accent colour as RGBColor."""
        return _hex_to_rgb(self.accent)

    @property
    def body_rgb(self) -> RGBColor:
        """Body text colour as RGBColor."""
        return _hex_to_rgb(self.body)

    @property
    def primary_hex(self) -> str:
        """Primary colour without # prefix."""
        return _strip_hash(self.primary)

    @property
    def accent_hex(self) -> str:
        """Accent colour without # prefix."""
        return _strip_hash(self.accent)

    @property
    def light_bg_hex(self) -> str:
        """Light background colour without # prefix."""
        return _strip_hash(self.light_bg)

    @property
    def border_hex(self) -> str:
        """Border colour without # prefix."""
        return _strip_hash(self.border)

    # --- Logo ---

    def logo_png(self, width_px: int = 200) -> BytesIO | None:
        """Return the logo as a PNG BytesIO, or None if no logo is set.

        If the logo is an SVG, it is converted to PNG via cairosvg.
        If the logo is already a raster image, it is resized to width_px.
        """
        if self.logo_path is None:
            return None
        path = Path(self.logo_path)
        if not path.exists():
            raise FileNotFoundError(f"Logo file not found: {path}")

        if path.suffix.lower() == ".svg":
            import cairosvg

            png_bytes = cairosvg.svg2png(
                url=str(path), output_width=width_px,
            )
            buf = BytesIO(png_bytes)
            buf.seek(0)
            return buf

        from PIL import Image

        img = Image.open(path)
        aspect = img.height / img.width
        new_height = int(width_px * aspect)
        img = img.resize((width_px, new_height), Image.LANCZOS)
        buf = BytesIO()
        img.save(buf, format="PNG")
        buf.seek(0)
        return buf

    # --- Serialization ---

    def save(self, path: Path) -> None:
        """Save brand configuration to a JSON file."""
        data = {
            "primary": self.primary,
            "accent": self.accent,
            "body": self.body,
            "light_bg": self.light_bg,
            "border": self.border,
            "font_family": self.font_family,
            "emphasis_weight": self.emphasis_weight,
            "heading_weight": self.heading_weight,
            "logo_path": str(self.logo_path) if self.logo_path else None,
        }
        Path(path).write_text(json.dumps(data, indent=2))

    @classmethod
    def load(cls, path: Path) -> Brand:
        """Load brand configuration from a JSON file."""
        data = json.loads(Path(path).read_text())
        logo = data.get("logo_path")
        return cls(
            primary=data["primary"],
            accent=data["accent"],
            body=data["body"],
            light_bg=data.get("light_bg", "#F7F9FC"),
            border=data.get("border", "#E2E8F0"),
            font_family=data["font_family"],
            emphasis_weight=data.get("emphasis_weight", 500),
            heading_weight=data.get("heading_weight", 700),
            logo_path=Path(logo) if logo else None,
        )
