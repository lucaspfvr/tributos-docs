"""Color roles and local appearance preferences."""

from dataclasses import dataclass


@dataclass(frozen=True)
class Palette:
    name: str
    nav: str
    active: str
    bg: str
    paper: str
    ink: str
    muted: str
    line: str
    accent: str
    tint: str
    hover: str
    canvas: str


PALETTES = {
    "operation": Palette("Azul operacional", "#eceef1", "#dce8f8", "#f1f4f8", "#ffffff", "#192b42",
                         "#465a71", "#c7d4e4", "#116aad", "#e9f0fa", "#0c538a", "#dbe4ee"),
    "paper": Palette("Papel e tinta", "#faf7f2", "#f6e3d3", "#f5f0e9", "#fffdfa", "#28211d",
                     "#68594e", "#ded3c7", "#a64216", "#faede2", "#843210", "#e8dfd5"),
    "blue": Palette("Azul névoa", "#f3f7fd", "#dae7fa", "#edf2f9", "#ffffff", "#19273d",
                     "#50617a", "#d1ddeb", "#2456a0", "#e7effd", "#1a447f", "#dce4ef"),
    "sage": Palette("Verde sálvia", "#f4f7f1", "#dfeadc", "#eef2eb", "#fdfefa", "#202d23",
                     "#4d6151", "#cedacd", "#2d6548", "#e5efe2", "#214e37", "#dbe4d9"),
    "plum": Palette("Ameixa", "#f8f3f8", "#edddee", "#f3edf4", "#fffaff", "#302136",
                     "#66546c", "#dfd1e2", "#793877", "#f5e7f4", "#62275f", "#e7dce8"),
}
DEFAULT_THEME = "operation"
