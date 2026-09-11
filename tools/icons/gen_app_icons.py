#!/usr/bin/env python3
"""Generate Airy application icons from the brand assets.

Inputs:

  tools/icons/source/airy_without_bg_no_text.png — 1254x1254 RGBA: the gradient
      mark alone (no wordmark, transparent background), committed so the icon
      set regenerates without the out-of-repo brand-assets directory.
  brand-assets/airy_logo.png               — full logo (mark + wordmark) on white,
      reused as docs/assets/airy_logo.png for the README banner. Outside the
      repo (only the README banner needs it).

Outputs:

  apps/shell/build/icon.png            — 1024x1024 master, ~8% transparent
                                         padding (electron-builder fallback
                                         and source for auto conversions)
  apps/shell/build/icon-mac.png        — 1024x1024 with the macOS icon grid
                                         margin (824/1024 content, same
                                         treatment as the file-type icons)
  apps/shell/build/icon.icns           — regenerated from icon-mac (PNG
                                         entries, macOS 10.7+)
  apps/shell/build/icon.ico            — regenerated from icon.png (PNG
                                         entries, Vista+)
  apps/shell/build/icons/*.png         — hicolor set for deb/rpm (linux
                                         `icon: 'build/icons'` is a directory)
  apps/docs/build/{icon.png,icon-mac.png,icon.icns,icon.ico}
                                       — same files for the legacy standalone
                                         docs app
  apps/shell/src/renderer/src/assets/app-icon.png
                                       — 1024x1024 copy of the master icon for
                                         the onboarding slide (shown at 60px)
  packages/ui/src/assets/airy-mark.png — 256x256 tight mark for the shared
                                         AiryMark component (all renderers)
  docs/assets/airy_logo.png            — full logo, width 960, for README

icns/ico are hand-rolled containers here (pure stdlib + PIL): icns is just a
type/length-framed bag of PNGs (ic07..ic14), ico is the same PNG-entry format
tools/gen-file-association-icons.mjs writes — no Xcode/iconutil needed, so the
shell app icon regenerates on Linux CI too.

Usage:  python3 tools/icons/gen_app_icons.py [--verify]
  --verify only prints the current size/mode of every generated file.
"""

from __future__ import annotations

import io
import struct
import sys
from pathlib import Path

from PIL import Image

REPO_ROOT = Path(__file__).resolve().parents[2]
BRAND_ASSETS = REPO_ROOT.parent / "brand-assets"
MARK_PNG = REPO_ROOT / "tools" / "icons" / "source" / "airy_without_bg_no_text.png"
LOGO_PNG = BRAND_ASSETS / "airy_logo.png"

SHELL_BUILD = REPO_ROOT / "apps" / "shell" / "build"
DOCS_BUILD = REPO_ROOT / "apps" / "docs" / "build"
LINUX_ICON_SET = SHELL_BUILD / "icons"
SHELL_RENDERER_ICON = REPO_ROOT / "apps" / "shell" / "src" / "renderer" / "src" / "assets" / "app-icon.png"
UI_MARK = REPO_ROOT / "packages" / "ui" / "src" / "assets" / "airy-mark.png"
DOCS_ASSETS = REPO_ROOT / "docs" / "assets"

MASTER_SIZE = 1024
# tight mark asset the shared AiryMark component renders in the UI
UI_MARK_SIZE = 256
# transparent margin around the mark, fraction of the canvas side
APP_PADDING = 0.08
# macOS app icons sit on a standard grid with 824/1024 content (the same
# ratio tools/gen-file-association-icons.mjs uses for document icons)
MAC_CONTENT_RATIO = 824 / 1024
LINUX_SET_SIZES = [16, 32, 48, 64, 128, 256, 512, 1024]
WIN_ICO_SIZES = [16, 24, 32, 48, 64, 128, 256]
# icns entry types -> pixel size (PNG-compressed entries, macOS 10.7+);
# @2x slots map onto the plain types, ic10 covers 512@2x/1024. This is the
# same ic07..ic14 set iconutil — and electron-builder's own png -> icns
# conversion — produces; 16px slots render from ic11 (16@2x).
ICNS_ENTRIES = [
    (b"ic07", 128),
    (b"ic08", 256),
    (b"ic09", 512),
    (b"ic10", 1024),
    (b"ic11", 32),
    (b"ic12", 64),
    (b"ic13", 256),
    (b"ic14", 512),
]
README_LOGO_WIDTH = 960


def load_mark() -> Image.Image:
    """Tight-crop the icon-only mark.

    The source is the bare gradient mark on a transparent background (no
    wordmark anymore), so its bbox is simply the alpha bounding box.
    """
    with Image.open(MARK_PNG) as img:
        rgba = img.convert("RGBA")
        bbox = rgba.getchannel("A").getbbox()
        if bbox is None:
            raise SystemExit(f"{MARK_PNG}: fully transparent")
        return rgba.crop(bbox)


def compose(mark: Image.Image, size: int, content_ratio: float) -> Image.Image:
    """Center the mark in a transparent square, content at `content_ratio`."""
    canvas = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    content = round(size * content_ratio)
    scale = min(content / mark.width, content / mark.height)
    resized = mark.resize(
        (max(1, round(mark.width * scale)), max(1, round(mark.height * scale))),
        Image.Resampling.LANCZOS,
    )
    canvas.paste(
        resized,
        ((size - resized.width) // 2, (size - resized.height) // 2),
        resized,
    )
    return canvas


def png_bytes(img: Image.Image) -> bytes:
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    return buf.getvalue()


def build_icns(master: Image.Image) -> bytes:
    """icns container: "icns" + BE total length, then type-framed PNGs."""
    entries = []
    for icon_type, size in ICNS_ENTRIES:
        png = png_bytes(master.resize((size, size), Image.Resampling.LANCZOS))
        entries.append(icon_type + struct.pack(">I", 8 + len(png)) + png)
    total = 8 + sum(len(entry) for entry in entries)
    return b"icns" + struct.pack(">I", total) + b"".join(entries)


def build_ico(master: Image.Image) -> bytes:
    """ICO container with PNG-compressed entries (supported since Vista)."""
    entries = []
    for size in WIN_ICO_SIZES:
        entries.append(
            (size, png_bytes(master.resize((size, size), Image.Resampling.LANCZOS)))
        )
    header = struct.pack("<HHH", 0, 1, len(entries))
    directory = b""
    blob = b""
    offset = 6 + 16 * len(entries)
    for size, png in entries:
        byte = 0 if size >= 256 else size
        directory += struct.pack("<BBBBHHII", byte, byte, 0, 0, 1, 32, len(png), offset)
        blob += png
        offset += len(png)
    return header + directory + blob


def report(paths: list[Path]) -> None:
    for path in sorted(paths):
        with Image.open(path) as img:
            print(f"{path.relative_to(REPO_ROOT)}: {img.size[0]}x{img.size[1]} {img.mode}")


def main() -> None:
    if "--verify" in sys.argv[1:]:
        targets = [
            SHELL_BUILD / "icon.png",
            SHELL_BUILD / "icon-mac.png",
            SHELL_BUILD / "icon.icns",
            SHELL_BUILD / "icon.ico",
            *[LINUX_ICON_SET / f"{size}x{size}.png" for size in LINUX_SET_SIZES],
            DOCS_BUILD / "icon.png",
            DOCS_BUILD / "icon-mac.png",
            DOCS_BUILD / "icon.icns",
            DOCS_BUILD / "icon.ico",
            SHELL_RENDERER_ICON,
            UI_MARK,
            DOCS_ASSETS / "airy_logo.png",
        ]
        report(targets)
        return

    for required in (MARK_PNG, LOGO_PNG):
        if not required.exists():
            raise SystemExit(f"missing brand asset: {required}")

    mark = load_mark()
    app_master = compose(mark, MASTER_SIZE, 1.0 - 2 * APP_PADDING)
    mac_master = compose(mark, MASTER_SIZE, MAC_CONTENT_RATIO)

    written: list[Path] = []

    def write(path: Path, data: bytes) -> None:
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_bytes(data)
        written.append(path)

    write(SHELL_BUILD / "icon.png", png_bytes(app_master))
    write(SHELL_BUILD / "icon-mac.png", png_bytes(mac_master))
    write(SHELL_BUILD / "icon.icns", build_icns(mac_master))
    write(SHELL_BUILD / "icon.ico", build_ico(app_master))
    for size in LINUX_SET_SIZES:
        write(
            LINUX_ICON_SET / f"{size}x{size}.png",
            png_bytes(app_master.resize((size, size), Image.Resampling.LANCZOS)),
        )

    # legacy standalone docs app ships the same brand icon set
    for name in ("icon.png", "icon-mac.png"):
        write(
            DOCS_BUILD / name,
            png_bytes(app_master if name == "icon.png" else mac_master),
        )
    write(DOCS_BUILD / "icon.icns", build_icns(mac_master))
    write(DOCS_BUILD / "icon.ico", build_ico(app_master))

    # onboarding slide 1 shows the app icon (CSS caps it at 60px)
    write(SHELL_RENDERER_ICON, png_bytes(app_master))

    # tight mark for the shared AiryMark component across the renderers
    write(UI_MARK, png_bytes(compose(mark, UI_MARK_SIZE, 1.0)))

    # full logo (mark + wordmark on white) for the README banner
    with Image.open(LOGO_PNG) as logo:
        height = round(logo.height * README_LOGO_WIDTH / logo.width)
        resized = logo.resize((README_LOGO_WIDTH, height), Image.Resampling.LANCZOS)
        write(DOCS_ASSETS / "airy_logo.png", png_bytes(resized))

    report(written)


if __name__ == "__main__":
    main()
