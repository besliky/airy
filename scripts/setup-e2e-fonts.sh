#!/usr/bin/env bash
# Install the document fonts the e2e pixel baselines are pinned to.
#
# The docs-visual suite (e2e/docs-visual.spec.ts) compares page renders
# pixel-for-pixel against baselines that CI rendered with the font stack
# pinned in .github/workflows/ci.yml ("Install document fonts (visual
# baselines)"): Carlito/Caladea (metric-compatible with Calibri/Cambria)
# and Noto CJK. Running the visual specs without them produces pixel-diff
# failures that do not reproduce on CI, so install the same packages
# locally before `npm run test:e2e:xvfb`.
#
# Two layers:
# 1. Distro packages — the same `apt-get install` CI runs. Idempotent:
#    already-installed packages are detected via dpkg and skipped. On
#    ubuntu-22.04 (the CI runner image) this alone matches CI.
# 2. Exact-CI-build overlay — newer Ubuntu/Debian releases ship redesigned
#    Carlito/Caladea builds (same metrics, different glyph outlines), which
#    still diff against the baselines. When the installed files are not the
#    CI builds, the script downloads the ubuntu-22.04 .debs pinned by hash,
#    extracts the faces into e2e/.e2e-fonts/ (gitignored, no root needed)
#    and writes a fontconfig include there. scripts/run-e2e-xvfb.sh picks
#    the overlay up via FONTCONFIG_FILE, so only the e2e run is affected.
set -euo pipefail

# CI (ubuntu-22.04) builds; .deb hashes pin the download against tampering.
CARLITO_DEB_URL='http://archive.ubuntu.com/ubuntu/pool/universe/f/fonts-crosextra-carlito/fonts-crosextra-carlito_20130920-1.1_all.deb'
CARLITO_DEB_SHA='7385475cde807e1363c3361976576571870373032466c7f525d5900852b6f420'
CALADEA_DEB_URL='http://archive.ubuntu.com/ubuntu/pool/universe/f/fonts-crosextra-caladea/fonts-crosextra-caladea_20130214-2.1_all.deb'
CALADEA_DEB_SHA='1330d25dfa5bab2e9b712b4950d2855cdb63a2b2b9451e2ffb93618c77e1f242'
# sha256 of the Carlito-Regular.ttf shipped by the CI .deb: this one file is
# enough to tell the CI build from the redesigned post-2023 distro builds.
CARLITO_CI_SHA='b4ff23ba370cc95a3c349336b73f9c28514a1371210f89832efc85c4b1ea7131'

PACKAGES=(fonts-crosextra-carlito fonts-crosextra-caladea fonts-noto-cjk)
REPO_ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
OVERLAY_DIR="$REPO_ROOT/e2e/.e2e-fonts"

if ! command -v dpkg-query >/dev/null 2>&1 || ! command -v apt-get >/dev/null 2>&1; then
  echo "error: this script targets Debian/Ubuntu (dpkg/apt-get), matching the CI runner." >&2
  echo "Install the font packages with your distro's package manager instead:" >&2
  echo "  ${PACKAGES[*]}" >&2
  exit 1
fi

missing=()
for pkg in "${PACKAGES[@]}"; do
  if [ "$(dpkg-query -W -f='${Status}' "$pkg" 2>/dev/null)" != "install ok installed" ]; then
    missing+=("$pkg")
  fi
done

if [ "${#missing[@]}" -gt 0 ]; then
  sudo=()
  if [ "$(id -u)" -ne 0 ]; then
    if ! command -v sudo >/dev/null 2>&1; then
      echo "error: sudo is required to install ${missing[*]} (no sudo on PATH, not running as root)." >&2
      echo "Install the packages manually, then re-run this script:" >&2
      echo "  apt-get install -y --no-install-recommends ${missing[*]}" >&2
      exit 1
    fi
    sudo=(sudo)
  fi
  echo "Installing missing e2e font packages: ${missing[*]}"
  "${sudo[@]}" apt-get update
  "${sudo[@]}" apt-get install -y --no-install-recommends "${missing[@]}"
else
  echo "distro e2e font packages already installed: ${PACKAGES[*]}"
fi

# Layer 2: skip when the distro already ships the CI builds (ubuntu-22.04).
system_carlito=/usr/share/fonts/truetype/crosextra/Carlito-Regular.ttf
if [ -f "$system_carlito" ] && [ "$(sha256sum "$system_carlito" | cut -d' ' -f1)" = "$CARLITO_CI_SHA" ]; then
  rm -rf "$OVERLAY_DIR"
  echo "installed Carlito matches the CI build; no font overlay needed"
  exit 0
fi

if [ ! -f "$OVERLAY_DIR/fonts/Carlito-Regular.ttf" ] ||
  [ "$(sha256sum "$OVERLAY_DIR/fonts/Carlito-Regular.ttf" | cut -d' ' -f1)" != "$CARLITO_CI_SHA" ]; then
  if ! command -v curl >/dev/null 2>&1 || ! command -v dpkg-deb >/dev/null 2>&1; then
    echo "error: curl and dpkg-deb are required to pin the exact CI font builds." >&2
    exit 1
  fi
  echo "distro Carlito/Caladea builds differ from CI; pinning the CI builds into e2e/.e2e-fonts/"
  mkdir -p "$OVERLAY_DIR/deb" "$OVERLAY_DIR/fonts"
  deb_tmp="$OVERLAY_DIR/deb"
  fetch_pinned_deb() {
    local url=$1 want_sha=$2 out
    out="$deb_tmp/$(basename "$url")"
    if [ -f "$out" ] && [ "$(sha256sum "$out" | cut -d' ' -f1)" = "$want_sha" ]; then
      return 0
    fi
    curl -fsSL -o "$out" "$url"
    if [ "$(sha256sum "$out" | cut -d' ' -f1)" != "$want_sha" ]; then
      echo "error: downloaded $(basename "$url") does not match the pinned sha256" >&2
      rm -f "$out"
      exit 1
    fi
  }
  fetch_pinned_deb "$CARLITO_DEB_URL" "$CARLITO_DEB_SHA"
  fetch_pinned_deb "$CALADEA_DEB_URL" "$CALADEA_DEB_SHA"
  rm -rf "$OVERLAY_DIR/extract"
  mkdir -p "$OVERLAY_DIR/extract"
  dpkg-deb -x "$deb_tmp/$(basename "$CARLITO_DEB_URL")" "$OVERLAY_DIR/extract"
  dpkg-deb -x "$deb_tmp/$(basename "$CALADEA_DEB_URL")" "$OVERLAY_DIR/extract"
  find "$OVERLAY_DIR/extract" -name '*.ttf' -exec cp -t "$OVERLAY_DIR/fonts" {} +
  rm -rf "$OVERLAY_DIR/extract"
fi

# Regenerate the fontconfig include each run: it must reject every installed
# Carlito/Caladea copy that is not the pinned one (fc-list knows them all,
# including per-user font dirs) and search the pinned dir first.
{
  echo '<?xml version="1.0"?>'
  echo '<!DOCTYPE fontconfig SYSTEM "fonts.dtd">'
  echo '<fontconfig>'
  echo "  <dir>$OVERLAY_DIR/fonts</dir>"
  echo '  <include ignore_missing="yes">/etc/fonts/fonts.conf</include>'
  echo '  <selectfont>'
  echo '    <rejectfont>'
  echo '      <glob>/usr/share/fonts/truetype/crosextra/*</glob>'
  echo '    </rejectfont>'
  fc-list : family file | awk -F': ' '$2 == "Carlito" || $2 == "Caladea" { print $1 }' |
    grep -v "^$OVERLAY_DIR/fonts/" | sort -u |
    while IFS= read -r file; do
      printf '    <rejectfont>\n      <glob>%s</glob>\n    </rejectfont>\n' "$file"
    done
  echo '  </selectfont>'
  echo '</fontconfig>'
} >"$OVERLAY_DIR/fontconfig.conf"

echo "font overlay ready: $OVERLAY_DIR/fontconfig.conf"
echo "scripts/run-e2e-xvfb.sh (npm run test:e2e:xvfb) applies it automatically"
