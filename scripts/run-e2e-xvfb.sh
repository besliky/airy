#!/usr/bin/env bash
# Run the e2e suite under xvfb (Linux dev machines; CI runs test:e2e inside
# its own xvfb-run). When scripts/setup-e2e-fonts.sh created the pinned-font
# overlay (distro Carlito/Caladea builds differ from the CI ones), apply it
# to this run only via FONTCONFIG_FILE — no system font config is touched.
set -euo pipefail
cd "$(dirname "${BASH_SOURCE[0]}")/.."
if [ -f e2e/.e2e-fonts/fontconfig.conf ]; then
  export FONTCONFIG_FILE="$PWD/e2e/.e2e-fonts/fontconfig.conf"
fi

# A Wayland desktop session leaks WAYLAND_DISPLAY / XDG_SESSION_TYPE /
# GDK_BACKEND into this environment, and playwright's electron.launch
# forwards the whole host env to the app. Electron 43's ozone platform
# auto-detection then picks Wayland, ignores the xvfb DISPLAY and blocks in
# a mojo handshake with the host compositor before main.js ever runs —
# every spec times out (OBS-1665 / OBS-SHELL-HOME-0: ready 94ms with X11
# vs a 180s launch timeout on Wayland). Strip the session variables so the
# app boots on X11 like it does in CI (whose environment is clean), and pin
# the ozone platform to x11 as belt-and-suspenders.
unset WAYLAND_DISPLAY XDG_SESSION_TYPE GDK_BACKEND
export ELECTRON_OZONE_PLATFORM_HINT=x11

# Regression guard: must stay after the unset above. If a future edit
# reintroduces the leak (reordered lines, re-export, wrapper calling this
# one), warn loudly instead of failing the suite with opaque timeouts.
if [ -n "${WAYLAND_DISPLAY:-}" ] || [ -n "${XDG_SESSION_TYPE:-}" ] || [ -n "${GDK_BACKEND:-}" ]; then
  echo "WARNING: Wayland session variables are visible inside run-e2e-xvfb.sh (WAYLAND_DISPLAY='${WAYLAND_DISPLAY:-}' XDG_SESSION_TYPE='${XDG_SESSION_TYPE:-}' GDK_BACKEND='${GDK_BACKEND:-}') — electron may hang on the host compositor instead of xvfb" >&2
fi

exec xvfb-run -a npm run test:e2e "$@"
