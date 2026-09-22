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
exec xvfb-run -a npm run test:e2e "$@"
