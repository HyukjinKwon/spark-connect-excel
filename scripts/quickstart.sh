#!/usr/bin/env bash
# SPDX-License-Identifier: Apache-2.0
#
# One-command setup for Spark Connect for Excel. Clones (if needed), installs
# deps, sets up the local HTTPS dev cert, and serves the add-in.
#
#   # Default: clone + install + serve the add-in and standalone web demo at
#   # https://localhost:3000 (open /demo/demo.html for an in-browser run)
#   curl -fsSL https://raw.githubusercontent.com/HyukjinKwon/spark-connect-excel/main/scripts/quickstart.sh | bash
#
#   # Windows/Mac desktop: also sideload + open Excel
#   curl -fsSL https://raw.githubusercontent.com/HyukjinKwon/spark-connect-excel/main/scripts/quickstart.sh | bash -s -- desktop
#
# Run it from anywhere; it clones into ./spark-connect-excel (or reuses the repo
# if you are already inside it).
set -euo pipefail

MODE="${1:-web}" # web | desktop
REPO="https://github.com/HyukjinKwon/spark-connect-excel"
DIR="spark-connect-excel"

say() { printf '\033[1;36m[quickstart]\033[0m %s\n' "$*"; }
die() {
  printf '\033[1;31m[quickstart] %s\033[0m\n' "$*" >&2
  exit 1
}

command -v git >/dev/null 2>&1 || die "git is required"
command -v npm >/dev/null 2>&1 || die "Node.js + npm are required (Node 20)"

# 1. Get the source: reuse if we're already in the repo, else clone/update.
if [ -f package.json ] && grep -q '"name": "spark-connect-excel"' package.json 2>/dev/null; then
  say "using the current directory (already the repo)"
elif [ -d "$DIR/.git" ]; then
  say "updating existing ./$DIR"
  git -C "$DIR" pull --ff-only || true
  cd "$DIR"
else
  say "cloning $REPO"
  git clone --depth 1 "$REPO" "$DIR"
  cd "$DIR"
fi

# 2. Dependencies.
say "installing dependencies (npm install)"
npm install

# 2b. Vendor the same-origin runtime assets (Pyodide + the wheels + glue) the
#     engine needs to boot. Idempotent: skipped if already present. This is the
#     exact set the end-to-end CI gate vendors. Needs python3 + pip + curl + tar.
if [ -f public/pyodide/pyodide.mjs ] && ls public/pyspark_client-*.whl >/dev/null 2>&1; then
  say "runtime assets already vendored (public/pyodide + wheels) - skipping"
else
  command -v python3 >/dev/null 2>&1 || die "python3 is required to vendor the runtime (pip builds the wheels)"
  say "vendoring runtime assets (Pyodide + wheels) - first run only, downloads ~30 MB"
  npm run vendor:runtime
fi

# 3. Trusted local HTTPS cert (idempotent; may prompt for keychain/sudo once).
say "installing the local HTTPS dev cert (one time)"
npx --yes office-addin-dev-certs install || say "cert step skipped/failed - rerun 'npx office-addin-dev-certs install' if Excel rejects the cert"

# 4. Serve (and optionally sideload).
if [ "$MODE" = "desktop" ]; then
  say "starting the add-in server in the background and sideloading into Excel"
  (npm run dev:https >/tmp/scx-dev.log 2>&1 &)
  npx --yes office-addin-debugging start manifest.xml desktop --app excel
else
  say "serving the add-in and standalone web demo at https://localhost:3000"
  say "Next: open https://localhost:3000/demo/demo.html for an in-browser run,"
  say "      or re-run with 'desktop' to sideload into Excel desktop."
  npm run dev:https
fi
