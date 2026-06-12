#!/usr/bin/env bash
# SPDX-License-Identifier: Apache-2.0
#
# check-coi-headers.sh — assert that the mandatory cross-origin isolation
# headers are present in vite.config.ts and the deploy/ Envoy configs.
#
# Used by .github/workflows/ci.yml to guard against accidental removal of
# the COOP/COEP headers (DECISIONS #1/#2). Returns non-zero (fails CI) if
# any required header is missing.
#
# Headers checked:
#   Cross-Origin-Opener-Policy: same-origin
#   Cross-Origin-Embedder-Policy: credentialless
#
# Files checked:
#   vite.config.ts           — dev server + preview block
#   deploy/envoy.yaml        — dev Envoy static-host listener
#   deploy/envoy.prod.yaml   — prod Envoy static-host listener

set -euo pipefail

FAIL=0

check_header() {
  local file="$1"
  local header="$2"
  local value="$3"

  if grep -q "${header}" "${file}" && grep -q "${value}" "${file}"; then
    echo "  OK  ${file}: ${header}: ${value}"
  else
    echo "  FAIL  ${file}: missing '${header}' with value '${value}'"
    FAIL=1
  fi
}

echo "=== COI header guard ==="
echo "Checking COOP/COEP headers in vite.config.ts and deploy/ ..."
echo ""

# ---------------------------------------------------------------------------
# vite.config.ts — must have both COOP and COEP in the coiHeaders object
# ---------------------------------------------------------------------------
echo "-- vite.config.ts --"
check_header "vite.config.ts" "Cross-Origin-Opener-Policy" "same-origin"
check_header "vite.config.ts" "Cross-Origin-Embedder-Policy" "credentialless"

# ---------------------------------------------------------------------------
# deploy/envoy.yaml — dev static-host listener response headers
# ---------------------------------------------------------------------------
if [ -f "deploy/envoy.yaml" ]; then
  echo ""
  echo "-- deploy/envoy.yaml --"
  check_header "deploy/envoy.yaml" "Cross-Origin-Opener-Policy" "same-origin"
  check_header "deploy/envoy.yaml" "Cross-Origin-Embedder-Policy" "credentialless"
else
  echo "  WARN  deploy/envoy.yaml not found (skipping)"
fi

# ---------------------------------------------------------------------------
# deploy/envoy.prod.yaml — prod static-host listener response headers
# ---------------------------------------------------------------------------
if [ -f "deploy/envoy.prod.yaml" ]; then
  echo ""
  echo "-- deploy/envoy.prod.yaml --"
  check_header "deploy/envoy.prod.yaml" "Cross-Origin-Opener-Policy" "same-origin"
  check_header "deploy/envoy.prod.yaml" "Cross-Origin-Embedder-Policy" "credentialless"
else
  echo "  WARN  deploy/envoy.prod.yaml not found (skipping)"
fi

echo ""
if [ "${FAIL}" -eq 0 ]; then
  echo "=== All COI header checks passed ==="
  exit 0
else
  echo "=== COI header guard FAILED ==="
  echo ""
  echo "One or more required cross-origin isolation headers are missing."
  echo "Both of the following must be present in every checked file:"
  echo ""
  echo "  Cross-Origin-Opener-Policy: same-origin"
  echo "  Cross-Origin-Embedder-Policy: credentialless"
  echo ""
  echo "These headers are required for SharedArrayBuffer (DECISIONS #1/#2)."
  echo "See docs/architecture.md and docs/security.md for rationale."
  exit 1
fi
