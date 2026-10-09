#!/usr/bin/env bash
set -uo pipefail
cd "$HOME/Herald-OS"

SHA="f97608f178d1ffeca59860195ab7da295f7c8e5f"
URL="https://codeload.github.com/NousResearch/hermes-agent/tar.gz/$SHA"
TARBALL="upstream/upstream.tar.gz"

echo "==> downloading hermes-agent snapshot ($(date))"
curl -fL --retry 6 --retry-delay 15 --retry-all-errors -o "$TARBALL" "$URL"
echo "==> download done, size:"; ls -la "$TARBALL"

echo "==> extracting"
rm -rf upstream/tmp-extract
mkdir -p upstream/tmp-extract
tar -xzf "$TARBALL" -C upstream/tmp-extract --strip-components=1
echo "==> replacing upstream/hermes-agent"
rm -rf upstream/hermes-agent
mv upstream/tmp-extract upstream/hermes-agent
echo "$SHA" > upstream/hermes-agent/.herald-os-upstream-sha
rm -f "$TARBALL"

echo "==> verify"
ls upstream/hermes-agent/apps/shared/src/ | head -20
echo "DONE"
