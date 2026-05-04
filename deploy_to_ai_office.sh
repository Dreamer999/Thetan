#!/usr/bin/env bash
# Deploy Eric's AI Organization Morning demo to Dreamer999/Ai_office_DEMO
# and enable GitHub Pages — single command runner.
#
# Usage on your Mac:
#   curl -fsSL https://raw.githubusercontent.com/Dreamer999/Thetan/claude/demo-ai-organization-pKYSH/deploy_to_ai_office.sh | bash

set -euo pipefail

SRC_REPO="Dreamer999/Thetan"
SRC_BRANCH="claude/demo-ai-organization-pKYSH"
DST_REPO="Dreamer999/Ai_office_DEMO"
WORK_DIR="${TMPDIR:-/tmp}/ai_office_deploy_$$"

step() { printf "\n\033[1;33m▶ %s\033[0m\n" "$*"; }
ok()   { printf "\033[1;32m✓ %s\033[0m\n" "$*"; }
die()  { printf "\033[1;31m✗ %s\033[0m\n" "$*" >&2; exit 1; }

command -v git >/dev/null || die "git not found"
command -v gh  >/dev/null || die "gh CLI not found — install: brew install gh && gh auth login"
gh auth status >/dev/null 2>&1 || die "gh not authenticated — run: gh auth login"

step "Cloning $DST_REPO"
mkdir -p "$WORK_DIR"
cd "$WORK_DIR"
git clone "https://github.com/$DST_REPO.git" repo
cd repo

# Make sure we're on main (create if empty repo)
if ! git show-ref --verify --quiet refs/heads/main; then
    git checkout -b main
else
    git checkout main
fi

step "Pulling demo files from $SRC_REPO@$SRC_BRANCH (via gh, supports private repos)"
fetch() {
    local file="$1"
    gh api "repos/$SRC_REPO/contents/$file?ref=$SRC_BRANCH" \
        -H "Accept: application/vnd.github.raw" > "$file"
}
fetch demo.html
fetch index.html
ok "demo.html  ($(wc -c < demo.html  | tr -d ' ') bytes)"
ok "index.html ($(wc -c < index.html | tr -d ' ') bytes)"

step "Committing"
git add demo.html index.html
if git diff --cached --quiet; then
    ok "No changes — already up to date"
else
    git commit -m "Deploy Eric's AI Organization Morning demo

7-scene live demo (cold open + 6 product mockups) showing a single
operator running an AI-powered morning workflow. Source kept in
$SRC_REPO@$SRC_BRANCH."
fi

step "Pushing to $DST_REPO/main"
git push -u origin main

step "Enabling GitHub Pages"
if gh api "repos/$DST_REPO/pages" >/dev/null 2>&1; then
    ok "Pages already enabled — pushing updates is enough"
else
    gh api -X POST "repos/$DST_REPO/pages" \
        -f 'source[branch]=main' -f 'source[path]=/' \
        && ok "Pages enabled" \
        || die "Failed to enable Pages — enable manually: https://github.com/$DST_REPO/settings/pages"
fi

URL="https://${DST_REPO%%/*}.github.io/${DST_REPO##*/}/"
URL_LC="$(echo "$URL" | tr '[:upper:]' '[:lower:]')"

cat <<EOF

────────────────────────────────────────────────────
$(printf "\033[1;32m✓ DONE\033[0m")
Deployed to:   $DST_REPO@main
Live in ~60s:  $URL_LC
Direct demo:   ${URL_LC}demo.html
────────────────────────────────────────────────────

Tip: GitHub Pages URLs are case-insensitive but always lowercase the repo name.
EOF

cd /
rm -rf "$WORK_DIR"
