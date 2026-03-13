#!/usr/bin/env bash
# Auto-push current branch after every Claude Code session.
set -euo pipefail

cd "$(git -C "$(dirname "$0")" rev-parse --show-toplevel)"

BRANCH=$(git rev-parse --abbrev-ref HEAD)

# Skip detached HEAD
if [[ "$BRANCH" == "HEAD" ]]; then
  exit 0
fi

# Nothing to push if working tree is clean and branch is up to date
UNPUSHED=$(git rev-list --count "origin/$BRANCH..HEAD" 2>/dev/null || echo 1)
if [[ "$UNPUSHED" == "0" ]]; then
  echo "[auto-push] Already up to date."
  exit 0
fi

git push -u origin "$BRANCH"
echo "[auto-push] Pushed $UNPUSHED commit(s) on '$BRANCH'."
