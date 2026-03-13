#!/usr/bin/env bash
# Auto-create or update a PR whenever Claude Code finishes a session with new commits.
set -euo pipefail

cd "$(git -C "$(dirname "$0")" rev-parse --show-toplevel)"

BRANCH=$(git rev-parse --abbrev-ref HEAD)
BASE="main"

# Do nothing on main / master
if [[ "$BRANCH" == "main" || "$BRANCH" == "master" ]]; then
  exit 0
fi

# Ensure the branch is pushed
git push -u origin "$BRANCH" --quiet 2>/dev/null || true

# Count commits ahead of base (try remote first, fall back to local)
COMMITS=$(git rev-list --count "origin/$BASE..$BRANCH" 2>/dev/null \
          || git rev-list --count "$BASE..$BRANCH" 2>/dev/null \
          || echo 0)

if [[ "$COMMITS" == "0" ]]; then
  echo "[auto-pr] No new commits on $BRANCH — skipping PR."
  exit 0
fi

# Check if a PR already exists for this branch
EXISTING=$(gh pr list --head "$BRANCH" --json number,url \
           --jq '.[0] | "\(.number) \(.url)"' 2>/dev/null || true)

if [[ -n "$EXISTING" ]]; then
  echo "[auto-pr] PR already exists: $EXISTING"
  exit 0
fi

# Collect commit messages for the PR body
LOG=$(git log "origin/$BASE..$BRANCH" --oneline 2>/dev/null \
      || git log "$BASE..$BRANCH" --oneline)

TITLE="[Auto PR] $(git log -1 --pretty=%s)"
BODY="$(cat <<EOF
## 自动生成的 PR

**分支:** \`$BRANCH\`
**新增提交:** $COMMITS 个

### 提交记录
\`\`\`
$LOG
\`\`\`

---
*此 PR 由 Claude Code Stop Hook 自动创建。*
EOF
)"

gh pr create \
  --title "$TITLE" \
  --body "$BODY" \
  --head "$BRANCH" \
  --base "$BASE" \
  && echo "[auto-pr] PR created for branch: $BRANCH" \
  || echo "[auto-pr] Warning: gh pr create failed (may need gh auth login)"
