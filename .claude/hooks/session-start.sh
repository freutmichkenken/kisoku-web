#!/bin/bash
set -euo pipefail

# Web（クラウド）セッションのみ実行
if [ "${CLAUDE_CODE_REMOTE:-}" != "true" ]; then
  exit 0
fi

# 回帰テストに必要な python-docx（wheels/ と同じ 1.2.0）。導入済みなら何もしない
if ! python3 -c "import docx" 2>/dev/null; then
  pip install --quiet --break-system-packages "python-docx==1.2.0" \
    || pip install --quiet "python-docx==1.2.0"
fi
