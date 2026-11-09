#!/usr/bin/env bash
# 国企PPT demo 运行包装：自动选择带 python-pptx 的 Python。
# 用法：
#   bash scripts/run_demo.sh
#   PYTHON=/path/to/python3 bash scripts/run_demo.sh --out /tmp/demo.pptx
set -euo pipefail

HERE="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PY="${PYTHON:-python3}"

if ! "$PY" -c "import pptx" >/dev/null 2>&1; then
  # 优先使用 DSH 捆绑运行时（例如本机 ~/.dsh/dsh-runtimes/*/dependencies/python/bin/python3）
  for cand in "${DSH_PYTHON:-}" "$HOME"/.dsh/dsh-runtimes/*/dependencies/python/bin/python3; do
    [ -n "$cand" ] || continue
    [ -x "$cand" ] || continue
    if "$cand" -c "import pptx" >/dev/null 2>&1; then
      PY="$cand"
      break
    fi
  done
fi

if ! "$PY" -c "import pptx" >/dev/null 2>&1; then
  echo "ERROR: 找不到带 python-pptx 的 Python。请 pip install python-pptx，或设置 PYTHON=/path/to/python3。" >&2
  exit 127
fi

echo "using python: $PY"
exec "$PY" "$HERE/build_demo.py" "$@"
