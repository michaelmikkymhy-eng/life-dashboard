#!/bin/sh
# ===== 铁律守卫（2026-11-22 用户立规）=====
# Excel 台账(xlsx) 与 Markdown 日记(md) 必须逐项一致，不一致则拒绝提交。
# 覆盖项：热量/蛋白/脂肪/碳水、消费合计、准提咒、睡眠时长、体重。
# 手动在 Obsidian 里改日记导致失配时，此钩子会拦住 commit。
set -e
ROOT=$(git rev-parse --show-toplevel)
cd "$ROOT" || exit 1

changed=$(git -c core.quotePath=false diff --cached --name-only --diff-filter=ACM \
  | grep -E '^生活档案/日记/[0-9]{4}-[0-9]{2}-[0-9]{2}\.md$' \
  | sed -E 's#.*/##; s#\.md$##' || true)

[ -z "$changed" ] && exit 0

fail=0
for d in $changed; do
  out=$(cd 生活档案 && PYTHONIOENCODING=utf-8 python3 sync.py check --date "$d" 2>&1) || true
  case "$out" in
    *"校验通过"*) printf '  ✅ %s 一致\n' "$d" ;;
    *)
      printf '  ❌ %s 不一致：\n' "$d"
      printf '%s\n' "$out" | tail -5 | sed 's/^/     /'
      fail=1
      ;;
  esac
done

if [ "$fail" -eq 1 ]; then
  echo ""
  echo "❌ 铁律拦截：Markdown 与 Excel 台账不一致，提交已中止。"
  echo "   处理：让 AI 重新同步该日数据（sync.py 对应命令），或撤销日记里的手改。"
  exit 1
fi

exit 0
