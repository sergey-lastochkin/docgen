#!/usr/bin/env bash
# make.sh — комплект документов из описания заказа одной фразой.
#   ./make.sh "ООО Ромашка, настройка обмена УТ с Ozon, 40 тысяч, с 12 по 20 октября"
#   ./make.sh examples/romashka-ozon.json     # собрать по готовой спецификации
# Спецификацию по образцу examples/romashka-ozon.json составляет Claude Code (claude -p).
# Неизвестные реквизиты клиента остаются «______» — допишите их в specs/*.json и пересоберите.
set -euo pipefail
cd "$(dirname "$0")"
PY=${PYTHON:-python3}
$PY -c 'import docx' 2>/dev/null || { echo "нужен python-docx: pip install -r requirements.txt"; exit 1; }
[ -f executor.json ] || { echo "нет executor.json — скопируйте executor.example.json и впишите свои реквизиты"; exit 1; }

if [ -f "${1:-}" ]; then spec="$1"; else
  desc="${1:?опишите заказ в кавычках}"
  command -v claude >/dev/null || { echo "нужен Claude Code CLI (claude) или готовая спецификация JSON"; exit 1; }
  mkdir -p specs; spec="specs/$(date +%Y%m%d-%H%M%S).json"
  prompt="Составь спецификацию заказа для генератора документов. Образец формата и стиля — файл examples/romashka-ozon.json (прочитай его). Описание нового заказа: «$desc».
Правила: верни ТОЛЬКО JSON того же формата, без пояснений. number — номер в формате ДД-ММ/ГГ по дате договора (сегодня $(date +%F), если не сказано иное); date/start/end — ГГГГ-ММ-ДД. Сумма items равна цене из описания. Неизвестные реквизиты клиента — «______». Работы, исходные данные, ТЗ, критерии приёмки и «не входит» — конкретно под этот заказ, деловым стилем образца."
  out=$(claude -p "$prompt" --permission-mode bypassPermissions)
  echo "$out" | $PY -c 'import sys,json; s=sys.stdin.read(); j=json.loads(s[s.index("{"):s.rindex("}")+1]); print(json.dumps(j,ensure_ascii=False,indent=2))' > "$spec" \
    || { echo "не удалось получить спецификацию:"; echo "$out" | tail -20; exit 1; }
  echo "спецификация: $spec"
fi
$PY docgen.py "$spec"
grep -q '______' "$spec" && echo "ВНИМАНИЕ: в спецификации пропуски «______» — допишите реквизиты в $spec и запустите ./make.sh $spec" || true
