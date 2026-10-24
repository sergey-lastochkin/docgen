# docgen

Комплект документов по заказу **из одной строки**: смета, договор возмездного оказания услуг, техническое задание (приложение № 1) и акт — в `.docx`, готовые к подписи.

```bash
./make.sh "ООО Ромашка, настройка обмена УТ с Ozon, 40 тысяч, с 12 по 20 октября"
```

1. Claude Code (`claude -p`) по описанию составляет спецификацию заказа в JSON — по образцу `examples/romashka-ozon.json`: работы и цены, сроки, исходные данные, критерии приёмки, что не входит.
2. `docgen.py` проверяет спецификацию и собирает четыре документа на `python-docx`: Times New Roman 12, нумерация пунктов, реквизиты сторон, сумма прописью, подписи.
3. Неизвестные реквизиты клиента остаются «______» — дописываются в JSON, и комплект пересобирается за секунду.

| Смета | Договор |
|---|---|
| ![Смета](docs/estimate.png) | ![Договор](docs/contract.png) |

## Что внутри договора
- предоплата и остаток после акта; **молчаливая приёмка**: если заказчик не подписал акт и не прислал замечаний за 3 рабочих дня, услуги считаются принятыми;
- срок привязан к получению доступов и исходных данных, перенос — только письменно;
- работа на тестовых данных, КЭП и закрытые ключи остаются у заказчика;
- необязательная пеня с потолком, факсимиле подписи (`signature.png`) по флагу `facsimile`.

## Запуск
```bash
pip install -r requirements.txt
cp executor.example.json executor.json        # свои реквизиты, в git не попадает
python3 docgen.py examples/romashka-ozon.json  # → out/12-10-26/*.docx
python3 docgen.py examples/romashka-ozon.json --check   # только проверить спецификацию
./make.sh "описание заказа"                    # спецификация через Claude Code + документы
```

## Тесты
```bash
python3 -m unittest discover -s tests -t .
```
Сумма прописью (род, падежи тысяч и миллионов), итог по работам, проверка дат и обязательных полей.

---
**English.** Generates a ready-to-sign set of Russian service-contract documents (estimate, contract, statement of work, acceptance act) from a JSON spec; the spec itself can be drafted from a one-line order description by Claude Code. Python + python-docx.

Автор: Сергей Гончаров — 1С, интеграции и автоматизация. [lastochkin1c.ru](https://lastochkin1c.ru) · Telegram [t.me/metaanswer](https://t.me/metaanswer)
