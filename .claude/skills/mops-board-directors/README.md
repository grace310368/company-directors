# mops-board-directors

Claude Code skill：從公開資訊觀測站的重大訊息推導公司現任董事名單，並產出 Excel。

使用方式：在這個 repo 開 Claude Code，直接說「幫我查以下公司的最新董事名單：彰銀、第一銀、……」，或輸入 `/mops-board-directors`。

也可以單獨執行腳本：

```bash
python3 scripts/mops.py lookup 彰化商業銀行 第一商業銀行
python3 scripts/mops.py news 2892 --from 112/01 --to 115/10 -o work/news_2892.json
python3 scripts/mops.py filter work/news_2892.json --subsidiary 第一銀行
python3 scripts/build_excel.py examples/sample_input.json -o 範例.xlsx   # 需要 openpyxl
```

觀測站會封鎖部分雲端 IP。抓取失敗時，請改在自己的電腦上執行 `mops.py`。
