# C/O チェックリスト

施設×氏名ごとの C / O タップ漏れを確認するスマホ用チェックリスト。

- 公開URL: https://kahoo-de3.github.io/cl-9d037c71/ （noindex・robots.txtで検索避け）
- 施設名・氏名データはPINから導出した鍵でAES-GCM暗号化して埋め込み（PINはリポジトリにもページにも保存されない）。初回にPIN入力、以降は同じ端末なら自動で解錠
- チェック状態は端末のブラウザ（localStorage）に年月ごとに自動保存

## リストの更新（Excel差し替え）

```
pip install openpyxl cryptography
python tools/generate.py <新しいBook.xlsx> <PIN>
git add docs/index.html && git commit -m "update list" && git push
```

PINを変えたい場合も同じコマンドで再生成する（全端末で再入力が必要になる）。

- Excelのレイアウト前提: B列=施設名, C列=氏名, D列=C, E列=O（5行目以降）
- 外字「祐」(U+E682) は自動で通常の「祐」に置換される
