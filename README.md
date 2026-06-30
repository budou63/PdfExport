# PdfExport

`budou63/youtihosyou` から抽出された PDF / Excel エクスポート関連のコードです。

## 含まれるファイル

- `frmPdfExport.txt`
- `modPdfExport.txt`
- `modApi.txt`
- `modNoFillExport.txt`

## Excel出力時のイベント・再計算制御

任意の一部シートだけを新規ブックへコピーして `.xlsx` として保存する場合、コピー先には選択されたシートのシートモジュールだけが残り、元ブックの標準モジュールや未選択シートが存在しないことがあります。
その状態で `Worksheet_Change`、`Worksheet_Activate`、`Worksheet_Calculate` などのイベントや、UDF を含む再計算が実行されると、未定義の定数・関数、未選択シート参照、マスタシート参照などにより出力エラーになる可能性があります。

このリポジトリでは、Excel出力の実処理側で次の Application 状態を保存し、出力中だけ一時変更して、正常終了・エラー終了のどちらでも元の値へ戻します。

- `Application.EnableEvents`
- `Application.ScreenUpdating`
- `Application.DisplayAlerts`
- `Application.Calculation`

出力中は以下の状態にします。

```vba
Application.EnableEvents = False
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Application.Calculation = xlCalculationManual
```

対象範囲は、空の新規ブック作成、最初のシート転記、2枚目以降のシート転記、コピー先での値貼り付け、外部参照の正規化、シート名変更等の出力後処理、`xlOpenXMLWorkbook` 形式での `SaveAs`、一時ブックの `Close` までです。
保存・終了が完了する前に `EnableEvents` を戻さないことに加えて、Excel出力では `ws.Copy` を使わず、空の新規ブックへセル内容・書式・印刷設定を転記します。これにより、元シートのシートモジュール（例: `Worksheet_Change`）をコピー先ブックへ持ち込まず、`RYOHI_REASON_CELL` のような元ブック側の標準モジュール定数に依存するコードが `.xlsx` 保存時にコンパイルされることを防ぎます。

### 対象関数

今回のイベント・再計算制御は、次の実処理に実装しています。日本語名ラッパーが英数字名へ委譲している場合は、二重実装を避けるため実処理側のみで制御します。

| 出力種別 | 対象関数 | 変更理由 |
| --- | --- | --- |
| 単一シート・値貼り付け | `Excel_出力実行` | `valueRange.Value = valueRange.Value` を含むコピー先セル変更中に `Worksheet_Change` を発火させないため |
| 複数シート・値貼り付け | `Excel_MultiSheetExport`（`Excel_複数シート出力実行` から委譲） | 任意の複数シートコピーと値貼り付け中に、未選択シートや標準モジュールへ依存するイベントを発火させないため |
| 単一シート・関数保持 | `Excel_出力実行_関数保持` | シートコピー、外部参照正規化、保存・終了までイベントと自動再計算を止めるため |
| 複数シート・関数保持 | `Excel_複数シート出力実行_関数保持` | 任意の複数シートコピー、外部参照正規化、保存・終了までイベントと自動再計算を止めるため |

## 関数保持モードの注意点

- VBAイベントやシートモジュール内マクロが、未出力のシート、元ブックの標準モジュールの `Public Const`、共通関数などへ依存している場合でも、出力中はイベントを停止し、Excel出力先へシートモジュールをコピーしないことで出力エラーを防止します。
- 関数保持モードでは、数式自体が未選択シートを参照している場合、参照切れや外部参照が残る可能性があります。今回の修正対象は、コピーされたシートモジュール内のVBAイベントやマクロ実行に起因する出力エラーの防止です。
- 値貼り付けモードでは、出力時点の計算結果を値として保存します。

## 実機確認について

この修正はコード上の静的確認まで実施していますが、この環境では Excel 実機を起動できないため、Excel上での実行確認は未実施です。利用者環境では、以下を確認してください。

1. 4シート中2シートを出力できること。
2. 4シート中3シートを出力できること。
3. 5シート中1シートを出力できること。
4. 選択シートに `Worksheet_Change` があっても出力できること。
5. `Worksheet_Change` が未選択のマスタシートを参照していても出力できること。
6. `Worksheet_Change` が標準モジュールの `Public Const` を参照していても出力できること。
7. `Worksheet_Change` が標準モジュールの `Public Function` を呼んでいても出力できること。
8. 未選択シートや標準モジュールがコピー先に存在しなくても出力できること。
9. 値貼り付け単一出力が成功すること。
10. 値貼り付け複数出力が成功すること。
11. 関数保持単一出力が成功すること。
12. 関数保持複数出力が成功すること。
13. 出力された `.xlsx` を開いてもVBAコンパイルエラーが表示されないこと。
14. 出力後、元ブックのイベントが正常に動作すること。
15. 出力中にエラーが起きても `Application.EnableEvents` が元の値へ戻ること。
16. 元々 `Application.EnableEvents = False` の場合、終了後も `False` のままであること。
17. `Application.Calculation`、`Application.ScreenUpdating`、`Application.DisplayAlerts` も元の値へ戻ること。
18. エラー後に一時的な `Book1` 等が残らないこと。

## 変更履歴メモ

- `modPdfExport.txt`: Excel出力4系統で Application 状態を保存・一時変更・復元する共通の終了処理を追加しました。さらにExcel出力では `ws.Copy` を使わず、空の新規ブックへ表示内容・書式・印刷設定だけを転記することで、コピー先 `.xlsx` にシートモジュールを持ち込まないようにしました。複数シート出力ではエラー復元後に呼び出し側へ再送出します。
- `README.md`: イベント停止の目的、対象関数、関数保持モードの注意点、Excel実機での確認手順を追記しました。

## テスト結果

- `rg -n "Application\.(EnableEvents|ScreenUpdating|DisplayAlerts|Calculation)|CleanUp:|Err\.Raise" modPdfExport.txt` により、対象処理の状態保存・停止・復元・再送出箇所を確認しました。
- `rg -n "valueRange\.Value = valueRange\.Value" modPdfExport.txt` により、値貼り付け処理が維持されていることを確認しました。
- `rg -n "ws\.Copy|CopyWorksheetContentWithoutCode|Workbooks\.Add\(xlWBATWorksheet\)" modPdfExport.txt` により、Excel出力がシートモジュールをコピーする `ws.Copy` ではなく、マクロなし転記用ヘルパーを使用していることを確認しました。
- `git diff --check` により、差分に空白エラーがないことを確認しました。
- Excel 実機がないため、上記18項目の実動作確認は未実施です。
