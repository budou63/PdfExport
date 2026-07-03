# PdfExport

`budou63/youtihosyou` から抽出された PDF / Excel エクスポート関連のコードです。

## 含まれるファイル

- `frmPdfExport.txt`
- `modPdfExport.txt`
- `modApi.txt`
- `modNoFillExport.txt`

## Excel出力時の一時ブック方式

任意の一部シートだけを成果物として保存する場合、シート単体の `ws.Copy` ではコピー先にシートモジュールだけが残り、元ブックの標準モジュールや未選択シートが存在しない状態になることがあります。
その状態で `Worksheet_Change`、`Worksheet_Activate`、`Worksheet_Calculate` などのイベントや、UDF を含む再計算が実行・コンパイルされると、未定義の定数・関数、未選択シート参照、マスタシート参照などにより出力エラーになる可能性があります。

現在のExcel出力では、空ブックへセルを貼り直す方式ではなく、元ブック全体を `SaveCopyAs` で一時コピーし、その一時ブックを加工して `xlOpenXMLWorkbook` の `.xlsx` として保存します。これにより、図形・画像・テキストボックス・ページ設定・印刷範囲・列幅・行高・結合セルなど、テンプレートブックの見た目をできるだけ自然に保持します。

出力中は次の Application 状態を保存し、処理中だけ一時変更して、正常終了・エラー終了のどちらでも元の値へ戻します。

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

対象範囲は、元ブック全体の一時コピー作成、一時ブックのオープン、対象シートの表示、数式の値化または対象外シートの非表示、`xlOpenXMLWorkbook` 形式での `SaveAs`、一時ブックの `Close`、一時ファイル削除までです。
`ws.Copy` はExcel出力の主処理では使用しません。また、VBProject / VBComponents を操作してコードを削除する方式も使用しません。

### 対象関数

今回の制御は、次の実処理に実装しています。日本語名ラッパーが英数字名へ委譲している場合は、二重実装を避けるため実処理側のみで制御します。

| 出力種別 | 対象関数 | 変更理由 |
| --- | --- | --- |
| 単一シート・値貼り付け | `Excel_出力実行` | 元ブック全体の一時コピー上で対象シートを値化し、対象外シートを削除するため |
| 複数シート・値貼り付け | `Excel_MultiSheetExport`（`Excel_複数シート出力実行` から委譲） | 対象シートを先に値化してから対象外シートを削除し、参照切れや結合セルエラーを避けるため |
| 単一シート・関数保持 | `Excel_出力実行_関数保持` | 対象シートを表示し、参照元になり得る対象外シートを削除せず `xlSheetVeryHidden` で残すため |
| 複数シート・関数保持 | `Excel_複数シート出力実行_関数保持` | 選択対象外シートを非表示で残し、数式参照を壊しにくくするため |

## 値貼り付けモードと関数保持モード

- 値貼り付けモードでは、対象シートだけを成果物に残します。対象外シートを削除する前に、対象シート上の数式セルだけを `ConvertSheetFormulasToValuesSafely` で値化します。
- 値貼り付けモードの数式値化では、結合セルの左上セルだけを処理し、元シート上で空白に見える数式セルは空白として確定します。これにより、結合セルに対する一括値貼り付けの1004エラーや、空白に見える数式セルが不要に `0` になる問題を避けます。
- 関数保持モードでは、数式は保持します。対象シートは表示状態にし、対象外シートは削除せず `xlSheetVeryHidden` にして残します。そのため、数式自体が未選択シートを参照している場合でも参照切れを起こしにくくなります。
- どちらのモードも最終保存形式は `.xlsx` です。保存形式によりVBAコードは成果物に残らないため、出力ファイルを開いたときの `RYOHI_REASON_CELL` 未定義などのVBAコンパイルエラーを防ぎます。

## 実機確認について

この修正はコード上の静的確認まで実施していますが、この環境では Excel 実機を起動できないため、Excel上での実行確認は未実施です。利用者環境では、以下を確認してください。

1. 通常Excel出力で、結合セル1004エラーが出ないこと。
2. 通常Excel出力で、対象シートだけが成果物に残ること。
3. 通常Excel出力で、数式は値に変換されること。
4. 通常Excel出力で、元シートで空白に見えていた関数セルが不要に `0` にならないこと。
5. 関数保持Excel出力で、対象シートは表示され、対象外シートは `xlSheetVeryHidden` になること。
6. 関数保持Excel出力で、参照元シートを非表示で残すため、数式参照が壊れにくいこと。
7. 図形・画像・テキストボックスが出力ファイルに残ること。
8. 罫線、塗りつぶし、フォント、列幅、行高、結合セル、印刷範囲、ページ設定が維持されること。
9. 出力ファイルは `.xlsx` として保存され、マクロ警告が出ないこと。
10. 出力ファイルを開いても、`RYOHI_REASON_CELL` 未定義などのコンパイルエラーが出ないこと。
11. 元ブック自体は変更されないこと。
12. 一時ファイルが残らないこと。
13. エラー時にも `Application.EnableEvents` などの設定が元に戻ること。

## 変更履歴メモ

- `modPdfExport.txt`: Excel出力4系統を、空ブック再構築方式から `SaveCopyAs` による元ブック全体の一時コピー加工方式へ変更しました。通常出力では対象シートを値化してから対象外シートを削除し、関数保持出力では対象外シートを `xlSheetVeryHidden` で残します。
- `README.md`: 一時ブック方式、値貼り付けモードと関数保持モードの違い、Excel実機での確認手順を更新しました。

## テスト結果

- `rg -n "ExportWorkbookCopyAsXlsx|SaveCopyAs|ConvertSheetFormulasToValuesSafely|xlSheetVeryHidden" modPdfExport.txt` により、一時コピー方式・安全な数式値化・関数保持時のVeryHidden化が実装されていることを確認しました。
- `rg -n "CopyWorksheetContentWithoutCode|ConvertFormulasToSourceDisplayedValues|Workbooks\.Add\(xlWBATWorksheet\)" modPdfExport.txt; test $? -eq 1` により、空ブック再構築方式の主処理が残っていないことを確認しました。
- `rg -n "ws\.Copy" modPdfExport.txt; test $? -eq 1` により、シート単体コピーを使用していないことを確認しました。
- `git diff --check` により、差分に空白エラーがないことを確認しました。
- Excel 実機がないため、上記13項目の実動作確認は未実施です。
