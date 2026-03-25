## Sync_Schedule

このシステムは、GoogleドライブにあるExcelファイルから自分のGoogleカレンダーに予定を同期するシステムです。
基本的に、一度スケジューラを設定すれば自動で動き続け、全てGoogleドライブ上で完結するためアプリのインストールなどは必要ありません。

### 免責
完全自己責任でご使用ください。

### 動作環境
* Google App Script

### 使い方
#### Schedule_base.xlsx

1. シート名は"yyyyMM"形式で書いてください。（例：2026年1月の場合、202601）

2. 1行目は日付用です。

3. A列には苗字や名前、あだ名など、分かるものを入力してください。
   ![RM_Name](img\RM_Name.png)

4. 名前の行にそれぞれ案件名や、現場名などわかるものを記載してください。
   ![RM_PJName](img\RM_PJName.png)

> [!Tip]
> 拡張子を変更しなければファイル名は変更しても問題ありません。


#### convert_Spreadsheet_this_file.xlsx

1. Googleドライブにスプレッドシートとして保存してください。
   ![RM_Spreadsheets](img\RM_Spreadsheets.png)

2. Apps Scriptのプロジェクトの設定にある「スクリプトプロパティを編集」からConfigシートのConfigNameと自分の環境の値を入力してください。
   ![RM_Spreadsheets](img\RM_Config.png)
3. こんな感じになります
   ![RM_Property.png](img\RM_Property.png)

> [!IMPORTANT]
   > バグによりConfigシートから値をとるようになっているためConfigシートにもURLなどの設定を記載しておいてください。

### 実装予定機能（予定は未定😏）
* サイボウズOfficeの予定登録に対応したCSVファイル出力
