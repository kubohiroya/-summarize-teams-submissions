# summarize-teams-submissions

1. Microsoft Teamsの課題として学生が提出済みのExcelファイル群を、SharePointから上位フォルダごとダウンロードします。zipを展開したときに 「./res/提出済みのファイル/学生氏名/課題名/バージョン1 /提出ファイル名」 となるような形で配置します。
2. 評価対象とする学生名簿のCSVファイル(タブ区切り)を用意し、「./res/member.csv」として配置します。
3. src/index.ts を開き、次の内容を定義します。
   * 定数ASSIGNMENT_TITLESによって、抽出する課題名を文字列の配列で定義します。
   * 定数HEADER_COLUMNSによって、表のヘッダを文字列の配列で定義します。
   * 定数SUBMISSION_VERSION_TAGによって、抽出するバージョン番号を定義します。
   * 関数getSrcItem内のコードによって、学生が提出したExcelファイル1つから、最終的にExcelファイルとして出力する内容の1行分を、文字列または数値または真偽値またはnull値の配列として返す内容を定義します。
   
5. 新規に作成対象とするExcelファイルの名称を、たとえば「output.xlsb」とし、その中に作成するシート名を、たとえば「2021」のように決めて、以下コマンドを実行します。
```bash
ts-node index.ts ./res/member.csv ./res/提出済みのファイル output.xlsb 2020
```
4. 処理結果が、output.xlsxに出力されます。
