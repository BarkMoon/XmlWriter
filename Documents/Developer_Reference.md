# XmlWriter 開発者リファレンス

本ドキュメントでは、`XmlWriter` ツールの内部構造と主要なクラス・メソッドについて解説します。

## プロジェクト構成

*   **Program.cs**: エントリポイント。CLI引数の有無により、GUIモードとCLIモードを振り分けます。
*   **Form1.cs**: GUIモードの実装。ユーザー操作のハンドリングと、GUI用ビジネスロジックを含みます。
*   **Utility/CommandRunner.cs**: CLIモードの実装。GUIに依存しない処理ロジックを提供します。
*   **Utility/IniFile.cs**: 設定保存用のINIファイル操作クラス。

## 主要クラス詳細

### 1. XmlWriter.Program
アプリケーションの開始点です。

*   `Main(string[] args)`
    *   `args` が空の場合: `Application.Run(new Form1())` を呼び出し、GUIを起動します。
    *   `args` がある場合: 引数を解析し、`CommandRunner.Run` を呼び出します。

### 2. XmlWriter.Utility.CommandRunner
CLI実行時の中核ロジックを担当する静的クラスです。

#### メソッド
*   `Run(RunOptions options)`
    *   処理フローの制御を行います。引数は `RunOptions`（`ExcelPath` / `OutputDir` / `TemplateFilePath` / `Mode` / `TargetTableName`）にまとめて渡します。
    *   テンプレートファイルの読み込み、Excelワークブックのオープン(ClosedXML使用)、各ワークシート内の各テーブルごとの生成メソッド呼び出しを、`Mode` に応じて振り分けます。

*   `GenerateXmlFromExcel(XLWorkbook workbook, string tableName, string baseOutputDir, string excelFilePath)`
    *   指定されたテーブルのデータを読み取り、XMLファイルを生成・保存します。
    *   カラムヘッダーから `ColumnInfo` を生成し、ネスト構造(`XElement`)を構築します。

*   `GenerateCSharpFromTemplate(XLWorkbook workbook, string tableName, string baseOutputDir, string template)`
    *   指定されたテンプレート文字列とテーブル構造を使用して、C#ソースコードを生成します（内部で `GenerateCSharpCode` を呼び出します）。

*   `GenerateScriptFromData(XLWorkbook workbook, string tableName, string outputDir, string template)`
    *   行データを `${...}` 変数として展開し、データスクリプト（`<TableName>_Data.cs`）を生成します（`DataCode` モード）。

#### 内部ロジック (テンプレートエンジン)
以下のメソッド群で独自のテンプレートマクロを処理しています。

*   `ProcessTemplateMacros(...)`: `#ForAllSubClasses` ループを処理します。
*   `ProcessInnerPropertyMacros(...)`: `#ForAllSubClassProperties` ループを処理します。
*   `ProcessConditionals(...)`: 式マクロ(`#Eq`等)と条件分岐(`#If`)を一括処理するエントリーポイントです。
*   `ProcessExpressionMacros(...)`: 正規表現で最内の関数マクロ(`#Eq`・`#Replace`等)を評価・置換します。変化が無くなるまで繰り返し適用することでネストを解決します。
*   `ProcessIfBlocks(...)`: `#If` ～ `#Endif` のブロック構造を（ネストを考慮して）解析し、条件に合致するテキストブロックを抽出します。
*   `ProcessDataMacros(...)`: `#ForAllData` ループを処理し、行データの変数(`${Id}`等)置換を行います。
*   `ProcessEraseDuplicatedLines(...)`: `#EraseDuplicatedLine` ～ `#EndErase` ブロック内の重複行（`Trim()` 一致・空行含む）を削除します。

#### その他のユーティリティ
*   `OpenWorkbookReadOnly(path)`: **ファイル全体をメモリ(`MemoryStream`)へ読み込んでから `XLWorkbook` を構築し、ファイルハンドルを保持しません。** これにより XmlWriter で開いている間も Excel 側での編集・保存（元ファイルの置換）が共有違反になりません。読み取りは `FileShare.ReadWrite | Delete` で行います。
*   `TableExists(workbook, tableName)`: 指定テーブルがワークブック内に存在するかを返します（生成前の存在確認に使用）。

### 3. XmlWriter.Form1 (GUI)
Windows Forms画面の実装クラスです。
※ 以前は `CommandRunner` と `Form1` で処理が重複していましたが、リファクタリングにより主要なビジネスロジックは `CommandRunner` に集約されました。`Form1` はUI操作のみを担当し、実際の処理は `CommandRunner` に委譲しています。

#### 主なイベントハンドラ
*   `btnBrowse_Click`: Excelファイルの選択。INIから前回のフォルダパスを復元します。
*   `btnReload_Click`: 開いているExcelのテーブル一覧を再取得します（`LoadTableNames`）。選択中のテーブル名が再取得後も存在すればその選択を維持します。
*   `btnGenerate_Click`: XML生成の実行。
    *   フォルダ選択には標準の `FolderBrowserDialog` ではなく、Vistaスタイル（エクスプローラー形式）のUIを提供するため、`SaveFileDialog` をフォルダ選択モード風に流用する実装を行っています。ユーザーにはダミーファイル名の保存を促すことでフォルダパスを取得します。
*   `btnGenerateClass_Click`: C#クラス生成の実行。
*   `btnGenerateScriptFromData_Click`: データからのスクリプト生成実行。`SaveFileDialog` で出力先フォルダを選択し、`GenerateScriptFromData` を呼び出します。チェックボックス `chkAppendDataSuffix`（既定OFF）がONのときのみ既定ファイル名に `_Data` を付けます。
*   いずれの生成ハンドラも、実行前に `WarnIfTableMissing` で選択テーブルの存在を確認し、Excel側で削除・改名されていた場合は警告を出して中止します（「リロード」で一覧更新を促します）。

## 依存ライブラリ

*   **ClosedXML**: Excelファイルの読み書きに使用。
*   **DocumentFormat.OpenXml**: ClosedXMLの依存関係。

## 注意事項

*   テンプレートエンジンのマクロ解析は正規表現と文字列操作に基づいています。非常に複雑なネストや想定外の記法には対応しきれない場合があります。
*   CLIモードのエラーはコンソール出力(`System.Console`)に行われます。終了コードによる制御が必要な場合は `Environment.ExitCode` の設定を追加実装してください。
