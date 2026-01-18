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
*   `Run(string excelPath, string outputDir, string templateFilePath)`
    *   処理フローの制御を行います。
    *   テンプレートファイルの読み込み、Excelワークブックのオープン(ClosedXML使用)、各テーブルごとの生成メソッド呼び出しを行います。

*   `GenerateXmlFromExcel(XLWorkbook workbook, string tableName, string baseOutputDir, string excelFilePath)`
    *   指定されたテーブルのデータを読み取り、XMLファイルを生成・保存します。
    *   カラムヘッダーから `ColumnInfo` を生成し、ネスト構造(`XElement`)を構築します。

*   `GenerateCSharpFromTemplate(XLWorkbook workbook, string tableName, string baseOutputDir, string template)`
    *   指定されたテンプレート文字列とテーブル構造を使用して、C#ソースコードを生成します。

#### 内部ロジック (テンプレートエンジン)
以下のメソッド群で独自のテンプレートマクロを処理しています。

*   `ProcessTemplateMacros(...)`: `#ForAllSubClasses` ループを処理します。
*   `ProcessInnerPropertyMacros(...)`: `#ForAllSubClassProperties` ループを処理します。
*   `ProcessConditionals(...)`: 条件分岐(`#If`)と式マクロ(`#Eq`等)を一括処理するエントリーポイントです。
*   `ProcessExpressionMacros(...)`: 正規表現を用いて関数マクロ(`#Eq`等)を再帰的に評価・置換します。
*   `ProcessIfBlocks(...)`: `#If` ～ `#Endif` のブロック構造を解析し、条件に合致するテキストブロックを抽出します。
*   `ProcessDataMacros(...)`: `#ForAllData` ループを処理し、行データの変数(`${Id}`等)置換を行います。

### 3. XmlWriter.Form1 (GUI)
Windows Forms画面の実装クラスです。
※ 現状、`CommandRunner` と `Form1` で処理ロジックのコードが重複している部分があります。将来的な保守の際は、共通ロジックを別クラス(`CoreLogic` 等)に切り出すリファクタリングが推奨されます。

#### 主なイベントハンドラ
*   `btnBrowse_Click`: Excelファイルの選択。INIから前回のフォルダパスを復元します。
*   `btnGenerate_Click`: XML生成の実行。
    *   フォルダ選択には標準の `FolderBrowserDialog` ではなく、Vistaスタイル（エクスプローラー形式）のUIを提供するため、`SaveFileDialog` をフォルダ選択モード風に流用する実装を行っています。ユーザーにはダミーファイル名の保存を促すことでフォルダパスを取得します。
*   `btnGenerateClass_Click`: C#クラス生成の実行。
*   `btnGenerateScriptFromData_Click`: データからのスクリプト生成実行。`SaveFileDialog` で出力先フォルダを選択し、`GenerateScriptFromData` を呼び出します。

## 依存ライブラリ

*   **ClosedXML**: Excelファイルの読み書きに使用。
*   **DocumentFormat.OpenXml**: ClosedXMLの依存関係。

## 注意事項

*   テンプレートエンジンのマクロ解析は正規表現と文字列操作に基づいています。非常に複雑なネストや想定外の記法には対応しきれない場合があります。
*   CLIモードのエラーはコンソール出力(`System.Console`)に行われます。終了コードによる制御が必要な場合は `Environment.ExitCode` の設定を追加実装してください。
