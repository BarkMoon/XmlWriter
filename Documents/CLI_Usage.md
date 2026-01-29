# XmlWriter CLI (Command Line Interface) 機能

`XmlWriter.exe` はGUIでの操作に加え、コマンドライン引数を指定することで、自動処理（バッチ実行）が可能になりました。
バージョン2.0より、柔軟な引数指定と実行モードの選択が可能になりました。

## 使用方法

### 基本構文

```powershell
XmlWriter.exe <ExcelFilePath> -mode <Mode> [Options...]
```

第1引数にExcelファイル（`.xlsx` または `.xlsm`）のパスを指定し、以降のオプションで動作を制御します。

### モード (`-mode` / `-m`)

以下のいずれかのモードを指定してください。（大文字小文字は区別されません）

| モード名 | 説明 | 必須オプション |
| :--- | :--- | :--- |
| `All` | XML生成とC#コード生成の両方を行います。（デフォルト） | `-output`, `-template` |
| `Xml` | XML生成のみを行います。 | `-output` |
| `Code` | C#コード生成のみを行います。 | `-output`, `-template` |
| `DataCode` | データ行からのスクリプト生成（`GenerateScriptFromData`）を行います。 | `-output`, `-template` |
| `List` | Excelファイル内のテーブル一覧をコンソールに出力します。 | なし |

### オプション引数

| オプション (短縮) | 引数 | 説明 |
| :--- | :--- | :--- |
| `-output` (`-o`) | `<DirPath>` | 出力先ディレクトリのパス。`output/xml` や `output/code` が作成されます。 |
| `-template` (`-t`) | `<FilePath>` | C#コード生成に使用するテンプレートファイルのパス。Code/Allモードで必須。 |
| `-target` (`-table`) | `<TableName>` | 特定のテーブルのみを処理対象にする場合にテーブル名を指定します。省略時は全テーブル対象。 |

## 実行例

### 1. 全テーブルのXMLとコードを生成 (通常使用)

```powershell
XmlWriter.exe "Data.xlsx" -m All -o "OutDir" -t "Template.cs"
```

### 2. XMLのみ生成

```powershell
XmlWriter.exe "Data.xlsx" -m Xml -o "OutDir"
```

### 3. 特定のテーブル ("Card_01") のみコード生成

```powershell
XmlWriter.exe "Data.xlsx" -m Code -o "OutDir" -t "Template.cs" -target "Card_01"
```

### 4. データスクリプトの生成

全データの変数展開を行うスクリプトを生成します。

```powershell
XmlWriter.exe "Data.xlsx" -m DataCode -o "OutDir" -t "Template_Data.cs"
```

### 5. テーブル一覧の確認

```powershell
XmlWriter.exe "Data.xlsx" -m List
```

出力例:
```text
Card_Monster
Card_Item
System_Config
```

## 注意事項

*   出力先ディレクトリ構成 (`xml/` および `code/`) は維持されます。
*   `List` モードは結果を標準出力 (Stdout) に返します。他のツールとパイプで連携可能です。
*   ファイルパス等に空白が含まれる場合は、ダブルクォーテーション `"` で囲んでください。
