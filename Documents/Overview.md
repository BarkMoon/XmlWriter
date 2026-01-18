# XmlWriter 概要仕様書

**XmlWriter** は、Excelファイル(`.xlsx`)で定義されたデータテーブルを読み込み、アプリケーションで使用するための **XMLデータ** と **C#クラスコード** を自動生成するツールです。

GUIモードとコマンドライン(CLI)モードの両方をサポートしています。

## 1. データ定義 (Excel仕様)

Excelの1行目（ヘッダー行）に特別な記法を使用することで、出力されるデータの型や構造を定義します。

### 基本フォーマット
```text
PropertyPath:Type
```

*   **PropertyPath**: ドット(`.`)区切りで階層構造を表現します。
    *   例: `Status.Hp` → `<Status><Hp>...</Hp></Status>`
*   **Type**: データの型を指定します（省略時は `string`）。

### サポートされている型
| 型名 | 説明 |
| :--- | :--- |
| `int` | 整数 |
| `long` | 64ビット整数 |
| `float` | 浮動小数点数 |
| `double` | 倍精度浮動小数点数 |
| `bool` | 真偽値 (TRUE/FALSE) |
| `date` / `datetime` | 日時 |
| `string` | 文字列 (デフォルト) |

### 配列 (リスト)
型名の末尾に `[]` を付けることで、そのカラムを配列として扱います。
Excelのセル内では、値以外の区切り文字（カンマ `,` 等）で複数の値を記述します。

*   例: `Tags:string[]`
*   セル値: `Fire, Magic, Rare` -> `<Tags>Fire</Tags><Tags>Magic</Tags>...`

---

## 2. テンプレートシステム仕様

C#クラス生成時には、指定されたテンプレートファイル(`.cs`)をベースにコード変数が置換・展開されます。

### 変数 (Variables)
| 変数名 | 説明 |
| :--- | :--- |
| `@TableName` | 現在処理中のテーブル名（ルートクラス名） |
| `@GeneratedDate` | 生成日時 |
| `@SubClassName` | ネストされたサブクラス（タグ）の名前 |
| `@SubClassPropertyName` | クラス内のプロパティ名 |
| `@SubClassTagName` | XML上でのタグ名 |

### マクロ制御 (Control Flow)

#### ループ
*   `#ForAllSubClasses` ... `#EndForAllSubClasses`
    *   データ構造に含まれるすべてのクラス（ルート以外）に対してループします。
*   `#ForAllSubClassProperties` ... `#EndForAllSubClassProperties`
    *   クラス内のすべてのプロパティに対してループします。

#### 条件分岐
*   `#If(...)`, `#Elif(...)`, `#Else`, `#Endif`
    *   条件式に基づいてコードの出力を制御します。

### 関数・式 (Expressions)
条件分岐の条件として以下の関数が使用可能です。ネストも可能です。

*   `#Eq(A, B)`: AとBが等しいか
*   `#Not(A)`: Aの否定
*   `#Contains(A, B)`: AがBを含むか
*   `#Or(A, B)`: A または B
*   `#And(A, B)`: A かつ B
*   `#Replace(Text, Old, New)`: 文字列置換（条件式以外でも使用可能）

---

## 3. 出力ファイル

### XMLデータ
*   テーブルの各行が1つのXMLファイルになります。
*   ID列がある場合、ファイル名にIDが付与されます（6桁ゼロ埋め）。
    *   例: `Card_000001.xml`

### C#コード
*   テーブル全体に対応するクラス定義ファイルが生成されます。
*   テンプレートのマクロに従い、必要なプロパティやメソッドが自動実装されます。
