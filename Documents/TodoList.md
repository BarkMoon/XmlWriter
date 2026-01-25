# XmlWriter 今後の実装項目

## 必須項目

*   テーブル中のパラメータで、名前の頭に#が付いているパラメータについてはXMLに書き出しを行わず、無視するようにする。
*   .xlsxに加えて.xlsmも読めるようにする。
*   Param_A:SubClass(Name:string)[]の表記の実装
    * このParam_Aが含まれるテーブルをTable_Aとしたとき、これはTable_A_SubClassという名称のテーブルのName:string行を検索して、一致するものがあればその行全体のデータをサブクラスとして埋め込む。