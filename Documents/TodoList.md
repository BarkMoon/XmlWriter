# XmlWriter 今後の実装項目

## 必須項目

*   テーブル中のパラメータで、名前の頭に#が付いているパラメータについてはXMLに書き出しを行わず、無視するようにします。
*   .xlsxに加えて.xlsmも読めるようにします。
*   Param_A:Table_SubClass(ParamName)[]の表記の実装
    * このParam_Aが含まれるテーブルをTable_Aとしたとき、これはTable_SubClassという名称のテーブルのParamName行を検索して、一致するものがあればその行全体のデータをサブクラスとして埋め込みます。
    * 例えば、Table_AがパラメータId:int, Name.OriginalName, Properties.Abilities:Table_SubClass(Name:string)[]を持ち、Table_SubClassがパラメータId:int, Name:string, FuncName:stringを持つとします。
    このとき、Table_Aの行
    101 | SCTest | Ability_0, Ability_1
    とTable_SubClassの行
    0 | Ability_0 | Func_Test
    1 | Ability_1 | Func_TestEnd
    があるとすると、ここから出力されるXMLは
    <Id>101</Id>
    <Name>
      <OriginalName>SCTest</OriginalName>
    </Name>
    <Properties>
      <Abilities>
        <Id>0</Id>
        <Name>Ability_0</Name>
        <FuncName>Func_Test</FuncName>
      </Abilities>
      <Abilities>
        <Id>1</Id>
        <Name>Ability_1</Name>
        <FuncName>Func_TestEnd</FuncName>
      </Abilities>
    </Properties>
    といった形式になることを想定しています。
    * テスト用にTest_Excel\CardGameElementsTable_MTG.xlsmのCard_FIN_Gテーブルを用意しました。仮テーブルが必要であればそれを使ってください。