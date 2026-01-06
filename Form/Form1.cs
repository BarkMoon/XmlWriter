using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions; // ★追加
using System.Windows.Forms;
using System.Xml.Linq;

namespace XmlWriter
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            cmbSheetName.Enabled = false;
        }

        // ---------------------------------------------------------
        // ★ 変更: 配列フラグ (IsArray) を追加したヘッダー解析クラス
        // ---------------------------------------------------------
        public class ColumnInfo
        {
            public string OriginalHeader { get; set; }
            public string[] PathParts { get; set; }
            public string PropertyName { get; set; }
            public string TypeName { get; set; }
            public bool IsArray { get; set; } // ★ 配列かどうか

            public ColumnInfo(string header)
            {
                OriginalHeader = header;
                string namePart = header;
                TypeName = "string";
                IsArray = false;

                // 型情報の分離 ("Name:int[]" -> "Name", "int", IsArray=true)
                if (header.Contains(":"))
                {
                    var parts = header.Split(':');
                    namePart = parts[0].Trim();
                    if (parts.Length > 1)
                    {
                        string typeRaw = parts[1].Trim().ToLower();
                        // "[]" がついていたら配列扱い
                        if (typeRaw.EndsWith("[]"))
                        {
                            IsArray = true;
                            TypeName = typeRaw.Replace("[]", ""); // "int[]" -> "int"
                        }
                        else
                        {
                            TypeName = typeRaw;
                        }
                    }
                }

                PathParts = namePart.Split('.');
                PropertyName = PathParts.Last();
            }
        }

        // ---------------------------------------------------------
        // ★ 新規: 読み取り専用でExcelを開くヘルパーメソッド
        // ---------------------------------------------------------
        private XLWorkbook OpenWorkbookReadOnly(string path)
        {
            // FileShare.ReadWrite を指定することで、Excelがファイルを開いていても読み込めるようにする
            FileStream fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            return new XLWorkbook(fs);
        }

        // --- イベントハンドラ ---
        private void btnBrowse_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog ofd = new OpenFileDialog())
            {
                ofd.Filter = "Excel Files|*.xlsx";
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    txtFilePath.Text = ofd.FileName;
                    LoadTableNames(ofd.FileName);
                }
            }
        }

        private void LoadTableNames(string filePath)
        {
            cmbSheetName.Items.Clear();
            cmbSheetName.Enabled = false;
            try
            {
                // ★ ヘルパーメソッドを使用
                using (var workbook = OpenWorkbookReadOnly(filePath))
                {
                    foreach (var ws in workbook.Worksheets)
                        foreach (var tbl in ws.Tables) cmbSheetName.Items.Add(tbl.Name);
                }
                if (cmbSheetName.Items.Count > 0) { cmbSheetName.SelectedIndex = 0; cmbSheetName.Enabled = true; }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"読み込みエラー: {ex.Message}\n(パス: {filePath})", "エラー");
            }
        }

        // --- XML生成 ---
        private void btnGenerate_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(txtFilePath.Text) || cmbSheetName.SelectedItem == null) return;
            try
            {
                UpdateStatus("XML生成中...");
                GenerateXmlFromExcel(txtFilePath.Text, cmbSheetName.SelectedItem.ToString());
                UpdateStatus("完了");
                MessageBox.Show("XML生成完了");
            }
            catch (Exception ex) { MessageBox.Show(ex.Message); }
        }

        private void GenerateXmlFromExcel(string filePath, string tableName)
        {
            string outputDir = Path.Combine(Path.GetDirectoryName(filePath), "Output_XML", tableName);
            if (!Directory.Exists(outputDir)) Directory.CreateDirectory(outputDir);

            // ★ ヘルパーメソッドを使用
            using (var workbook = OpenWorkbookReadOnly(filePath))
            {
                var table = workbook.Table(tableName);
                var headers = table.HeadersRow().CellsUsed()
                    .Select(c => new ColumnInfo(c.GetValue<string>()))
                    .ToList();

                foreach (var row in table.DataRange.Rows())
                {
                    XElement rootElement = new XElement("Record");
                    string idValue = null;
                    int cellIndex = 0;

                    foreach (var cell in row.Cells())
                    {
                        if (cellIndex >= headers.Count) break;
                        var colInfo = headers[cellIndex];
                        string rawVal = cell.GetValue<string>();

                        XElement targetParent = rootElement;
                        for (int i = 0; i < colInfo.PathParts.Length - 1; i++)
                        {
                            string partName = colInfo.PathParts[i];
                            XElement existing = targetParent.Element(partName);
                            if (existing == null)
                            {
                                existing = new XElement(partName);
                                targetParent.Add(existing);
                            }
                            targetParent = existing;
                        }

                        // ★ 変更: 配列対応ロジック
                        if (colInfo.IsArray)
                        {
                            // カンマ区切りで分割（前後の空白除去）
                            // 空の場合は要素を作らない
                            if (!string.IsNullOrWhiteSpace(rawVal))
                            {
                                var values = rawVal.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
                                foreach (var val in values)
                                {
                                    targetParent.Add(new XElement(colInfo.PropertyName, val.Trim()));
                                }
                            }
                        }
                        else
                        {
                            // 通常の単一値
                            targetParent.Add(new XElement(colInfo.PropertyName, rawVal));
                        }

                        if (colInfo.PropertyName.Equals("ID", StringComparison.OrdinalIgnoreCase)) idValue = rawVal;
                        cellIndex++;
                    }

                    if (string.IsNullOrEmpty(idValue)) idValue = row.WorksheetRow().RowNumber().ToString();

                    // IDが数値としてパースできるか確認し、できるなら0埋め、できないならそのまま
                    string formattedId = idValue;
                    if (long.TryParse(idValue, out long idNum))
                    {
                        formattedId = idNum.ToString("D6");
                    }

                    rootElement.Save(Path.Combine(outputDir, $"{tableName}_{formattedId}.xml"));
                }
            }
        }

        // --- C#クラス生成 ---
        private void btnGenerateClass_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(txtFilePath.Text) || cmbSheetName.SelectedItem == null) return;

            string tableName = cmbSheetName.SelectedItem.ToString();
            try
            {
                UpdateStatus("クラス生成中...");
                string templateContent = GetTemplateContent(); // 埋め込みリソースから取得

                // ★ ヘルパーメソッドを使用
                using (var workbook = OpenWorkbookReadOnly(txtFilePath.Text))
                {
                    var table = workbook.Table(tableName);
                    var headers = table.HeadersRow().CellsUsed()
                        .Select(c => new ColumnInfo(c.GetValue<string>()))
                        .ToList();

                    string finalCode = GenerateCSharpFromTemplate(tableName, headers, templateContent);

                    using (SaveFileDialog sfd = new SaveFileDialog())
                    {
                        sfd.FileName = $"{tableName}.cs";
                        sfd.Filter = "C# File|*.cs";
                        if (sfd.ShowDialog() == DialogResult.OK)
                        {
                            File.WriteAllText(sfd.FileName, finalCode, Encoding.UTF8);
                            MessageBox.Show("クラスファイルを保存しました。", "成功");
                        }
                    }
                }
                UpdateStatus("完了");
            }
            catch (Exception ex) { MessageBox.Show($"エラー: {ex.Message}"); }
        }

        // 修正版: 外部ファイル (Template.cs) を読み込む
        private string GetTemplateContent()
        {
            // 実行ファイル(.exe)があるフォルダのパスを取得
            string exePath = Application.StartupPath;

            // ファイルパスを結合 (例: C:\...\bin\Debug\net8.0-windows\Template.cs)
            string templatePath = Path.Combine(exePath, "Template.cs");

            if (!File.Exists(templatePath))
            {
                throw new FileNotFoundException(
                    "テンプレートファイルが見つかりません。\n" +
                    $"以下の場所に 'Template.cs' があるか確認してください。\n{templatePath}");
            }

            // ファイルを読み込んで返す
            return File.ReadAllText(templatePath, Encoding.UTF8);
        }

        private string GenerateCSharpFromTemplate(string rootClassName, List<ColumnInfo> columns, string template)
        {
            // 1. ツリー構築
            var rootNode = new ClassNode(rootClassName, rootClassName);
            rootNode.IsRoot = true;

            foreach (var col in columns)
            {
                rootNode.AddPath(col.PathParts, col.TypeName, col.IsArray, rootClassName);
            }

            // 2. マクロブロック (#ForAllSubClasses) の処理
            // ここでループ内部のマクロは処理されます
            string processedTemplate = ProcessTemplateMacros(template, rootNode, rootClassName);

            // 3. ルートクラスのプロパティ生成
            string rootPropertiesCode = BuildPropertiesCodeOnly(rootNode, "        ");

            // 4. グローバル変数の置換
            // ★重要: 変数(@TableNameなど)を実際の値に置き換えます。
            // これにより #Contains(@TableName, ...) が #Contains(CardData_Hanafuda, ...) になります。
            string finalCode = processedTemplate
                .Replace("@TableName", rootClassName)
                .Replace("@GeneratedDate", DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss"))
                .Replace("@RootProperties", rootPropertiesCode.TrimEnd());

            // 5. ★新規追加: 最終的なマクロ一括処理
            // これにより、ループの外側にあった #If や #Replace が解決されます。
            // 変数置換の後に行うことで、置換後の文字列に対して判定・加工が機能します。
            finalCode = ProcessConditionals(finalCode);

            // 6. 改行コード統一 (CR+LF)
            return finalCode.Replace("\r\n", "\n").Replace("\n", "\r\n");
        }

        // 外側のループ (#ForAllSubClasses) を処理
        private string ProcessTemplateMacros(string template, ClassNode rootNode, string rootTableName)
        {
            string startTag = "#ForAllSubClasses";
            string endTag = "#EndForAllSubClasses";

            // マクロブロックを探して処理するループ (複数ブロック対応)
            while (true)
            {
                int startIdx = template.IndexOf(startTag);
                if (startIdx == -1) break;

                int endIdx = template.IndexOf(endTag, startIdx);
                if (endIdx == -1) break; // 閉じタグがない場合は終了

                // ブロックの中身を抽出
                int contentStart = startIdx + startTag.Length;
                int contentLength = endIdx - contentStart;
                string blockTemplate = template.Substring(contentStart, contentLength);

                // 先頭の改行除去
                blockTemplate = TrimStartNewline(blockTemplate);

                var subNodes = rootNode.GetAllNodes().Where(n => !n.IsRoot).ToList();
                StringBuilder loopContent = new StringBuilder();

                foreach (var node in subNodes)
                {
                    // 1. 変数置換
                    string instance = blockTemplate
                        .Replace("@SubClassName", node.ClassName)
                        .Replace("@SubClassTagName", node.XmlTagName)
                        .Replace("@TableName", rootTableName);

                    // 2. 内側のプロパティマクロを処理
                    instance = ProcessInnerPropertyMacros(instance, node);

                    // 3. ★新規: 条件分岐マクロを処理 (#Eq -> #If)
                    instance = ProcessConditionals(instance);

                    loopContent.Append(instance);
                }

                // 終了タグ後ろの改行除去範囲計算
                int removeEndIndex = GetRemoveEndIndex(template, endIdx + endTag.Length);
                int removeLength = removeEndIndex - startIdx;

                // 置換して更新 (次のループで次のブロックを探す)
                template = template.Remove(startIdx, removeLength).Insert(startIdx, loopContent.ToString());
            }

            return template;
        }

        // 内側のループ (#ForAllSubClassProperties) を処理
        private string ProcessInnerPropertyMacros(string template, ClassNode node)
        {
            string startTag = "#ForAllSubClassProperties";
            string endTag = "#EndForAllSubClassProperties";

            while (true)
            {
                int startIdx = template.IndexOf(startTag);
                if (startIdx == -1) break;

                int endIdx = template.IndexOf(endTag, startIdx);
                if (endIdx == -1) break;

                int contentStart = startIdx + startTag.Length;
                int contentLength = endIdx - contentStart;
                string blockTemplate = template.Substring(contentStart, contentLength);

                blockTemplate = TrimStartNewline(blockTemplate);

                StringBuilder loopContent = new StringBuilder();

                // プロパティと子クラスを統合リストとして扱う
                var allProperties = new List<dynamic>();
                foreach (var prop in node.Properties)
                    allProperties.Add(new { Name = prop.Name, Type = ConvertType(prop.TypeName, prop.IsArray) });
                foreach (var child in node.Children)
                    allProperties.Add(new { Name = child.XmlTagName, Type = child.ClassName });

                foreach (var prop in allProperties)
                {
                    // 1. 変数置換
                    string instance = blockTemplate
                        .Replace("@SubClassPropertyName", prop.Name)
                        .Replace("@SubClassPropertyType", prop.Type);

                    // 2. ★新規: 条件分岐マクロを処理 (#Eq -> #If)
                    // 変数置換後の文字列に対して評価を行う
                    instance = ProcessConditionals(instance);

                    loopContent.Append(instance);
                }

                int removeEndIndex = GetRemoveEndIndex(template, endIdx + endTag.Length);
                int removeLength = removeEndIndex - startIdx;

                template = template.Remove(startIdx, removeLength).Insert(startIdx, loopContent.ToString());
            }

            return template;
        }

        // ---------------------------------------------------------
        // 条件分岐処理エンジン (修正版)
        // ---------------------------------------------------------

        private string ProcessConditionals(string text)
        {
            // 1. ★新規: 式マクロ (#Eq, #Not, #And, #Or, #Contains, #Replace) を一括解決
            // ネストに対応するため、ループ処理を行うメソッドを呼び出す
            text = ProcessExpressionMacros(text);

            // 2. #If ... #Endif を構造解析して置換
            return ProcessIfBlocks(text);
        }

        // 式マクロを再帰的に解決するメソッド
        private string ProcessExpressionMacros(string text)
        {
            // 処理対象のマクロ名リスト
            var targetMacros = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                "Eq", "Not", "And", "Or", "Contains", "Replace"
            };

            // 正規表現: #MacroName(Args...) 
            // [^()]* とすることで、「括弧を含まない引数」つまり「最も内側のマクロ」にマッチさせる
            var regex = new Regex(@"#(\w+)\s*\(([^()]*)\)");

            bool changed = true;

            // ネストが解消されるまでループ置換 (例: #Not(#Eq(A,B)) -> #Not(True) -> False)
            while (changed)
            {
                changed = false;
                text = regex.Replace(text, match =>
                {
                    string macroName = match.Groups[1].Value;
                    string argsStr = match.Groups[2].Value;

                    // 対象外のマクロ（#Ifなど）は触らずそのまま返す
                    if (!targetMacros.Contains(macroName))
                    {
                        return match.Value;
                    }

                    // マクロを評価して結果文字列を取得
                    string result = EvaluateMacro(macroName, argsStr);

                    // 結果が変わっていればフラグを立てる
                    if (result != match.Value)
                    {
                        changed = true;
                        return result;
                    }
                    return match.Value;
                });
            }
            return text;
        }

        // 個別のマクロロジックを実行するメソッド
        private string EvaluateMacro(string name, string argsStr)
        {
            // 引数をカンマで分割してトリム
            // ※引数としての文字列内にカンマが含まれるケースは考慮しない簡易実装
            var args = argsStr.Split(',').Select(a => a.Trim()).ToArray();

            switch (name.ToLower()) // 小文字で判定
            {
                // --- 比較・論理 ---
                case "eq":
                    if (args.Length < 2) return "False";
                    return (args[0] == args[1]) ? "True" : "False";

                case "not":
                    if (args.Length < 1) return "False";
                    // TrueならFalse, それ以外ならTrue
                    return (args[0].Equals("True", StringComparison.OrdinalIgnoreCase)) ? "False" : "True";

                case "and":
                    // 引数がすべてTrueならTrue
                    if (args.Length == 0) return "False";
                    bool andResult = args.All(a => a.Equals("True", StringComparison.OrdinalIgnoreCase));
                    return andResult ? "True" : "False";

                case "or":
                    // 引数のどれかがTrueならTrue
                    if (args.Length == 0) return "False";
                    bool orResult = args.Any(a => a.Equals("True", StringComparison.OrdinalIgnoreCase));
                    return orResult ? "True" : "False";

                // --- 文字列操作 ---
                case "contains":
                    if (args.Length < 2) return "False";
                    return args[0].Contains(args[1]) ? "True" : "False";

                case "replace":
                    // #Replace(Source, Old, New)
                    if (args.Length < 3) return args.Length > 0 ? args[0] : "";
                    return args[0].Replace(args[1], args[2]);

                default:
                    // 未知のマクロはそのまま返す（通常ここには来ない）
                    return $"#{name}({argsStr})";
            }
        }

        // #If ブロックの処理 (ネスト対応・構造解析)
        private string ProcessIfBlocks(string text)
        {
            // 再帰的に処理するため、最も外側の #If を探す
            while (true)
            {
                int ifIndex = text.IndexOf("#If");
                if (ifIndex == -1) break; // もう #If はない

                // 対応する #Endif を探す (ネストを考慮)
                int endIndex = FindMatchingEndif(text, ifIndex);
                if (endIndex == -1) break; // 閉じタグが見つからない(異常系)

                // #If(...) ～ #Endif 全体を切り出す
                int length = (endIndex + "#Endif".Length) - ifIndex;

                // 行末の改行まで含めて削除範囲とする
                int removeEndIndex = GetRemoveEndIndex(text, ifIndex + length);
                int fullRemoveLength = removeEndIndex - ifIndex;

                // ブロックの中身を解析して、採用するテキストを決定
                string result = SolveIfBlock(text.Substring(ifIndex, length));

                // 結果の中にさらに #If があるかもしれないので再帰処理しないといけないが、
                // 今回はシンプルに「採用されたテキスト」を元の場所に埋め込んで、
                // whileループの次周で再度検索させることで解決する。
                // (埋め込んだテキスト内に #If があれば次に見つかる)

                text = text.Remove(ifIndex, fullRemoveLength).Insert(ifIndex, result);
            }

            return text;
        }

        // 特定の #If ブロックを解析し、条件に合う部分のテキストを返す
        private string SolveIfBlock(string block)
        {
            // ブロックは "#If(Cond)...#Endif" の形
            // これを行単位などでパースして、If, Elif, Else の区間を見つける
            // ただし、ネストされた #If があると単純な Split はできない。

            // 簡易パーサー: トップレベルの #If, #Elif, #Else を探す
            var segments = ParseIfSegments(block);

            foreach (var seg in segments)
            {
                // Else は無条件で採用
                if (seg.Type == "Else")
                {
                    return TrimStartNewline(seg.Content); // 改行調整
                }

                // If, Elif は条件判定
                if (seg.Condition == "True")
                {
                    return TrimStartNewline(seg.Content); // 改行調整
                }
            }

            return ""; // どの条件にも合致せずElseもない場合
        }

        private class IfSegment
        {
            public string Type; // If, Elif, Else
            public string Condition; // True, False
            public string Content;
        }

        // ブロックをセグメント(If, Elif, Else)に分解する修正版
        private List<IfSegment> ParseIfSegments(string block)
        {
            var list = new List<IfSegment>();

            using (StringReader sr = new StringReader(block))
            {
                string line;
                string currentType = null;
                string currentCond = null;
                StringBuilder currentContent = new StringBuilder();

                // depth: ネストの深さ
                // 0: ブロックの外
                // 1: 現在処理中の #If ブロックの直下 (ここにある #Elif/#Else が有効)
                // 2以上: さらにネストされた #If の内部
                int depth = 0;

                while ((line = sr.ReadLine()) != null)
                {
                    string trimmed = line.Trim();

                    // 1. #If (開始タグ)
                    if (trimmed.StartsWith("#If"))
                    {
                        if (depth == 0)
                        {
                            // トップレベルの開始
                            currentType = "If";
                            currentContent.Clear();
                            currentCond = ExtractCondition(trimmed);
                        }
                        else
                        {
                            // ネストされた #If はコンテンツとして扱う
                            currentContent.AppendLine(line);
                        }
                        depth++;
                        continue;
                    }

                    // 2. #Endif (終了タグ)
                    if (trimmed.StartsWith("#Endif"))
                    {
                        depth--;
                        if (depth == 0)
                        {
                            // トップレベルの終了。現在のセグメントを保存して終了
                            if (currentType != null)
                            {
                                list.Add(new IfSegment { Type = currentType, Condition = currentCond, Content = currentContent.ToString() });
                            }
                            currentType = null;
                        }
                        else
                        {
                            // ネストされた #Endif はコンテンツとして扱う
                            currentContent.AppendLine(line);
                        }
                        continue;
                    }

                    // 3. #Elif / #Else (中間分岐タグ)
                    // depth == 1 のときのみ、分岐として認識する
                    if (depth == 1)
                    {
                        if (trimmed.StartsWith("#Elif"))
                        {
                            // 前のセグメントを保存
                            if (currentType != null)
                            {
                                list.Add(new IfSegment { Type = currentType, Condition = currentCond, Content = currentContent.ToString() });
                            }
                            // 新しいセグメント開始
                            currentType = "Elif";
                            currentContent.Clear();
                            currentCond = ExtractCondition(trimmed);
                            continue;
                        }

                        if (trimmed.StartsWith("#Else"))
                        {
                            // 前のセグメントを保存
                            if (currentType != null)
                            {
                                list.Add(new IfSegment { Type = currentType, Condition = currentCond, Content = currentContent.ToString() });
                            }
                            // 新しいセグメント開始
                            currentType = "Else";
                            currentContent.Clear();
                            currentCond = "True"; // Elseは常にTrue
                            continue;
                        }
                    }

                    // 4. 通常のコンテンツ
                    // セグメントが確定している場合のみ追加
                    if (currentType != null)
                    {
                        currentContent.AppendLine(line);
                    }
                }
            }
            return list;
        }

        private string ExtractCondition(string line)
        {
            // #If(True) -> True
            int start = line.IndexOf('(');
            int end = line.LastIndexOf(')');
            if (start != -1 && end != -1)
            {
                return line.Substring(start + 1, end - start - 1).Trim();
            }
            return "False";
        }

        private int FindMatchingEndif(string text, int startIfIndex)
        {
            int depth = 0;
            int pos = startIfIndex;

            while (pos < text.Length)
            {
                int nextIf = text.IndexOf("#If", pos);
                int nextEnd = text.IndexOf("#Endif", pos);

                if (nextEnd == -1) return -1; // Endifがない

                // 次に近いのが If か Endif か
                if (nextIf != -1 && nextIf < nextEnd)
                {
                    // ネストしたIfが見つかった
                    depth++;
                    pos = nextIf + 3; // 進める
                }
                else
                {
                    // Endifが見つかった
                    depth--;
                    pos = nextEnd + 6; // 進める

                    if (depth == 0)
                    {
                        return nextEnd; // これが対応するEndif
                    }
                }
            }
            return -1;
        }


        // ---------------------------------------------------------
        // ユーティリティ (改行処理)
        // ---------------------------------------------------------

        // 先頭の改行を除去
        private string TrimStartNewline(string text)
        {
            if (string.IsNullOrEmpty(text)) return text;
            if (text.StartsWith("\r\n")) return text.Substring(2);
            if (text.StartsWith("\n")) return text.Substring(1);
            return text;
        }

        // 末尾の改行を含めた削除位置を計算
        private int GetRemoveEndIndex(string fullText, int indexAfterTag)
        {
            int idx = indexAfterTag;
            if (idx < fullText.Length && fullText[idx] == '\r') idx++;
            if (idx < fullText.Length && fullText[idx] == '\n') idx++;
            return idx;
        }

        // パターンA: プロパティ定義のみを生成する (ルートクラスの中身用)
        private string BuildPropertiesCodeOnly(ClassNode node, string indent)
        {
            var allItems = new List<string>();

            // 通常プロパティ
            foreach (var prop in node.Properties)
            {
                string type = ConvertType(prop.TypeName, prop.IsArray);
                string item = $"{indent}[XmlElement(\"{prop.Name}\")]\n" +
                              $"{indent}public {type} {prop.Name} {{ get; set; }}";
                allItems.Add(item);
            }

            // 子クラス(グループ)への参照プロパティ
            foreach (var child in node.Children)
            {
                string item = $"{indent}[XmlElement(\"{child.XmlTagName}\")]\n" +
                              $"{indent}public {child.ClassName} {child.XmlTagName} {{ get; set; }}";
                allItems.Add(item);
            }

            return string.Join("\n\n", allItems);
        }

        // パターンB: クラス定義全体を生成する (サブクラス用)
        private string BuildFullClassCode(ClassNode node, string indent)
        {
            StringBuilder sb = new StringBuilder();

            // サブクラス定義 (partial削除)
            sb.AppendLine($"{indent}public class {node.ClassName}");
            sb.AppendLine($"{indent}{{");

            // 中身は パターンA のロジックを再利用 (インデントを深くして呼ぶ)
            string innerIndent = indent + "    ";
            string properties = BuildPropertiesCodeOnly(node, innerIndent);

            if (!string.IsNullOrEmpty(properties))
            {
                sb.AppendLine(properties);
            }

            sb.Append($"{indent}}}");
            return sb.ToString();
        }

        // ★ 変更: IsArray引数を追加し、List<T> を返すように変更
        private string ConvertType(string typeName, bool isArray)
        {
            string baseType;
            switch (typeName.ToLower())
            {
                case "int": baseType = "int"; break;
                case "long": baseType = "long"; break;
                case "float": baseType = "float"; break;
                case "double": baseType = "double"; break;
                case "bool": baseType = "bool"; break;
                case "date":
                case "datetime": baseType = "DateTime"; break;
                default: baseType = "string"; break;
            }

            return isArray ? $"List<{baseType}>" : baseType;
        }

        private void UpdateStatus(string msg) { lblStatus.Text = msg; Application.DoEvents(); }

        class ClassNode
        {
            public string XmlTagName { get; set; }
            public string ClassName { get; set; }
            public bool IsRoot { get; set; } = false;
            public List<PropertyNode> Properties { get; set; } = new List<PropertyNode>();
            public List<ClassNode> Children { get; set; } = new List<ClassNode>();

            public ClassNode(string xmlTagName, string className) { XmlTagName = xmlTagName; ClassName = className; }

            // ★ 変更: isArray 引数を追加
            public void AddPath(string[] parts, string typeName, bool isArray, string parentPrefix)
            {
                if (parts.Length == 1)
                {
                    // ★ PropertyNodeに配列情報を保存
                    Properties.Add(new PropertyNode { Name = parts[0], TypeName = typeName, IsArray = isArray });
                    return;
                }

                string currentPartName = parts[0];
                var child = Children.FirstOrDefault(c => c.XmlTagName == currentPartName);
                if (child == null)
                {
                    string uniqueClassName = $"{parentPrefix}_{currentPartName}";
                    child = new ClassNode(currentPartName, uniqueClassName);
                    Children.Add(child);
                }
                // 子ノードへ進む (子のプロパティ自体はまだ確定していないので isArray はここでは使わないが、末端まで引き回す)
                child.AddPath(parts.Skip(1).ToArray(), typeName, isArray, child.ClassName);
            }

            public List<ClassNode> GetAllNodes()
            {
                var list = new List<ClassNode> { this };
                foreach (var child in Children) list.AddRange(child.GetAllNodes());
                return list;
            }
        }

        class PropertyNode
        {
            public string Name { get; set; }
            public string TypeName { get; set; }
            public bool IsArray { get; set; } // ★ 追加
        }
    }
}