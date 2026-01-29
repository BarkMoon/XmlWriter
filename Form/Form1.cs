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
        // INIファイル管理クラス
        private IniFile ini = new IniFile();

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
            public bool IsArray { get; set; }
            public string RefTableName { get; set; }
            public string RefKeyColumn { get; set; }
            public int ColumnNumber { get; set; }

            public ColumnInfo(string header)
            {
                OriginalHeader = header;
                string namePart = header;
                TypeName = "string";
                IsArray = false;
                RefTableName = null;
                RefKeyColumn = null;
                ColumnNumber = 0;

                if (header.Contains(":"))
                {
                    var parts = header.Split(new[] { ':' }, 2);
                    namePart = parts[0].Trim();
                    if (parts.Length > 1)
                    {
                        string typePart = parts[1].Trim();
                        var match = Regex.Match(typePart, @"^([^\(]+)\(([^\)]+)\)(\[\])?$");
                        
                        if (match.Success)
                        {
                            TypeName = match.Groups[1].Value;
                            RefTableName = TypeName;
                            RefKeyColumn = match.Groups[2].Value;
                            if (RefKeyColumn.Contains(":")) RefKeyColumn = RefKeyColumn.Split(':')[0].Trim();
                            IsArray = match.Groups[3].Success;
                        }
                        else
                        {
                            string typeRaw = typePart;
                            if (typeRaw.EndsWith("[]"))
                            {
                                IsArray = true;
                                TypeName = typeRaw.Substring(0, typeRaw.Length - 2);
                            }
                            else
                            {
                                TypeName = typeRaw;
                            }
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
                ofd.Filter = "Excel Files|*.xlsx;*.xlsm";

                // ★追加: INIから前回のフォルダを読み込んでセット
                string lastFolder = ini.Read("LastExcelFolder");
                if (!string.IsNullOrEmpty(lastFolder) && Directory.Exists(lastFolder))
                {
                    ofd.InitialDirectory = lastFolder;
                }

                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    txtFilePath.Text = ofd.FileName;
                    LoadTableNames(ofd.FileName);

                    // ★追加: 選択したフォルダをINIに保存
                    ini.Write("LastExcelFolder", Path.GetDirectoryName(ofd.FileName));
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

            string tableName = cmbSheetName.SelectedItem.ToString();
            string outputDir = null;

            // ★修正: ユーザー要望により SaveFileDialog を使用してフォルダを選択させる
            // (SaveFileDialogならVistaスタイルのダイアログが出るため)
            using (SaveFileDialog sfd = new SaveFileDialog())
            {
                sfd.Title = "XMLデータの出力先フォルダを選択してください (保存ボタンを押してください)";
                sfd.FileName = "SelectFolder"; // ダミーファイル名
                sfd.Filter = "Folder Selection|*.*";
                sfd.CheckFileExists = false;
                sfd.OverwritePrompt = false;

                // INIから前回のフォルダを読み込んでセット
                string lastFolder = ini.Read("LastXmlOutputFolder");
                if (!string.IsNullOrEmpty(lastFolder) && Directory.Exists(lastFolder))
                {
                    sfd.InitialDirectory = lastFolder;
                }

                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    // 選択されたファイルパスのディレクトリ部分を取得
                    outputDir = Path.GetDirectoryName(sfd.FileName);
                    
                    // 選択したフォルダをINIに保存
                    ini.Write("LastXmlOutputFolder", outputDir);
                }
                else
                {
                    return; // キャンセル時は何もしない
                }
            }

            try
            {
                UpdateStatus("XML生成中...");
                GenerateXmlFromExcel(txtFilePath.Text, tableName, outputDir);
                UpdateStatus("完了");
                MessageBox.Show("XML生成完了");
            }
            catch (Exception ex) { MessageBox.Show(ex.Message); }
        }

        private Dictionary<string, List<ColumnInfo>> _headerCache = new Dictionary<string, List<ColumnInfo>>();
        private Dictionary<string, Dictionary<string, IXLRangeRow>> _tableIndexCache = new Dictionary<string, Dictionary<string, IXLRangeRow>>();

        private List<ColumnInfo> GetHeaders(XLWorkbook workbook, string tableName)
        {
            if (_headerCache.TryGetValue(tableName, out var cached)) return cached;

            var table = workbook.Table(tableName);
            var headers = table.HeadersRow().CellsUsed()
                .Where(c => !c.GetValue<string>().TrimStart().StartsWith("#"))
                .Select(c => new ColumnInfo(c.GetValue<string>()) { ColumnNumber = c.WorksheetColumn().ColumnNumber() })
                .ToList();
            
            _headerCache[tableName] = headers;
            return headers;
        }

        private IXLRangeRow FindRow(XLWorkbook workbook, string tableName, string keyColumn, string value)
        {
            string cacheKey = tableName + "::" + keyColumn;
            if (!_tableIndexCache.TryGetValue(cacheKey, out var index))
            {
                index = new Dictionary<string, IXLRangeRow>();
                try {
                    var table = workbook.Table(tableName);
                    var headers = GetHeaders(workbook, tableName);
                    var keyColInfo = headers.FirstOrDefault(h => h.PropertyName.Equals(keyColumn, StringComparison.OrdinalIgnoreCase));
                
                    if (keyColInfo != null)
                    {
                        foreach (var row in table.DataRange.Rows())
                        {
                            try
                            {
                                if (row.IsEmpty()) continue;
                                string val = row.WorksheetRow().Cell(keyColInfo.ColumnNumber).GetValue<string>();
                                if (!index.ContainsKey(val)) index[val] = row;
                            }
                            catch { }
                        }
                    }
                } catch { }
                _tableIndexCache[cacheKey] = index;
            }

            if (index.TryGetValue(value, out var foundRow)) return foundRow;
            return null;
        }

        private XElement CreateElementFromRow(IXLRangeRow row, List<ColumnInfo> headers, XLWorkbook workbook, string elementName)
        {
            XElement rootElement = new XElement(elementName);

            foreach (var colInfo in headers)
            {
                string rawVal = row.WorksheetRow().Cell(colInfo.ColumnNumber).GetValue<string>();
                
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

                if (colInfo.RefTableName != null)
                {
                    if (!string.IsNullOrWhiteSpace(rawVal))
                    {
                        var keys = rawVal.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
                        foreach (var key in keys)
                        {
                            var trimmedKey = key.Trim();
                            var targetRow = FindRow(workbook, colInfo.RefTableName, colInfo.RefKeyColumn, trimmedKey);
                            if (targetRow != null)
                            {
                                var targetHeaders = GetHeaders(workbook, colInfo.RefTableName);
                                var childElement = CreateElementFromRow(targetRow, targetHeaders, workbook, colInfo.PropertyName);
                                targetParent.Add(childElement);
                            }
                        }
                    }
                }
                else if (colInfo.IsArray)
                {
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
                    targetParent.Add(new XElement(colInfo.PropertyName, rawVal));
                }
            }
            return rootElement;
        }

        private void GenerateXmlFromExcel(string filePath, string tableName, string baseOutputDir)
        {
            string outputDir = Path.Combine(baseOutputDir, tableName);
            if (!Directory.Exists(outputDir)) Directory.CreateDirectory(outputDir);

            using (var workbook = OpenWorkbookReadOnly(filePath))
            {
                // Clear Caches
                _headerCache.Clear();
                _tableIndexCache.Clear();

                var headers = GetHeaders(workbook, tableName);

                foreach (var row in workbook.Table(tableName).DataRange.Rows())
                {
                    XElement rootElement = CreateElementFromRow(row, headers, workbook, "Record");

                    // ID Handling
                    string idValue = null;
                    var idCol = headers.FirstOrDefault(h => h.PropertyName.Equals("ID", StringComparison.OrdinalIgnoreCase));
                    if (idCol != null)
                    {
                        idValue = row.WorksheetRow().Cell(idCol.ColumnNumber).GetValue<string>();
                    }
                    if (string.IsNullOrEmpty(idValue)) idValue = row.WorksheetRow().RowNumber().ToString();

                    string formattedId = idValue;
                    if (long.TryParse(idValue, out long idNum))
                    {
                        formattedId = idNum.ToString("D6");
                    }

                    rootElement.Save(Path.Combine(outputDir, $"{tableName}_{formattedId}.xml"));
                }
            }
        }

        // ---------------------------------------------------------
        // ★新規: テンプレート参照ボタンの処理
        // ---------------------------------------------------------
        private void btnBrowseTemplate_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog ofd = new OpenFileDialog())
            {
                ofd.Filter = "C# Template File|*.cs";
                ofd.Title = "テンプレートファイルを選択してください";

                // ★修正: INIから前回のフォルダを読み込んでセット
                string lastFolder = ini.Read("LastTemplateFolder");

                if (!string.IsNullOrEmpty(lastFolder) && Directory.Exists(lastFolder))
                {
                    ofd.InitialDirectory = lastFolder;
                }
                else
                {
                    // 初回は実行ファイルのあるフォルダ
                    ofd.InitialDirectory = Application.StartupPath;
                }

                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    txtTemplatePath.Text = ofd.FileName;

                    // ★修正: 選択したフォルダをINIに保存
                    ini.Write("LastTemplateFolder", Path.GetDirectoryName(ofd.FileName));
                }
            }
        }


        // ---------------------------------------------------------
        // ★変更: C#クラス生成ボタンの処理
        // ---------------------------------------------------------
        private void btnGenerateClass_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(txtFilePath.Text) || cmbSheetName.SelectedItem == null)
            {
                MessageBox.Show("Excelファイルとテーブルを選択してください。");
                return;
            }

            string tableName = cmbSheetName.SelectedItem.ToString();
            string templatePath;

            // ★ テンプレートパスの決定
            if (!string.IsNullOrEmpty(txtTemplatePath.Text))
            {
                // ユーザー指定のテンプレートを使用
                templatePath = txtTemplatePath.Text;
            }
            else
            {
                // 未指定の場合はデフォルト (実行フォルダの Template.cs)
                templatePath = Path.Combine(Application.StartupPath, "Template.cs");
            }

            // ファイル存在確認
            if (!File.Exists(templatePath))
            {
                MessageBox.Show($"テンプレートファイルが見つかりません。\nパス: {Path.GetFullPath(templatePath)}", "エラー");
                return;
            }

            try
            {
                UpdateStatus("クラス生成中...");

                // テンプレート読み込み
                string templateContent = File.ReadAllText(templatePath, Encoding.UTF8);

                using (var workbook = OpenWorkbookReadOnly(txtFilePath.Text))
                {
                    var table = workbook.Table(tableName);
                    var headers = table.HeadersRow().CellsUsed()
                        .Where(c => !c.GetValue<string>().TrimStart().StartsWith("#"))
                        .Select(c => new ColumnInfo(c.GetValue<string>()))
                        .ToList();

                    // クラスコードの生成
                    string finalCode = GenerateCSharpFromTemplate(tableName, headers, templateContent);

                    // 保存
                    using (SaveFileDialog sfd = new SaveFileDialog())
                    {
                        sfd.FileName = $"{tableName}.cs";
                        sfd.Filter = "C# File|*.cs";

                        // ★追加: INIから前回の保存先フォルダを読み込んでセット
                        string lastFolder = ini.Read("LastClassOutputFolder");
                        if (!string.IsNullOrEmpty(lastFolder) && Directory.Exists(lastFolder))
                        {
                            sfd.InitialDirectory = lastFolder;
                        }

                        if (sfd.ShowDialog() == DialogResult.OK)
                        {
                            File.WriteAllText(sfd.FileName, finalCode, Encoding.UTF8);

                            // ★追加: 保存先のフォルダをINIに保存
                            ini.Write("LastClassOutputFolder", Path.GetDirectoryName(sfd.FileName));

                            MessageBox.Show("クラスファイルを保存しました。", "成功");
                        }
                    }
                }
                UpdateStatus("完了");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"エラー: {ex.Message}");
            }
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
                case "string": baseType = "string"; break;
                default: 
                    // 既知の型以外はそのまま（クラス名として）扱う
                    baseType = typeName; 
                    break;
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

        // --- データからスクリプト生成 ---
        private void btnGenerateScriptFromData_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(txtFilePath.Text) || cmbSheetName.SelectedItem == null) return;
            string tableName = cmbSheetName.SelectedItem.ToString();
            
            // 1. テンプレートファイルを選択 (共通ロジック)
            string templatePath;
            if (!string.IsNullOrEmpty(txtTemplatePath.Text))
            {
                templatePath = txtTemplatePath.Text;
            }
            else
            {
                templatePath = Path.Combine(Application.StartupPath, "Template.cs");
            }
            
            if (!File.Exists(templatePath))
            {
                MessageBox.Show($"テンプレートファイルが見つかりません。\nパス: {Path.GetFullPath(templatePath)}", "エラー");
                return;
            }

            // 2. 出力先ファイルを選択 (SaveFileDialogを使用)
            string outputPath = null;
            using (SaveFileDialog sfd = new SaveFileDialog())
            {
                sfd.Title = "保存ファイル名を指定してください";
                sfd.FileName = $"{tableName}_Data.cs";
                sfd.Filter = "C# File|*.cs|All Files|*.*";
                
                string lastFolder = ini.Read("LastScriptOutputFolder");
                if (!string.IsNullOrEmpty(lastFolder) && Directory.Exists(lastFolder))
                {
                    sfd.InitialDirectory = lastFolder;
                }

                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    outputPath = sfd.FileName;
                    ini.Write("LastScriptOutputFolder", Path.GetDirectoryName(outputPath));
                }
                else return;
            }

            try
            {
                UpdateStatus("スクリプト生成中...");
                string templateContent = File.ReadAllText(templatePath, Encoding.UTF8);

                // Helper Method call
                // CommandRunner will handle output dir creation if needed, 
                // but since we pass full path to CommandRunner (which was updated to handle it), this works.
                using (var workbook = OpenWorkbookReadOnly(txtFilePath.Text))
                {
                    XmlWriter.Utility.CommandRunner.GenerateScriptFromData(workbook, tableName, outputPath, templateContent);
                }

                UpdateStatus("完了");
                MessageBox.Show("スクリプト生成完了");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
    }
}