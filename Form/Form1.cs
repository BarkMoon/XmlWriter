using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml.Linq;

namespace XmlWriter
{
    public partial class Form1 : Form
    {
        // ---------------------------------------------------------
        // テンプレート定義 (外部ファイル化も容易です)
        // ---------------------------------------------------------
        private const string ClassTemplate = @"using System;
using System.Collections.Generic;
using System.Xml.Serialization;

namespace GeneratedClasses
{
    // テーブル: @TableName
    // このファイルは自動生成されています。
    // 手動でロジックを追加したい場合は、別ファイルで partial class を定義してください。

@ClassDefinitions
}
";
        // ---------------------------------------------------------

        public Form1()
        {
            InitializeComponent();
            cmbSheetName.Enabled = false;
        }

        // (ColumnInfoクラス定義などは前回と同じですが、再掲します)
        public class ColumnInfo
        {
            public string OriginalHeader { get; set; }
            public string[] PathParts { get; set; }
            public string PropertyName { get; set; }
            public string TypeName { get; set; }

            public ColumnInfo(string header)
            {
                OriginalHeader = header;
                string namePart = header;
                TypeName = "string";

                if (header.Contains(":"))
                {
                    var parts = header.Split(':');
                    namePart = parts[0].Trim();
                    if (parts.Length > 1) TypeName = parts[1].Trim().ToLower();
                }

                PathParts = namePart.Split('.');
                PropertyName = PathParts.Last();
            }
        }

        // --- イベントハンドラ (btnBrowse_Click, LoadTableNames は変更なし) ---
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
                using (var workbook = new XLWorkbook(filePath))
                {
                    foreach (var ws in workbook.Worksheets)
                        foreach (var tbl in ws.Tables) cmbSheetName.Items.Add(tbl.Name);
                }
                if (cmbSheetName.Items.Count > 0) { cmbSheetName.SelectedIndex = 0; cmbSheetName.Enabled = true; }
            }
            catch { }
        }

        // --- XML生成 (前回の修正版そのまま) ---
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

            using (var workbook = new XLWorkbook(filePath))
            {
                var table = workbook.Table(tableName);
                var headers = table.HeadersRow().CellsUsed().Select(c => new ColumnInfo(c.GetValue<string>())).ToList();

                foreach (var row in table.DataRange.Rows())
                {
                    XElement rootElement = new XElement("Record");
                    string idValue = null;
                    int cellIndex = 0;

                    foreach (var cell in row.Cells())
                    {
                        if (cellIndex >= headers.Count) break;
                        var colInfo = headers[cellIndex];
                        string val = cell.GetValue<string>();

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
                        targetParent.Add(new XElement(colInfo.PropertyName, val));

                        if (colInfo.PropertyName.Equals("ID", StringComparison.OrdinalIgnoreCase)) idValue = val;
                        cellIndex++;
                    }

                    if (string.IsNullOrEmpty(idValue)) idValue = row.WorksheetRow().RowNumber().ToString();
                    long.TryParse(idValue, out long idNum);
                    string formattedId = (idNum != 0) ? idNum.ToString("D6") : idValue;
                    rootElement.Save(Path.Combine(outputDir, $"{tableName}_{formattedId}.xml"));
                }
            }
        }

        // --- ★ 修正版 C#クラス生成処理 ---
        private void btnGenerateClass_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(txtFilePath.Text) || cmbSheetName.SelectedItem == null)
            {
                MessageBox.Show("Excelファイルとテーブルを選択してください。");
                return;
            }

            string tableName = cmbSheetName.SelectedItem.ToString();
            string templatePath = "Template.cs"; // 実行フォルダのTemplate.csを参照

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

                using (var workbook = new XLWorkbook(txtFilePath.Text))
                {
                    var table = workbook.Table(tableName);
                    // ヘッダー解析
                    // (ColumnInfoは以前の定義を使用してください)
                    var headers = table.HeadersRow().CellsUsed()
                        .Select(c => new ColumnInfo(c.GetValue<string>()))
                        .ToList();

                    // クラスコードの生成
                    string finalCode = GenerateCSharpFromTemplate(tableName, headers, templateContent);

                    // 保存
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
            catch (Exception ex)
            {
                MessageBox.Show($"エラー: {ex.Message}");
            }
        }

        // テンプレートに埋め込むロジック
        private string GenerateCSharpFromTemplate(string rootClassName, List<ColumnInfo> columns, string template)
        {
            // 1. ツリー構築
            // ルートノードの名前はテーブル名そのまま、Prefixなし
            var rootNode = new ClassNode(rootClassName, rootClassName);
            rootNode.IsRoot = true;

            foreach (var col in columns)
            {
                // パスを追加 (ここでユニークなクラス名が生成される)
                rootNode.AddPath(col.PathParts, col.TypeName, rootClassName);
            }

            // 2. 定義出力
            StringBuilder definitionsSb = new StringBuilder();
            var allClasses = rootNode.GetAllNodes();

            foreach (var node in allClasses)
            {
                definitionsSb.AppendLine(BuildClassCode(node));
                // クラス間の改行
                definitionsSb.AppendLine();
            }

            // 最後の余分な改行を削除
            string definitionsStr = definitionsSb.ToString().TrimEnd();

            // 3. 置換
            string code = template
                .Replace("@TableName", rootClassName)
                .Replace("@GeneratedDate", DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss"))
                .Replace("@ClassDefinitions", definitionsStr);

            return code;
        }

        private string BuildClassCode(ClassNode node)
        {
            StringBuilder sb = new StringBuilder();
            string indent = "    ";

            // ルート属性
            if (node.IsRoot)
            {
                sb.AppendLine($"{indent}[XmlRoot(\"Record\")]");
            }

            // クラス定義 (ClassNameはユニーク化された名前)
            sb.AppendLine($"{indent}public partial class {node.ClassName}");
            sb.AppendLine($"{indent}{{");

            // プロパティ出力
            // リストを作成して、最後かどうかを判定しやすくする
            var allItems = new List<string>();

            // 1. 通常のプロパティ
            foreach (var prop in node.Properties)
            {
                string type = ConvertType(prop.TypeName);
                string item = $"{indent}    [XmlElement(\"{prop.Name}\")]\n" +
                              $"{indent}    public {type} {prop.Name} {{ get; set; }}";
                allItems.Add(item);
            }

            // 2. 子クラス(グループ)プロパティ
            foreach (var child in node.Children)
            {
                // 型にはユニークな ClassName を使い、プロパティ名には元の XmlTagName (Name) を使う
                string item = $"{indent}    [XmlElement(\"{child.XmlTagName}\")]\n" +
                              $"{indent}    public {child.ClassName} {child.XmlTagName} {{ get; set; }}";
                allItems.Add(item);
            }

            // まとめて結合 (間に空行を入れる)
            if (allItems.Count > 0)
            {
                sb.AppendLine(string.Join("\n\n", allItems));
            }

            // 閉じ括弧 (直前の不要な改行は上記Joinロジックにより排除済)
            sb.Append($"{indent}}}");

            return sb.ToString();
        }

        private string ConvertType(string typeName)
        {
            switch (typeName.ToLower())
            {
                case "int": return "int";
                case "long": return "long";
                case "float": return "float";
                case "double": return "double";
                case "bool": return "bool";
                case "date": case "datetime": return "DateTime";
                default: return "string";
            }
        }

        private void UpdateStatus(string msg) { lblStatus.Text = msg; Application.DoEvents(); }

        // --- ★ 修正版 ClassNode (ユニーク名対応) ---
        class ClassNode
        {
            public string XmlTagName { get; set; } // XMLタグ用 (例: "User")
            public string ClassName { get; set; }  // C#クラス名用 (例: "Table1_User")

            public bool IsRoot { get; set; } = false;

            public List<PropertyNode> Properties { get; set; } = new List<PropertyNode>();
            public List<ClassNode> Children { get; set; } = new List<ClassNode>();

            public ClassNode(string xmlTagName, string className)
            {
                XmlTagName = xmlTagName;
                ClassName = className;
            }

            // パスを追加する際、親のコンテキスト(prefix)を引き継いでユニーク名を生成
            public void AddPath(string[] parts, string typeName, string parentPrefix)
            {
                // 末尾(プロパティ)の場合
                if (parts.Length == 1)
                {
                    Properties.Add(new PropertyNode { Name = parts[0], TypeName = typeName });
                    return;
                }

                // グループ(子クラス)の場合
                string currentPartName = parts[0];

                // 既存の子を探す
                var child = Children.FirstOrDefault(c => c.XmlTagName == currentPartName);
                if (child == null)
                {
                    // ★ ここでユニークな名前を作成
                    // 例: Parent="Table1", Current="User" -> ClassName="Table1_User"
                    string uniqueClassName = $"{parentPrefix}_{currentPartName}";

                    child = new ClassNode(currentPartName, uniqueClassName);
                    Children.Add(child);
                }

                // 再帰呼び出し (次の階層へPrefixを引き継ぐ)
                child.AddPath(parts.Skip(1).ToArray(), typeName, child.ClassName);
            }

            public List<ClassNode> GetAllNodes()
            {
                var list = new List<ClassNode>();
                list.Add(this);
                foreach (var child in Children)
                {
                    list.AddRange(child.GetAllNodes());
                }
                return list;
            }
        }

        class PropertyNode
        {
            public string Name { get; set; }
            public string TypeName { get; set; }
        }
    }
}