using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
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
            var rootNode = new ClassNode(rootClassName, rootClassName);
            rootNode.IsRoot = true;

            foreach (var col in columns)
            {
                // ★ 配列情報を渡す
                rootNode.AddPath(col.PathParts, col.TypeName, col.IsArray, rootClassName);
            }

            StringBuilder definitionsSb = new StringBuilder();
            var allClasses = rootNode.GetAllNodes();
            foreach (var node in allClasses)
            {
                definitionsSb.AppendLine(BuildClassCode(node));
                definitionsSb.AppendLine();
            }

            return template
                .Replace("@TableName", rootClassName)
                .Replace("@GeneratedDate", DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss"))
                .Replace("@ClassDefinitions", definitionsSb.ToString().TrimEnd());
        }

        private string BuildClassCode(ClassNode node)
        {
            StringBuilder sb = new StringBuilder();
            string indent = "    ";

            if (node.IsRoot) sb.AppendLine($"{indent}[XmlRoot(\"Record\")]");

            sb.AppendLine($"{indent}public partial class {node.ClassName}");
            sb.AppendLine($"{indent}{{");

            var allItems = new List<string>();
            foreach (var prop in node.Properties)
            {
                // ★ 配列かどうかに応じて型を変換
                string type = ConvertType(prop.TypeName, prop.IsArray);
                string item = $"{indent}    [XmlElement(\"{prop.Name}\")]\n" +
                              $"{indent}    public {type} {prop.Name} {{ get; set; }}";
                allItems.Add(item);
            }
            foreach (var child in node.Children)
            {
                string item = $"{indent}    [XmlElement(\"{child.XmlTagName}\")]\n" +
                              $"{indent}    public {child.ClassName} {child.XmlTagName} {{ get; set; }}";
                allItems.Add(item);
            }

            if (allItems.Count > 0) sb.AppendLine(string.Join("\n\n", allItems));
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