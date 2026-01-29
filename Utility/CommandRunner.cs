using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using ClosedXML.Excel;
using System.Xml.Linq;

namespace XmlWriter.Utility
{
    public static class CommandRunner
    {
        public enum ExecutionMode
        {
            All,
            Xml,
            Code,
            List,
            DataCode
        }

        public class RunOptions
        {
            public string ExcelPath { get; set; }
            public string OutputDir { get; set; }
            public string TemplateFilePath { get; set; }
            public ExecutionMode Mode { get; set; } = ExecutionMode.All;
            public string TargetTableName { get; set; } // Optional: null or empty means all
        }

        public static void Run(RunOptions options)
        {
            SetupLogger();
            Log($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] 実行開始");
            Log($"モード: {options.Mode}");
            Log($"Excelパス: {options.ExcelPath}");
            if (!string.IsNullOrEmpty(options.TargetTableName))
            {
                Log($"対象テーブル: {options.TargetTableName}");
            }

            if (!File.Exists(options.ExcelPath))
            {
                throw new FileNotFoundException("Excelファイルが見つかりません", options.ExcelPath);
            }

            // 必須パスのチェック
            if (options.Mode == ExecutionMode.Code || options.Mode == ExecutionMode.All || options.Mode == ExecutionMode.DataCode)
            {
                if (string.IsNullOrEmpty(options.TemplateFilePath) || !File.Exists(options.TemplateFilePath))
                {
                    throw new FileNotFoundException("テンプレートファイルが見つかりません（コード生成には必須です）", options.TemplateFilePath);
                }
            }

            using (var workbook = OpenWorkbookReadOnly(options.ExcelPath))
            {
                // キャッシュクリア
                _headerCache.Clear();
                _tableIndexCache.Clear();

                // リストモード
                if (options.Mode == ExecutionMode.List)
                {
                    ListTables(workbook);
                    return;
                }

                Log($"出力ディレクトリ: {options.OutputDir}");
                if (options.Mode == ExecutionMode.Code || options.Mode == ExecutionMode.All || options.Mode == ExecutionMode.DataCode)
                {
                    Log($"テンプレートファイル: {options.TemplateFilePath}");
                }

                // テンプレート読み込み
                string templateContent = null;
                if (options.Mode == ExecutionMode.Code || options.Mode == ExecutionMode.All || options.Mode == ExecutionMode.DataCode)
                {
                    templateContent = File.ReadAllText(options.TemplateFilePath, Encoding.UTF8);
                }

                foreach (var ws in workbook.Worksheets)
                {
                    foreach (var table in ws.Tables)
                    {
                        string tableName = table.Name;

                        // Filtering
                        if (!string.IsNullOrEmpty(options.TargetTableName))
                        {
                            if (!tableName.Equals(options.TargetTableName, StringComparison.OrdinalIgnoreCase))
                            {
                                continue;
                            }
                        }

                        Log($"テーブル処理中: {tableName}");

                        // 1. XML生成
                        if (options.Mode == ExecutionMode.Xml || options.Mode == ExecutionMode.All)
                        {
                            GenerateXmlFromExcel(workbook, tableName, options.OutputDir, options.ExcelPath);
                        }

                        // 2. C#コード生成
                        if (options.Mode == ExecutionMode.Code || options.Mode == ExecutionMode.All)
                        {
                            Log($"テンプレートを使用してクラスを生成: {Path.GetFileName(options.TemplateFilePath)}");
                            GenerateCSharpFromTemplate(workbook, tableName, options.OutputDir, templateContent);
                        }

                        // 3. データスクリプト生成
                        if (options.Mode == ExecutionMode.DataCode)
                        {
                            Log($"テンプレートを使用してデータスクリプトを生成: {Path.GetFileName(options.TemplateFilePath)}");
                            GenerateScriptFromData(workbook, tableName, options.OutputDir, templateContent);
                        }
                    }
                }
            }
        }

        private static void ListTables(XLWorkbook workbook)
        {
            foreach (var ws in workbook.Worksheets)
            {
                foreach (var table in ws.Tables)
                {
                    Console.WriteLine(table.Name);
                    LogFileOnly(table.Name);
                }
            }
        }

        private static string _logFilePath;

        private static void SetupLogger()
        {
            try
            {
                // 開発環境(bin/Release)の場合、プロジェクトルートのLogフォルダに出力したい
                string baseDir = AppDomain.CurrentDomain.BaseDirectory;
                string projectLogDir = Path.GetFullPath(Path.Combine(baseDir, @"..\..\Log"));

                string targetLogDir;
                if (Directory.Exists(projectLogDir))
                {
                    targetLogDir = projectLogDir;
                }
                else
                {
                    targetLogDir = Path.Combine(baseDir, "Log");
                    if (!Directory.Exists(targetLogDir))
                    {
                        Directory.CreateDirectory(targetLogDir);
                    }
                }
                _logFilePath = Path.Combine(targetLogDir, "execution_log.txt");
            }
            catch { }
        }

        private static void Log(string message)
        {
            Console.WriteLine(message);
            LogFileOnly(message);
        }

        private static void LogFileOnly(string message)
        {
            if (!string.IsNullOrEmpty(_logFilePath))
            {
                try
                {
                    File.AppendAllText(_logFilePath, message + Environment.NewLine, Encoding.UTF8);
                }
                catch { }
            }
        }

        private static XLWorkbook OpenWorkbookReadOnly(string path)
        {
            FileStream fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            return new XLWorkbook(fs);
        }

        private static Dictionary<string, List<ColumnInfo>> _headerCache = new Dictionary<string, List<ColumnInfo>>();
        private static Dictionary<string, Dictionary<string, IXLRangeRow>> _tableIndexCache = new Dictionary<string, Dictionary<string, IXLRangeRow>>();

        private static List<ColumnInfo> GetHeaders(XLWorkbook workbook, string tableName)
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

        private static IXLRangeRow FindRow(XLWorkbook workbook, string tableName, string keyColumn, string value)
        {
            string cacheKey = tableName + "::" + keyColumn;
            if (!_tableIndexCache.TryGetValue(cacheKey, out var index))
            {
                index = new Dictionary<string, IXLRangeRow>();
                var table = workbook.Table(tableName);
                var headers = GetHeaders(workbook, tableName);
                
                var keyColInfo = headers.FirstOrDefault(h => h.PropertyName.Equals(keyColumn, StringComparison.OrdinalIgnoreCase));
                
                if (keyColInfo != null)
                {
                    foreach (var row in table.DataRange.Rows())
                    {
                        try
                        {
                            // 行全体が空の場合はスキップ
                            if (row.IsEmpty()) continue;

                            string val = row.WorksheetRow().Cell(keyColInfo.ColumnNumber).GetValue<string>();
                            if (!index.ContainsKey(val)) index[val] = row;
                        }
                        catch { }
                    }
                }
                _tableIndexCache[cacheKey] = index;
            }

            if (index.TryGetValue(value, out var foundRow)) return foundRow;
            return null;
        }

        private static void GenerateXmlFromExcel(XLWorkbook workbook, string tableName, string baseOutputDir, string excelFilePath)
        {
            string outputDir = Path.Combine(baseOutputDir, "xml", tableName);
            if (!Directory.Exists(outputDir)) Directory.CreateDirectory(outputDir);

            var table = workbook.Table(tableName);
            var headers = GetHeaders(workbook, tableName);
            
            foreach (var row in table.DataRange.Rows())
            {
                XElement rootElement = CreateElementFromRow(row, headers, workbook, "Record");
                
                // ID取得（ファイル名用）
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
            Console.WriteLine($"Generated XMLs in {outputDir}");
        }

        private static XElement CreateElementFromRow(IXLRangeRow row, List<ColumnInfo> headers, XLWorkbook workbook, string elementName)
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
                    // テーブル参照の場合
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
                                // 再帰呼び出し。要素名はプロパティ名とする
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

        private static void GenerateCSharpFromTemplate(XLWorkbook workbook, string tableName, string baseOutputDir, string template)
        {
            var table = workbook.Table(tableName);
            var headers = table.HeadersRow().CellsUsed()
                .Select(c => c.GetValue<string>())
                .Where(h => !h.TrimStart().StartsWith("#")) // #で始まる列を無視
                .Select(h => new ColumnInfo(h))
                .ToList();

            string rootClassName = tableName;
            
            // 1. ツリー構築
            var rootNode = new ClassNode(rootClassName, rootClassName);
            rootNode.IsRoot = true;

            foreach (var col in headers)
            {
                rootNode.AddPath(col.PathParts, col.TypeName, col.IsArray, rootClassName);
            }

            // 2. マクロ処理
            string processedTemplate = ProcessTemplateMacros(template, rootNode, rootClassName);

            // 3. ルートプロパティ
            string rootPropertiesCode = BuildPropertiesCodeOnly(rootNode, "        ");

            // 4. グローバル変数置換
            string finalCode = processedTemplate
                .Replace("@TableName", rootClassName)
                .Replace("@GeneratedDate", DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss"))
                .Replace("@RootProperties", rootPropertiesCode.TrimEnd());

            // 5. 最終条件分岐
            finalCode = ProcessConditionals(finalCode);

            // 5.1 重複行削除
            finalCode = ProcessEraseDuplicatedLines(finalCode);

            // 6. 改行コード正規化
            finalCode = finalCode.Replace("\r\n", "\n").Replace("\n", "\r\n");

            // 保存
            // 出力パス: [baseOutputDir]/code/[TableName].cs
            string outputDir = Path.Combine(baseOutputDir, "code");
            if (!Directory.Exists(outputDir)) Directory.CreateDirectory(outputDir);
            
            string outputPath = Path.Combine(outputDir, $"{tableName}.cs");
            File.WriteAllText(outputPath, finalCode, Encoding.UTF8);
            Console.WriteLine($"Generated Class: {outputPath}");
        }

        public static void GenerateScriptFromData(XLWorkbook workbook, string tableName, string outputDir, string template)
        {
            var table = workbook.Table(tableName);
            var headers = table.HeadersRow().CellsUsed()
                .Select(c => c.GetValue<string>())
                .Where(h => !h.TrimStart().StartsWith("#")) // #で始まる列を無視
                .Select(h => new ColumnInfo(h))
                .ToList();

            // データ準備
            // List of Dictionary<PropertyPath, Value>
            var dataRows = new List<Dictionary<string, string>>();

            foreach (var row in table.DataRange.Rows())
            {
                var rowDict = new Dictionary<string, string>();
                int cellIndex = 0;
                foreach (var cell in row.Cells())
                {
                    if (cellIndex >= headers.Count) break;
                    var colInfo = headers[cellIndex];
                    string val = cell.GetValue<string>();
                    
                    // キー: "Id", "Properties.Suit" など
                    // ヘッダ情報からキーを再構築します
                    string key = string.Join(".", colInfo.PathParts);
                    rowDict[key] = val;
                    
                    cellIndex++;
                }
                dataRows.Add(rowDict);
            }

            // テンプレート処理
            string finalCode = ProcessDataMacros(template, dataRows);
            
            // スクリプトテンプレートでも @TableName マクロを使用可能にします
            finalCode = finalCode.Replace("@TableName", tableName)
                                 .Replace("@GeneratedDate", DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss"));

            // 最終条件分岐（ループ外に残っている可能性があるため）
            finalCode = ProcessConditionals(finalCode);

            // 重複行削除
            finalCode = ProcessEraseDuplicatedLines(finalCode);

            // 改行コード正規化
            finalCode = finalCode.Replace("\r\n", "\n").Replace("\n", "\r\n");

            // 保存
            // 出力パス: [baseOutputDir]/code/[TableName]_Data.cs
            // baseOutputDir がファイル指定の場合はそのパスを使用
            string outputPath;
            if (Path.HasExtension(outputDir))
            {
                outputPath = outputDir;
                string dir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir)) Directory.CreateDirectory(dir);
            }
            else
            {
                string dir = Path.Combine(outputDir, "code"); 
                if (!Directory.Exists(dir)) Directory.CreateDirectory(dir);
                outputPath = Path.Combine(dir, $"{tableName}_Data.cs");
            }
            
            File.WriteAllText(outputPath, finalCode, Encoding.UTF8);
            Console.WriteLine($"Generated Script: {outputPath}");
        }

        private static string ProcessDataMacros(string template, List<Dictionary<string, string>> dataRows)
        {
            string startTag = "#ForAllData";
            string endTag = "#EndForAllData";

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

                foreach (var rowDict in dataRows)
                {
                    string instance = blockTemplate;
                    
                    // 変数置換 ${Key}
                    foreach (var kvp in rowDict)
                    {
                        instance = instance.Replace($"${{{kvp.Key}}}", kvp.Value);
                    }

                    // 行ごとに条件分岐も処理します
                    instance = ProcessConditionals(instance);

                    loopContent.Append(instance);
                }

                int removeEndIndex = GetRemoveEndIndex(template, endIdx + endTag.Length);
                int removeLength = removeEndIndex - startIdx;

                template = template.Remove(startIdx, removeLength).Insert(startIdx, loopContent.ToString());
            }

            return template;
        }

        private static string ProcessTemplateMacros(string template, ClassNode rootNode, string rootTableName)
        {
            string startTag = "#ForAllSubClasses";
            string endTag = "#EndForAllSubClasses";

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

                var subNodes = rootNode.GetAllNodes().Where(n => !n.IsRoot).ToList();
                StringBuilder loopContent = new StringBuilder();

                foreach (var node in subNodes)
                {
                    string instance = blockTemplate
                        .Replace("@SubClassName", node.ClassName)
                        .Replace("@SubClassTagName", node.XmlTagName)
                        .Replace("@TableName", rootTableName);

                    instance = ProcessInnerPropertyMacros(instance, node);
                    instance = ProcessConditionals(instance);

                    loopContent.Append(instance);
                }

                int removeEndIndex = GetRemoveEndIndex(template, endIdx + endTag.Length);
                int removeLength = removeEndIndex - startIdx;

                template = template.Remove(startIdx, removeLength).Insert(startIdx, loopContent.ToString());
            }

            return template;
        }

        private static string ProcessInnerPropertyMacros(string template, ClassNode node)
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

                var allProperties = new List<dynamic>();
                foreach (var prop in node.Properties)
                    allProperties.Add(new { Name = prop.Name, Type = ConvertType(prop.TypeName, prop.IsArray) });
                foreach (var child in node.Children)
                    allProperties.Add(new { Name = child.XmlTagName, Type = child.ClassName });

                foreach (var prop in allProperties)
                {
                    string instance = blockTemplate
                        .Replace("@SubClassPropertyName", prop.Name)
                        .Replace("@SubClassPropertyType", prop.Type);

                    instance = ProcessConditionals(instance);

                    loopContent.Append(instance);
                }

                int removeEndIndex = GetRemoveEndIndex(template, endIdx + endTag.Length);
                int removeLength = removeEndIndex - startIdx;

                template = template.Remove(startIdx, removeLength).Insert(startIdx, loopContent.ToString());
            }

            return template;
        }

        private static string ProcessConditionals(string text)
        {
            text = ProcessExpressionMacros(text);
            return ProcessIfBlocks(text);
        }

        private static string ProcessExpressionMacros(string text)
        {
            var targetMacros = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                "Eq", "Not", "And", "Or", "Contains", "Replace"
            };

            var regex = new Regex(@"#(\w+)\s*\(([^()]*)\)");
            bool changed = true;

            while (changed)
            {
                changed = false;
                text = regex.Replace(text, match =>
                {
                    string macroName = match.Groups[1].Value;
                    string argsStr = match.Groups[2].Value;

                    if (!targetMacros.Contains(macroName)) return match.Value;

                    string result = EvaluateMacro(macroName, argsStr);
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

        private static string EvaluateMacro(string name, string argsStr)
        {
            var args = argsStr.Split(',').Select(a => a.Trim()).ToArray();
            switch (name.ToLower())
            {
                case "eq":
                    if (args.Length < 2) return "False";
                    return (args[0] == args[1]) ? "True" : "False";
                case "not":
                    if (args.Length < 1) return "False";
                    return (args[0].Equals("True", StringComparison.OrdinalIgnoreCase)) ? "False" : "True";
                case "and":
                    if (args.Length == 0) return "False";
                    return args.All(a => a.Equals("True", StringComparison.OrdinalIgnoreCase)) ? "True" : "False";
                case "or":
                    if (args.Length == 0) return "False";
                    return args.Any(a => a.Equals("True", StringComparison.OrdinalIgnoreCase)) ? "True" : "False";
                case "contains":
                    if (args.Length < 2) return "False";
                    return args[0].Contains(args[1]) ? "True" : "False";
                case "replace":
                    if (args.Length < 3) return args.Length > 0 ? args[0] : "";
                    return args[0].Replace(args[1], args[2]);
                default:
                    return $"#{name}({argsStr})";
            }
        }

        private static string ProcessEraseDuplicatedLines(string text)
        {
            string startTag = "#EraseDuplicatedLine";
            string endTag = "#EndErase";

            while (true)
            {
                int startIdx = text.IndexOf(startTag);
                if (startIdx == -1) break;

                int endIdx = text.IndexOf(endTag, startIdx);
                if (endIdx == -1) break;



                // startIdx を行全体を含むように拡張
                int realStartIdx = startIdx;
                int checkPos = startIdx - 1;
                while (checkPos >= 0)
                {
                    char c = text[checkPos];
                    if (c == '\n' || c == '\r') break; // 行頭（改行の後）
                    if (c != ' ' && c != '\t')
                    {
                        // 同一行に他のコンテンツがある場合（例: "code(); #Erase"）
                        // この場合は行全体を削除できません。
                        realStartIdx = startIdx;
                        break;
                    }
                    realStartIdx = checkPos;
                    checkPos--;
                }
                
                // endIdx を行末まで拡張
                int realEndIdx = endIdx + endTag.Length;
                
                // endタグ後の改行まで含める
                checkPos = realEndIdx;
                while (checkPos < text.Length)
                {
                    char c = text[checkPos];
                    if (c == '\n')
                    {
                        realEndIdx = checkPos + 1; // \n を含める
                        break;
                    }
                    if (c == '\r')
                    {
                        realEndIdx = checkPos + 1; // \r を含める
                         if (realEndIdx < text.Length && text[realEndIdx] == '\n')
                            realEndIdx++; // \r\n の \n も含める
                        break;
                    }
                    if (c != ' ' && c != '\t')
                    {
                        // タグの後ろにコンテンツがある場合は何もしない（通常は #EndErase でブロックが終わるはず）
                    }
                    checkPos++;
                }

                // タグの内側のコンテンツを取得
                int contentStart = startIdx + startTag.Length;
                int contentLength = endIdx - contentStart;
                
                string content = text.Substring(contentStart, contentLength);
                string processedContent = RemoveDuplicates(content);
                
                // 開始行、終了行、および内側のコンテンツを置換対象とするための範囲計算
                // 1. 開始タグ行のインデントをさかのぼる
                int removeStart = startIdx;
                bool removeStartLine = true;
                for (int i = startIdx - 1; i >= 0; i--)
                {
                    if (text[i] == '\n' || text[i] == '\r') { removeStart = i + 1; break; } // 改行の後（行頭）
                    if (text[i] != ' ' && text[i] != '\t') { removeStartLine = false; break; } // インデント以外が含まれる
                    removeStart = i;
                }
                
                // 2. 開始タグ後の改行まで進む
                int removeStartEnd = startIdx + startTag.Length;
                if (removeStartLine)
                {
                    if (removeStartEnd < text.Length && text[removeStartEnd] == '\r') removeStartEnd++;
                    if (removeStartEnd < text.Length && text[removeStartEnd] == '\n') removeStartEnd++;
                }

                // 3. 終了タグ行のインデントをさかのぼる
                int removeEndStart = endIdx;
                 bool removeEndLine = true;
                for (int i = endIdx - 1; i >= removeStartEnd; i--)
                {
                     if (text[i] == '\n' || text[i] == '\r') { removeEndStart = i + 1; break; }
                     if (text[i] != ' ' && text[i] != '\t') { removeEndLine = false; break; }
                     removeEndStart = i;
                }
                
                // 4. 終了タグ後の改行まで進む
                int removeEndEnd = endIdx + endTag.Length;
                if (removeEndLine)
                {
                     if (removeEndEnd < text.Length && text[removeEndEnd] == '\r') removeEndEnd++;
                     if (removeEndEnd < text.Length && text[removeEndEnd] == '\n') removeEndEnd++;
                }

                // 行全体でない場合は、純粋にタグ部分のみを削除対象とします
                int finalStart = removeStartLine ? removeStart : startIdx;
                int finalStartContent = removeStartLine ? removeStartEnd : (startIdx + startTag.Length);
                
                int finalEndContent = removeEndLine ? removeEndStart : endIdx;
                int finalEnd = removeEndLine ? removeEndEnd : (endIdx + endTag.Length);

                string innerContent = text.Substring(finalStartContent, finalEndContent - finalStartContent);
                string processed = RemoveDuplicates(innerContent);
                
                // 置換実行
                text = text.Remove(finalStart, finalEnd - finalStart).Insert(finalStart, processed);
            }
            return text;
        }

        private static string RemoveDuplicates(string content)
        {
            // 行分割
            // 元の改行構造を保持します
            var lines = new List<string>();
            using (StringReader sr = new StringReader(content))
            {
                string line;
                while ((line = sr.ReadLine()) != null)
                {
                    lines.Add(line);
                }
            }

            var seen = new HashSet<string>();
            var result = new StringBuilder();
            
            for (int i = 0; i < lines.Count; i++)
            {
                string line = lines[i];
                string trimmed = line.Trim();
                
                // 空行の場合はどうするか？
                // 通常、空行は構造的な意味を持つため維持したい場合が多いですが、
                // 重複除去のコンテキストでは、連続する空行などを整理したい場合もあります。
                // ここでは、「空行であっても、過去に出現していなければ出力」というロジックにします。
                // ただし、Trim() した結果が同一であれば重複とみなします（インデント違いの空行など）。
                if (string.IsNullOrWhiteSpace(trimmed))
                {
                    if (!seen.Contains(trimmed))
                    {
                        seen.Add(trimmed);
                        result.AppendLine(line);
                    }
                }
                else
                {
                    if (!seen.Contains(trimmed))
                    {
                        seen.Add(trimmed);
                        result.AppendLine(line);
                    }
                }
            }
            // 末尾の改行を除去するか？
            // StringReader/AppendLine は標準的な改行を追加します。
            // 元の文字列の末尾状態を維持したいですが、AppendLineしているので末尾に改行がつきます。
            // 結合時の不具合を防ぐため、そのまま返します。
            return result.ToString(); 
        }

        private static string ProcessIfBlocks(string text)
        {
            while (true)
            {
                int ifIndex = text.IndexOf("#If");
                if (ifIndex == -1) break;

                int endIndex = FindMatchingEndif(text, ifIndex);
                if (endIndex == -1) break;

                int length = (endIndex + "#Endif".Length) - ifIndex;
                int removeEndIndex = GetRemoveEndIndex(text, ifIndex + length);
                
                // 修正: #If の前のインデントも削除する
                int removeStartIndex = ifIndex;
                int currentPos = ifIndex - 1;
                while (currentPos >= 0)
                {
                    char c = text[currentPos];
                    if (c == ' ' || c == '\t')
                    {
                        removeStartIndex = currentPos;
                        currentPos--;
                    }
                    else if (c == '\n' || c == '\r')
                    {
                        // 行頭（または改行後の空白）が見つかった
                        break;
                    }
                    else
                    {
                        // 同一行に非空白文字が見つかった
                        // #If がインラインで記述されている場合はインデント削除を行わない
                        removeStartIndex = ifIndex;
                        break;
                    }
                }

                int fullRemoveLength = removeEndIndex - removeStartIndex;

                string result = SolveIfBlock(text.Substring(ifIndex, length));
                text = text.Remove(removeStartIndex, fullRemoveLength).Insert(removeStartIndex, result);
            }
            return text;
        }

        private static string SolveIfBlock(string block)
        {
            var segments = ParseIfSegments(block);
            foreach (var seg in segments)
            {
                if (seg.Type == "Else") return TrimStartNewline(seg.Content);
                if (seg.Condition == "True") return TrimStartNewline(seg.Content);
            }
            return "";
        }

        private class IfSegment
        {
            public string Type;
            public string Condition;
            public string Content;
        }

        private static List<IfSegment> ParseIfSegments(string block)
        {
            var list = new List<IfSegment>();
            using (StringReader sr = new StringReader(block))
            {
                string line;
                string currentType = null;
                string currentCond = null;
                StringBuilder currentContent = new StringBuilder();
                int depth = 0;

                while ((line = sr.ReadLine()) != null)
                {
                    string trimmed = line.Trim();
                    if (trimmed.StartsWith("#If"))
                    {
                        if (depth == 0)
                        {
                            currentType = "If";
                            currentContent.Clear();
                            currentCond = ExtractCondition(trimmed);
                        }
                        else
                        {
                            currentContent.AppendLine(line);
                        }
                        depth++;
                        continue;
                    }
                    if (trimmed.StartsWith("#Endif"))
                    {
                        depth--;
                        if (depth == 0)
                        {
                            if (currentType != null)
                            {
                                list.Add(new IfSegment { Type = currentType, Condition = currentCond, Content = currentContent.ToString() });
                            }
                            currentType = null;
                        }
                        else
                        {
                            currentContent.AppendLine(line);
                        }
                        continue;
                    }
                    if (depth == 1)
                    {
                        if (trimmed.StartsWith("#Elif"))
                        {
                            if (currentType != null)
                                list.Add(new IfSegment { Type = currentType, Condition = currentCond, Content = currentContent.ToString() });
                            currentType = "Elif";
                            currentContent.Clear();
                            currentCond = ExtractCondition(trimmed);
                            continue;
                        }
                        if (trimmed.StartsWith("#Else"))
                        {
                            if (currentType != null)
                                list.Add(new IfSegment { Type = currentType, Condition = currentCond, Content = currentContent.ToString() });
                            currentType = "Else";
                            currentContent.Clear();
                            currentCond = "True";
                            continue;
                        }
                    }
                    if (currentType != null) currentContent.AppendLine(line);
                }
            }
            return list;
        }

        private static string ExtractCondition(string line)
        {
            int start = line.IndexOf('(');
            int end = line.LastIndexOf(')');
            if (start != -1 && end != -1) return line.Substring(start + 1, end - start - 1).Trim();
            return "False";
        }

        private static int FindMatchingEndif(string text, int startIfIndex)
        {
            int depth = 0;
            int pos = startIfIndex;
            while (pos < text.Length)
            {
                int nextIf = text.IndexOf("#If", pos);
                int nextEnd = text.IndexOf("#Endif", pos);
                if (nextEnd == -1) return -1;
                if (nextIf != -1 && nextIf < nextEnd)
                {
                    depth++;
                    pos = nextIf + 3;
                }
                else
                {
                    depth--;
                    pos = nextEnd + 6;
                    if (depth == 0) return nextEnd;
                }
            }
            return -1;
        }

        private static string TrimStartNewline(string text)
        {
            if (string.IsNullOrEmpty(text)) return text;
            if (text.StartsWith("\r\n")) return text.Substring(2);
            if (text.StartsWith("\n")) return text.Substring(1);
            return text;
        }

        private static int GetRemoveEndIndex(string fullText, int indexAfterTag)
        {
            int idx = indexAfterTag;
            if (idx < fullText.Length && fullText[idx] == '\r') idx++;
            if (idx < fullText.Length && fullText[idx] == '\n') idx++;
            return idx;
        }

        private static string BuildPropertiesCodeOnly(ClassNode node, string indent)
        {
            var allItems = new List<string>();
            foreach (var prop in node.Properties)
            {
                string type = ConvertType(prop.TypeName, prop.IsArray);
                string item = $"{indent}[XmlElement(\"{prop.Name}\")]\n{indent}public {type} {prop.Name} {{ get; set; }}";
                allItems.Add(item);
            }
            foreach (var child in node.Children)
            {
                string item = $"{indent}[XmlElement(\"{child.XmlTagName}\")]\n{indent}public {child.ClassName} {child.XmlTagName} {{ get; set; }}";
                allItems.Add(item);
            }
            return string.Join("\n\n", allItems);
        }

        private static string ConvertType(string typeName, bool isArray)
        {
            string baseType;
            // プリミティブ型判定（小文字で比較）
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

        public class ColumnInfo
        {
            public string OriginalHeader { get; set; }
            public string[] PathParts { get; set; }
            public string PropertyName { get; set; }
            public string TypeName { get; set; }
            public bool IsArray { get; set; }
            
            // 参照テーブル情報
            public string RefTableName { get; set; }
            public string RefKeyColumn { get; set; }
            public int ColumnNumber { get; set; } // 1-based absolute column number

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
                        // 構文解析: TableName(KeyName)[] 
                        // Regexでの抽出
                        // 例: Table_SubClass(ParamName)[] -> Group1:Table_SubClass, Group2:ParamName, Group3:[]
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
                            // 従来の型指定
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

        public class ClassNode
        {
            public string XmlTagName { get; set; }
            public string ClassName { get; set; }
            public bool IsRoot { get; set; } = false;
            public List<PropertyNode> Properties { get; set; } = new List<PropertyNode>();
            public List<ClassNode> Children { get; set; } = new List<ClassNode>();
            public ClassNode(string xmlTagName, string className) { XmlTagName = xmlTagName; ClassName = className; }
            public void AddPath(string[] parts, string typeName, bool isArray, string parentPrefix)
            {
                if (parts.Length == 1)
                {
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
                child.AddPath(parts.Skip(1).ToArray(), typeName, isArray, child.ClassName);
            }
            public List<ClassNode> GetAllNodes()
            {
                var list = new List<ClassNode> { this };
                foreach (var child in Children) list.AddRange(child.GetAllNodes());
                return list;
            }
        }

        public class PropertyNode
        {
            public string Name { get; set; }
            public string TypeName { get; set; }
            public bool IsArray { get; set; }
        }
    }
}
