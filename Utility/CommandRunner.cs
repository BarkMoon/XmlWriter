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
            List
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
            Log($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] Start Execution");
            Log($"Mode: {options.Mode}");
            Log($"Excel: {options.ExcelPath}");
            if (!string.IsNullOrEmpty(options.TargetTableName))
            {
                Log($"Target Table: {options.TargetTableName}");
            }

            if (!File.Exists(options.ExcelPath))
            {
                throw new FileNotFoundException("Excel file not found", options.ExcelPath);
            }

            // Mode Check for required paths
            if (options.Mode == ExecutionMode.Code || options.Mode == ExecutionMode.All)
            {
                if (string.IsNullOrEmpty(options.TemplateFilePath) || !File.Exists(options.TemplateFilePath))
                {
                    throw new FileNotFoundException("Template file not found (Required for Code generation)", options.TemplateFilePath);
                }
            }

            using (var workbook = OpenWorkbookReadOnly(options.ExcelPath))
            {
                // List Mode
                if (options.Mode == ExecutionMode.List)
                {
                    ListTables(workbook);
                    return;
                }

                Log($"Output Directory: {options.OutputDir}");
                if (options.Mode == ExecutionMode.Code || options.Mode == ExecutionMode.All)
                {
                    Log($"Template File: {options.TemplateFilePath}");
                }

                // Prepare Template Content
                string templateContent = null;
                if (options.Mode == ExecutionMode.Code || options.Mode == ExecutionMode.All)
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

                        Log($"Processing Table: {tableName}");

                        // 1. Generate XML
                        if (options.Mode == ExecutionMode.Xml || options.Mode == ExecutionMode.All)
                        {
                            GenerateXmlFromExcel(workbook, tableName, options.OutputDir, options.ExcelPath);
                        }

                        // 2. Generate C#
                        if (options.Mode == ExecutionMode.Code || options.Mode == ExecutionMode.All)
                        {
                            Log($"Generating Class using template: {Path.GetFileName(options.TemplateFilePath)}");
                            GenerateCSharpFromTemplate(workbook, tableName, options.OutputDir, templateContent);
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

        private static void GenerateXmlFromExcel(XLWorkbook workbook, string tableName, string baseOutputDir, string excelFilePath)
        {
            // Output path: <baseOutputDir> (User specified)
            // Original tool logic: [ExcelDir]/Output_XML/[TableName]
            // Modified logic: [baseOutputDir]/Output_XML/[TableName] ? No, User said "CardGameElements_Data\out"
            // Let's output to [baseOutputDir]/[TableName]_XML/ or just [baseOutputDir] ?
            // The tool originally creates a subdirectory "Output_XML/[TableName]".
            // Let's respect the structure but put it inside baseOutputDir.
            // Actually, the user asked to generate "xml data and source code under CardGameElements_Data\out".
            // So maybe "out/xml/[TableName]" and "out/code/[TableName].cs"?
            // Or "out/xml/[TableName]_[ID].xml" flatly?
            // The original tool does: [Dir]/Output_XML/[TableName]/[TableName]_[ID].xml
            // Let's do: [baseOutputDir]/xml/[TableName]/...
            
            // To match original behavior somewhat but adapt to "out folder":
            // I'll create [baseOutputDir]/xml/[TableName] for XMLs.
            string outputDir = Path.Combine(baseOutputDir, "xml", tableName);
            if (!Directory.Exists(outputDir)) Directory.CreateDirectory(outputDir);

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

                    if (colInfo.IsArray)
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

                    if (colInfo.PropertyName.Equals("ID", StringComparison.OrdinalIgnoreCase)) idValue = rawVal;
                    cellIndex++;
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

        private static void GenerateCSharpFromTemplate(XLWorkbook workbook, string tableName, string baseOutputDir, string template)
        {
            var table = workbook.Table(tableName);
            var headers = table.HeadersRow().CellsUsed()
                .Select(c => new ColumnInfo(c.GetValue<string>()))
                .ToList();

            string rootClassName = tableName;
            
            // 1. Build Tree
            var rootNode = new ClassNode(rootClassName, rootClassName);
            rootNode.IsRoot = true;

            foreach (var col in headers)
            {
                rootNode.AddPath(col.PathParts, col.TypeName, col.IsArray, rootClassName);
            }

            // 2. Process Macros
            string processedTemplate = ProcessTemplateMacros(template, rootNode, rootClassName);

            // 3. Root Properties
            string rootPropertiesCode = BuildPropertiesCodeOnly(rootNode, "        ");

            // 4. Global Variables replacement
            string finalCode = processedTemplate
                .Replace("@TableName", rootClassName)
                .Replace("@GeneratedDate", DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss"))
                .Replace("@RootProperties", rootPropertiesCode.TrimEnd());

            // 5. Final Conditionals
            finalCode = ProcessConditionals(finalCode);

            // 6. Newline normalization
            finalCode = finalCode.Replace("\r\n", "\n").Replace("\n", "\r\n");

            // Save
            // Output path: [baseOutputDir]/code/[TableName].cs
            string outputDir = Path.Combine(baseOutputDir, "code");
            if (!Directory.Exists(outputDir)) Directory.CreateDirectory(outputDir);
            
            string outputPath = Path.Combine(outputDir, $"{tableName}.cs");
            File.WriteAllText(outputPath, finalCode, Encoding.UTF8);
            Console.WriteLine($"Generated Class: {outputPath}");
        }

        // --- Logic Methods from Form1.cs (Adapted) ---

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
                int fullRemoveLength = removeEndIndex - ifIndex;

                string result = SolveIfBlock(text.Substring(ifIndex, length));
                text = text.Remove(ifIndex, fullRemoveLength).Insert(ifIndex, result);
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

        public class ColumnInfo
        {
            public string OriginalHeader { get; set; }
            public string[] PathParts { get; set; }
            public string PropertyName { get; set; }
            public string TypeName { get; set; }
            public bool IsArray { get; set; }
            
            public ColumnInfo(string header)
            {
                OriginalHeader = header;
                string namePart = header;
                TypeName = "string";
                IsArray = false;
                if (header.Contains(":"))
                {
                    var parts = header.Split(':');
                    namePart = parts[0].Trim();
                    if (parts.Length > 1)
                    {
                        string typeRaw = parts[1].Trim().ToLower();
                        if (typeRaw.EndsWith("[]"))
                        {
                            IsArray = true;
                            TypeName = typeRaw.Replace("[]", "");
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
