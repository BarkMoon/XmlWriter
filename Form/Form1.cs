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
            // Logger setting
            XmlWriter.Utility.CommandRunner.SetExternalLogger(msg => UpdateStatus(msg));
        }

        // ---------------------------------------------------------




        // ---------------------------------------------------------
        // ★ 新規: 読み取り専用でExcelを開くヘルパーメソッド
        // ---------------------------------------------------------


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
                using (var workbook = XmlWriter.Utility.CommandRunner.OpenWorkbookReadOnly(filePath))
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

        private void GenerateXmlFromExcel(string filePath, string tableName, string baseOutputDir)
        {
            using (var workbook = XmlWriter.Utility.CommandRunner.OpenWorkbookReadOnly(filePath))
            {
                XmlWriter.Utility.CommandRunner.GenerateXmlFromExcel(workbook, tableName, baseOutputDir, filePath);
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

                using (var workbook = XmlWriter.Utility.CommandRunner.OpenWorkbookReadOnly(txtFilePath.Text))
                {
                    // クラスコードの生成 (ロジックはCommandRunnerに委譲)
                    string finalCode = XmlWriter.Utility.CommandRunner.GenerateCSharpCode(workbook, tableName, templateContent);

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

        private void UpdateStatus(string msg) { lblStatus.Text = msg; Application.DoEvents(); }

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
                using (var workbook = XmlWriter.Utility.CommandRunner.OpenWorkbookReadOnly(txtFilePath.Text))
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