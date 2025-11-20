using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using System.Xml.Linq; // ★ XML生成に使用するライブラリ

namespace XmlWriter
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            cmbSheetName.Enabled = false;
        }

        // 参照ボタンクリック時 (変更なし)
        private void btnBrowse_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog ofd = new OpenFileDialog())
            {
                ofd.Filter = "Excel Files|*.xlsx";
                ofd.Title = "Excelファイルを選択してください";
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    txtFilePath.Text = ofd.FileName;
                    LoadTableNames(ofd.FileName);
                }
            }
        }

        // Excelからテーブル名一覧を読み込む (変更なし)
        private void LoadTableNames(string filePath)
        {
            cmbSheetName.Items.Clear();
            cmbSheetName.Enabled = false;
            UpdateStatus("テーブル名を読み込み中...");

            if (!File.Exists(filePath))
            {
                UpdateStatus("ファイルが見つかりません。");
                return;
            }

            try
            {
                using (var workbook = new XLWorkbook(filePath))
                {
                    bool foundTable = false;
                    foreach (var worksheet in workbook.Worksheets)
                    {
                        foreach (var table in worksheet.Tables)
                        {
                            cmbSheetName.Items.Add(table.Name);
                            foundTable = true;
                        }
                    }

                    if (foundTable)
                    {
                        cmbSheetName.SelectedIndex = 0;
                        cmbSheetName.Enabled = true;
                        UpdateStatus("テーブル名の読み込み完了。");
                    }
                    else
                    {
                        UpdateStatus("Excelファイルに「挿入＞テーブル」で作成されたテーブルが見つかりません。");
                    }
                }
            }
            catch (Exception ex)
            {
                cmbSheetName.Items.Add("読み込みエラー");
                UpdateStatus($"テーブル名読み込みエラー: {ex.Message}");
                MessageBox.Show($"テーブル名読み込み中にエラーが発生しました:\n{ex.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


        // XML生成ボタンクリック時 (変更なし)
        private void btnGenerate_Click(object sender, EventArgs e)
        {
            string excelPath = txtFilePath.Text;

            if (cmbSheetName.SelectedItem == null)
            {
                MessageBox.Show("テーブル名を選択してください。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            string tableName = cmbSheetName.SelectedItem.ToString();

            if (string.IsNullOrEmpty(excelPath) || !File.Exists(excelPath))
            {
                MessageBox.Show("有効なExcelファイルを選択してください。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            try
            {
                UpdateStatus("処理中...");

                GenerateXmlFromExcel(excelPath, tableName);

                UpdateStatus("完了");
                MessageBox.Show("XMLファイルの生成が完了しました。", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                UpdateStatus("エラー発生");
                MessageBox.Show($"エラーが発生しました:\n{ex.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // ★ 修正: XElementを使用してカスタムXMLを生成し、フォルダ構成を変更
        private void GenerateXmlFromExcel(string filePath, string tableName)
        {
            // 1. 出力フォルダの設定
            // Output_XMLの下に、テーブル名と同名のフォルダを作成
            string baseDir = Path.GetDirectoryName(filePath);
            string outputDir = Path.Combine(baseDir, "Output_XML", tableName);

            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            using (var workbook = new XLWorkbook(filePath))
            {
                IXLTable table = null;
                try
                {
                    table = workbook.Table(tableName);
                }
                catch (KeyNotFoundException)
                {
                    throw new InvalidOperationException($"指定されたテーブル名 '{tableName}' がブック内に見つかりません。");
                }

                // ヘッダーを取得
                var headerRow = table.HeadersRow();
                List<string> headers = new List<string>();
                foreach (var cell in headerRow.CellsUsed())
                {
                    headers.Add(cell.GetValue<string>());
                }

                // データ行をループ
                foreach (var row in table.DataRange.Rows())
                {
                    XElement rootElement = new XElement("Record"); // XMLのルート要素

                    string idValue = null;

                    int cellIndex = 0;
                    foreach (var cell in row.Cells())
                    {
                        string colName = headers[cellIndex];
                        string val = cell.GetValue<string>();

                        // ★ <XX>Value</XX> 形式で要素を作成し、ルートに追加
                        rootElement.Add(new XElement(colName, val));

                        // ファイル名用のID取得
                        if (colName.Equals("ID", StringComparison.OrdinalIgnoreCase))
                        {
                            idValue = val;
                        }

                        cellIndex++;
                    }

                    // 2. ファイル名の生成とXML書き出し

                    // ID値のバリデーションとゼロ埋め
                    if (string.IsNullOrEmpty(idValue))
                    {
                        throw new InvalidOperationException($"テーブル '{tableName}' の行 (シート:{row.WorksheetRow().RowNumber()}行目) に 'ID' 列の値が見つかりませんでした。");
                    }

                    // IDを数値として解析し、6桁にゼロ埋め
                    if (!long.TryParse(idValue, out long idNum))
                    {
                        // IDが数値ではない場合、そのまま使用するが警告
                        MessageBox.Show($"ID列の値 '{idValue}' が数値ではないため、ゼロ埋めせずそのまま使用します。", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }

                    // 6桁のゼロ埋め文字列を作成 (IDが数値でなければそのままの文字列)
                    string formattedId = (idNum != 0) ? idNum.ToString("D6") : idValue;

                    // ファイル名: X_YYYYYY.xml
                    string fileName = $"{tableName}_{formattedId}.xml";
                    string fullPath = Path.Combine(outputDir, fileName);

                    // XMLファイルを保存 (XElement.Saveを使用)
                    rootElement.Save(fullPath);
                }
            }
        }

        private void UpdateStatus(string msg)
        {
            lblStatus.Text = msg;
            Application.DoEvents();
        }
    }
}