namespace XmlWriter
{
    partial class XmlWriter
    {
        /// <summary>
        /// 必要なデザイナー変数です。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 使用中のリソースをすべてクリーンアップします。
        /// </summary>
        /// <param name="disposing">マネージド リソースを破棄する場合は true を指定し、その他の場合は false を指定します。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows フォーム デザイナーで生成されたコード

        /// <summary>
        /// デザイナー サポートに必要なメソッドです。このメソッドの内容を
        /// コード エディターで変更しないでください。
        /// </summary>
        private void InitializeComponent()
        {
            this.buttonOpenFileDialog = new System.Windows.Forms.Button();
            this.textBoxExcelFilePath = new System.Windows.Forms.TextBox();
            this.buttonExcelToTables = new System.Windows.Forms.Button();
            this.listBoxTables = new System.Windows.Forms.ListBox();
            this.buttonTableToRows = new System.Windows.Forms.Button();
            this.treeViewRows = new System.Windows.Forms.TreeView();
            this.buttonWriteOutput = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // buttonOpenFileDialog
            // 
            this.buttonOpenFileDialog.Location = new System.Drawing.Point(58, 219);
            this.buttonOpenFileDialog.Name = "buttonOpenFileDialog";
            this.buttonOpenFileDialog.Size = new System.Drawing.Size(75, 23);
            this.buttonOpenFileDialog.TabIndex = 0;
            this.buttonOpenFileDialog.Text = "参照";
            this.buttonOpenFileDialog.UseVisualStyleBackColor = true;
            // 
            // textBoxExcelFilePath
            // 
            this.textBoxExcelFilePath.Location = new System.Drawing.Point(12, 159);
            this.textBoxExcelFilePath.Multiline = true;
            this.textBoxExcelFilePath.Name = "textBoxExcelFilePath";
            this.textBoxExcelFilePath.Size = new System.Drawing.Size(171, 19);
            this.textBoxExcelFilePath.TabIndex = 1;
            // 
            // buttonExcelToTables
            // 
            this.buttonExcelToTables.Location = new System.Drawing.Point(202, 136);
            this.buttonExcelToTables.Name = "buttonExcelToTables";
            this.buttonExcelToTables.Size = new System.Drawing.Size(72, 65);
            this.buttonExcelToTables.TabIndex = 2;
            this.buttonExcelToTables.Text = "Next";
            this.buttonExcelToTables.UseVisualStyleBackColor = true;
            // 
            // listBoxTables
            // 
            this.listBoxTables.FormattingEnabled = true;
            this.listBoxTables.ItemHeight = 12;
            this.listBoxTables.Location = new System.Drawing.Point(296, 41);
            this.listBoxTables.Name = "listBoxTables";
            this.listBoxTables.Size = new System.Drawing.Size(222, 256);
            this.listBoxTables.TabIndex = 3;
            // 
            // buttonTableToRows
            // 
            this.buttonTableToRows.Location = new System.Drawing.Point(542, 136);
            this.buttonTableToRows.Name = "buttonTableToRows";
            this.buttonTableToRows.Size = new System.Drawing.Size(72, 65);
            this.buttonTableToRows.TabIndex = 4;
            this.buttonTableToRows.Text = "Next";
            this.buttonTableToRows.UseVisualStyleBackColor = true;
            // 
            // treeViewRows
            // 
            this.treeViewRows.Location = new System.Drawing.Point(633, 41);
            this.treeViewRows.Name = "treeViewRows";
            this.treeViewRows.Size = new System.Drawing.Size(223, 255);
            this.treeViewRows.TabIndex = 5;
            // 
            // buttonWriteOutput
            // 
            this.buttonWriteOutput.Location = new System.Drawing.Point(750, 329);
            this.buttonWriteOutput.Name = "buttonWriteOutput";
            this.buttonWriteOutput.Size = new System.Drawing.Size(106, 34);
            this.buttonWriteOutput.TabIndex = 6;
            this.buttonWriteOutput.Text = "出力";
            this.buttonWriteOutput.UseVisualStyleBackColor = true;
            this.buttonWriteOutput.Click += new System.EventHandler(this.button4_Click);
            // 
            // XmlWriter
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(885, 376);
            this.Controls.Add(this.buttonWriteOutput);
            this.Controls.Add(this.treeViewRows);
            this.Controls.Add(this.buttonTableToRows);
            this.Controls.Add(this.listBoxTables);
            this.Controls.Add(this.buttonExcelToTables);
            this.Controls.Add(this.textBoxExcelFilePath);
            this.Controls.Add(this.buttonOpenFileDialog);
            this.Name = "XmlWriter";
            this.Text = "XmlWriter";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button buttonOpenFileDialog;
        private System.Windows.Forms.TextBox textBoxExcelFilePath;
        private System.Windows.Forms.Button buttonExcelToTables;
        private System.Windows.Forms.ListBox listBoxTables;
        private System.Windows.Forms.Button buttonTableToRows;
        private System.Windows.Forms.TreeView treeViewRows;
        private System.Windows.Forms.Button buttonWriteOutput;
    }
}

