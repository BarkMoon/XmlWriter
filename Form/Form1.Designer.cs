namespace XmlWriter
{
    partial class Form1
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
            this.lblFilePath = new System.Windows.Forms.Label();
            this.txtFilePath = new System.Windows.Forms.TextBox();
            this.btnBrowse = new System.Windows.Forms.Button();
            this.lblSheetName = new System.Windows.Forms.Label();
            this.cmbSheetName = new System.Windows.Forms.ComboBox();
            this.btnGenerate = new System.Windows.Forms.Button();
            this.lblStatus = new System.Windows.Forms.Label();

            // ★追加: テンプレート選択用
            this.lblTemplate = new System.Windows.Forms.Label();
            this.txtTemplatePath = new System.Windows.Forms.TextBox();
            this.btnBrowseTemplate = new System.Windows.Forms.Button();

            this.btnGenerateClass = new System.Windows.Forms.Button();
            // Start Add
            this.btnGenerateScriptFromData = new System.Windows.Forms.Button();
            // End Add
            this.SuspendLayout();
            // 
            // lblFilePath
            // 
            this.lblFilePath.AutoSize = true;
            this.lblFilePath.Location = new System.Drawing.Point(24, 24);
            this.lblFilePath.Name = "lblFilePath";
            this.lblFilePath.Size = new System.Drawing.Size(101, 12);
            this.lblFilePath.TabIndex = 0;
            this.lblFilePath.Text = "Excelファイルパス:";
            // 
            // txtFilePath
            // 
            this.txtFilePath.Location = new System.Drawing.Point(26, 48);
            this.txtFilePath.Name = "txtFilePath";
            this.txtFilePath.ReadOnly = true;
            this.txtFilePath.Size = new System.Drawing.Size(350, 19);
            this.txtFilePath.TabIndex = 1;
            // 
            // btnBrowse
            // 
            this.btnBrowse.Location = new System.Drawing.Point(390, 46);
            this.btnBrowse.Name = "btnBrowse";
            this.btnBrowse.Size = new System.Drawing.Size(75, 23);
            this.btnBrowse.TabIndex = 2;
            this.btnBrowse.Text = "参照...";
            this.btnBrowse.UseVisualStyleBackColor = true;
            this.btnBrowse.Click += new System.EventHandler(this.btnBrowse_Click);
            // 
            // lblSheetName
            // 
            this.lblSheetName.AutoSize = true;
            this.lblSheetName.Location = new System.Drawing.Point(24, 90);
            this.lblSheetName.Name = "lblSheetName";
            this.lblSheetName.Size = new System.Drawing.Size(126, 12);
            this.lblSheetName.TabIndex = 3;
            this.lblSheetName.Text = "データテーブル名を選択:";
            // 
            // cmbSheetName
            // 
            this.cmbSheetName.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbSheetName.FormattingEnabled = true;
            this.cmbSheetName.Location = new System.Drawing.Point(26, 115);
            this.cmbSheetName.Name = "cmbSheetName";
            this.cmbSheetName.Size = new System.Drawing.Size(200, 20);
            this.cmbSheetName.TabIndex = 4;
            // 
            // btnGenerate
            // 
            this.btnGenerate.Location = new System.Drawing.Point(26, 150);
            this.btnGenerate.Name = "btnGenerate";
            this.btnGenerate.Size = new System.Drawing.Size(439, 40);
            this.btnGenerate.TabIndex = 5;
            this.btnGenerate.Text = "XML生成実行";
            this.btnGenerate.UseVisualStyleBackColor = true;
            this.btnGenerate.Click += new System.EventHandler(this.btnGenerate_Click);
            // 
            // lblStatus
            // 
            this.lblStatus.AutoSize = true;
            this.lblStatus.ForeColor = System.Drawing.Color.Gray;
            this.lblStatus.Location = new System.Drawing.Point(26, 350); // 位置調整
            this.lblStatus.Name = "lblStatus";
            this.lblStatus.Size = new System.Drawing.Size(53, 12);
            this.lblStatus.TabIndex = 9; // Index調整
            this.lblStatus.Text = "準備完了";
            // 
            // lblTemplate
            // 
            this.lblTemplate.AutoSize = true;
            this.lblTemplate.Location = new System.Drawing.Point(24, 210);
            this.lblTemplate.Name = "lblTemplate";
            this.lblTemplate.Size = new System.Drawing.Size(177, 12);
            this.lblTemplate.TabIndex = 6;
            this.lblTemplate.Text = "テンプレートファイル (未指定でデフォルト):";
            // 
            // txtTemplatePath
            // 
            this.txtTemplatePath.Location = new System.Drawing.Point(26, 230);
            this.txtTemplatePath.Name = "txtTemplatePath";
            this.txtTemplatePath.ReadOnly = true;
            this.txtTemplatePath.Size = new System.Drawing.Size(350, 19);
            this.txtTemplatePath.TabIndex = 7;
            // 
            // btnBrowseTemplate
            // 
            this.btnBrowseTemplate.Location = new System.Drawing.Point(390, 228);
            this.btnBrowseTemplate.Name = "btnBrowseTemplate";
            this.btnBrowseTemplate.Size = new System.Drawing.Size(75, 23);
            this.btnBrowseTemplate.TabIndex = 8;
            this.btnBrowseTemplate.Text = "参照...";
            this.btnBrowseTemplate.UseVisualStyleBackColor = true;
            this.btnBrowseTemplate.Click += new System.EventHandler(this.btnBrowseTemplate_Click);
            // 
            // btnGenerateClass
            // 
            this.btnGenerateClass.Location = new System.Drawing.Point(26, 270);
            this.btnGenerateClass.Name = "btnGenerateClass";
            this.btnGenerateClass.Size = new System.Drawing.Size(439, 40);
            this.btnGenerateClass.TabIndex = 9;
            this.btnGenerateClass.Text = "C#クラス定義の生成 (.cs)";
            this.btnGenerateClass.UseVisualStyleBackColor = true;
            this.btnGenerateClass.Click += new System.EventHandler(this.btnGenerateClass_Click);
            // 
            // btnGenerateScriptFromData
            // 
            this.btnGenerateScriptFromData.Location = new System.Drawing.Point(26, 310); // Offset Y
            this.btnGenerateScriptFromData.Name = "btnGenerateScriptFromData";
            this.btnGenerateScriptFromData.Size = new System.Drawing.Size(439, 40);
            this.btnGenerateScriptFromData.TabIndex = 10;
            this.btnGenerateScriptFromData.Text = "データからスクリプトを作成";
            this.btnGenerateScriptFromData.UseVisualStyleBackColor = true;
            this.btnGenerateScriptFromData.Click += new System.EventHandler(this.btnGenerateScriptFromData_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(500, 380); // 高さを拡張
            this.Controls.Add(this.btnGenerateScriptFromData); // Add this
            this.Controls.Add(this.btnGenerateClass);
            this.Controls.Add(this.btnBrowseTemplate);
            this.Controls.Add(this.txtTemplatePath);
            this.Controls.Add(this.lblTemplate);
            this.Controls.Add(this.lblStatus);
            this.Controls.Add(this.btnGenerate);
            this.Controls.Add(this.cmbSheetName);
            this.Controls.Add(this.lblSheetName);
            this.Controls.Add(this.btnBrowse);
            this.Controls.Add(this.txtFilePath);
            this.Controls.Add(this.lblFilePath);
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Excel to XML Converter (XmlWriter)";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label lblFilePath;
        private System.Windows.Forms.TextBox txtFilePath;
        private System.Windows.Forms.Button btnBrowse;
        private System.Windows.Forms.Label lblSheetName;
        private System.Windows.Forms.ComboBox cmbSheetName;
        private System.Windows.Forms.Button btnGenerate;
        private System.Windows.Forms.Label lblStatus;
        // 新規追加
        private System.Windows.Forms.Label lblTemplate;
        private System.Windows.Forms.TextBox txtTemplatePath;
        private System.Windows.Forms.Button btnBrowseTemplate;
        private System.Windows.Forms.Button btnGenerateClass;
        // 新規追加
        private System.Windows.Forms.Button btnGenerateScriptFromData;
    }
}