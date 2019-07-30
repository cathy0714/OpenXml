namespace OpenXml
{
    partial class FrmStudent
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.BtnQuery = new System.Windows.Forms.Button();
            this.BtnOpenXmlExport = new System.Windows.Forms.Button();
            this.BtnOpenXmlImport = new System.Windows.Forms.Button();
            this.BtnNpoiExport = new System.Windows.Forms.Button();
            this.BtnNpoiImport = new System.Windows.Forms.Button();
            this.BtnExportWord = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.SuspendLayout();
            // 
            // dataGridView1
            // 
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.GridColor = System.Drawing.SystemColors.Highlight;
            this.dataGridView1.Location = new System.Drawing.Point(12, 115);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.RowTemplate.Height = 27;
            this.dataGridView1.Size = new System.Drawing.Size(895, 424);
            this.dataGridView1.TabIndex = 0;
            // 
            // BtnQuery
            // 
            this.BtnQuery.Location = new System.Drawing.Point(25, 32);
            this.BtnQuery.Name = "BtnQuery";
            this.BtnQuery.Size = new System.Drawing.Size(132, 41);
            this.BtnQuery.TabIndex = 1;
            this.BtnQuery.Text = "查询";
            this.BtnQuery.UseVisualStyleBackColor = true;
            this.BtnQuery.Click += new System.EventHandler(this.BtnQuery_Click);
            // 
            // BtnOpenXmlExport
            // 
            this.BtnOpenXmlExport.Location = new System.Drawing.Point(178, 12);
            this.BtnOpenXmlExport.Name = "BtnOpenXmlExport";
            this.BtnOpenXmlExport.Size = new System.Drawing.Size(130, 31);
            this.BtnOpenXmlExport.TabIndex = 2;
            this.BtnOpenXmlExport.Text = "OpenXml导出";
            this.BtnOpenXmlExport.UseVisualStyleBackColor = true;
            this.BtnOpenXmlExport.Click += new System.EventHandler(this.BtnOpenXmlExport_Click);
            // 
            // BtnOpenXmlImport
            // 
            this.BtnOpenXmlImport.Location = new System.Drawing.Point(327, 12);
            this.BtnOpenXmlImport.Name = "BtnOpenXmlImport";
            this.BtnOpenXmlImport.Size = new System.Drawing.Size(133, 31);
            this.BtnOpenXmlImport.TabIndex = 3;
            this.BtnOpenXmlImport.Text = "OpenXml导入";
            this.BtnOpenXmlImport.UseVisualStyleBackColor = true;
            this.BtnOpenXmlImport.Click += new System.EventHandler(this.BtnOpenXmlImport_Click);
            // 
            // BtnNpoiExport
            // 
            this.BtnNpoiExport.Location = new System.Drawing.Point(178, 49);
            this.BtnNpoiExport.Name = "BtnNpoiExport";
            this.BtnNpoiExport.Size = new System.Drawing.Size(130, 31);
            this.BtnNpoiExport.TabIndex = 4;
            this.BtnNpoiExport.Text = "Npoi导出";
            this.BtnNpoiExport.UseVisualStyleBackColor = true;
            this.BtnNpoiExport.Click += new System.EventHandler(this.BtnNpoiExport_Click);
            // 
            // BtnNpoiImport
            // 
            this.BtnNpoiImport.Location = new System.Drawing.Point(327, 49);
            this.BtnNpoiImport.Name = "BtnNpoiImport";
            this.BtnNpoiImport.Size = new System.Drawing.Size(133, 31);
            this.BtnNpoiImport.TabIndex = 5;
            this.BtnNpoiImport.Text = "Npoi导入";
            this.BtnNpoiImport.UseVisualStyleBackColor = true;
            this.BtnNpoiImport.Click += new System.EventHandler(this.BtnNpoiImport_Click);
            // 
            // BtnExportWord
            // 
            this.BtnExportWord.Location = new System.Drawing.Point(489, 23);
            this.BtnExportWord.Name = "BtnExportWord";
            this.BtnExportWord.Size = new System.Drawing.Size(132, 41);
            this.BtnExportWord.TabIndex = 6;
            this.BtnExportWord.Text = "导出文档";
            this.BtnExportWord.UseVisualStyleBackColor = true;
            this.BtnExportWord.Click += new System.EventHandler(this.BtnExportWord_Click);
            // 
            // FrmStudent
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(919, 551);
            this.Controls.Add(this.BtnExportWord);
            this.Controls.Add(this.BtnNpoiImport);
            this.Controls.Add(this.BtnNpoiExport);
            this.Controls.Add(this.BtnOpenXmlImport);
            this.Controls.Add(this.BtnOpenXmlExport);
            this.Controls.Add(this.BtnQuery);
            this.Controls.Add(this.dataGridView1);
            this.Name = "FrmStudent";
            this.Text = "FrmStudent";
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.Button BtnQuery;
        private System.Windows.Forms.Button BtnOpenXmlExport;
        private System.Windows.Forms.Button BtnOpenXmlImport;
        private System.Windows.Forms.Button BtnNpoiExport;
        private System.Windows.Forms.Button BtnNpoiImport;
        private System.Windows.Forms.Button BtnExportWord;
    }
}