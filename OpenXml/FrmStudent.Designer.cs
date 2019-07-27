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
            this.BtnExport = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.SuspendLayout();
            // 
            // dataGridView1
            // 
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Location = new System.Drawing.Point(12, 115);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.RowTemplate.Height = 27;
            this.dataGridView1.Size = new System.Drawing.Size(776, 323);
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
            // BtnExport
            // 
            this.BtnExport.Location = new System.Drawing.Point(180, 32);
            this.BtnExport.Name = "BtnExport";
            this.BtnExport.Size = new System.Drawing.Size(132, 41);
            this.BtnExport.TabIndex = 2;
            this.BtnExport.Text = "导出";
            this.BtnExport.UseVisualStyleBackColor = true;
            this.BtnExport.Click += new System.EventHandler(this.BtnExport_Click);
            // 
            // FrmStudent
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.BtnExport);
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
        private System.Windows.Forms.Button BtnExport;
    }
}