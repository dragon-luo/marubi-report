namespace Marubi
{
    partial class MainForm
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows 窗体设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.TxtDataFile = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.TxtRptFile = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.BtnGenReport = new System.Windows.Forms.Button();
            this.lstStatus = new System.Windows.Forms.ListBox();
            this.BtnImport = new System.Windows.Forms.Button();
            this.BtnBrow = new System.Windows.Forms.Button();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // TxtDataFile
            // 
            this.TxtDataFile.BackColor = System.Drawing.SystemColors.Info;
            this.TxtDataFile.Enabled = false;
            this.TxtDataFile.Font = new System.Drawing.Font("SimSun", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.TxtDataFile.Location = new System.Drawing.Point(53, 119);
            this.TxtDataFile.Margin = new System.Windows.Forms.Padding(4);
            this.TxtDataFile.Name = "TxtDataFile";
            this.TxtDataFile.Size = new System.Drawing.Size(779, 30);
            this.TxtDataFile.TabIndex = 4;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.ForeColor = System.Drawing.Color.RoyalBlue;
            this.label1.Location = new System.Drawing.Point(51, 99);
            this.label1.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(112, 15);
            this.label1.TabIndex = 5;
            this.label1.Text = "数据文件路径：";
            // 
            // groupBox1
            // 
            this.groupBox1.BackColor = System.Drawing.SystemColors.Control;
            this.groupBox1.Controls.Add(this.TxtRptFile);
            this.groupBox1.Controls.Add(this.label3);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.BtnGenReport);
            this.groupBox1.Controls.Add(this.lstStatus);
            this.groupBox1.Controls.Add(this.BtnImport);
            this.groupBox1.Controls.Add(this.TxtDataFile);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Controls.Add(this.BtnBrow);
            this.groupBox1.Location = new System.Drawing.Point(31, 26);
            this.groupBox1.Margin = new System.Windows.Forms.Padding(4);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Padding = new System.Windows.Forms.Padding(4);
            this.groupBox1.Size = new System.Drawing.Size(899, 574);
            this.groupBox1.TabIndex = 6;
            this.groupBox1.TabStop = false;
            // 
            // TxtRptFile
            // 
            this.TxtRptFile.BackColor = System.Drawing.SystemColors.Info;
            this.TxtRptFile.Enabled = false;
            this.TxtRptFile.Font = new System.Drawing.Font("SimSun", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.TxtRptFile.Location = new System.Drawing.Point(53, 525);
            this.TxtRptFile.Margin = new System.Windows.Forms.Padding(4);
            this.TxtRptFile.Name = "TxtRptFile";
            this.TxtRptFile.Size = new System.Drawing.Size(779, 30);
            this.TxtRptFile.TabIndex = 10;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.ForeColor = System.Drawing.Color.RoyalBlue;
            this.label3.Location = new System.Drawing.Point(51, 505);
            this.label3.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(82, 15);
            this.label3.TabIndex = 11;
            this.label3.Text = "生成报表：";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.ForeColor = System.Drawing.Color.RoyalBlue;
            this.label2.Location = new System.Drawing.Point(51, 164);
            this.label2.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(82, 15);
            this.label2.TabIndex = 9;
            this.label2.Text = "处理状态：";
            // 
            // BtnGenReport
            // 
            this.BtnGenReport.Image = global::Marubi.Properties.Resources.gennerate_report;
            this.BtnGenReport.Location = new System.Drawing.Point(471, 25);
            this.BtnGenReport.Margin = new System.Windows.Forms.Padding(4);
            this.BtnGenReport.Name = "BtnGenReport";
            this.BtnGenReport.Size = new System.Drawing.Size(159, 55);
            this.BtnGenReport.TabIndex = 8;
            this.BtnGenReport.Text = "  生成报表";
            this.BtnGenReport.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.BtnGenReport.UseVisualStyleBackColor = false;
            this.BtnGenReport.Click += new System.EventHandler(this.BtnGenReport_Click);
            // 
            // lstStatus
            // 
            this.lstStatus.BackColor = System.Drawing.SystemColors.Info;
            this.lstStatus.ForeColor = System.Drawing.SystemColors.WindowText;
            this.lstStatus.FormattingEnabled = true;
            this.lstStatus.ItemHeight = 15;
            this.lstStatus.Location = new System.Drawing.Point(53, 184);
            this.lstStatus.Margin = new System.Windows.Forms.Padding(4);
            this.lstStatus.Name = "lstStatus";
            this.lstStatus.Size = new System.Drawing.Size(779, 304);
            this.lstStatus.TabIndex = 7;
            // 
            // BtnImport
            // 
            this.BtnImport.Image = global::Marubi.Properties.Resources.import;
            this.BtnImport.Location = new System.Drawing.Point(260, 25);
            this.BtnImport.Margin = new System.Windows.Forms.Padding(4);
            this.BtnImport.Name = "BtnImport";
            this.BtnImport.Size = new System.Drawing.Size(159, 55);
            this.BtnImport.TabIndex = 6;
            this.BtnImport.Text = "  数据导入";
            this.BtnImport.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.BtnImport.UseVisualStyleBackColor = false;
            this.BtnImport.Click += new System.EventHandler(this.BtnImport_Click);
            // 
            // BtnBrow
            // 
            this.BtnBrow.Image = global::Marubi.Properties.Resources.browser_search;
            this.BtnBrow.Location = new System.Drawing.Point(53, 25);
            this.BtnBrow.Margin = new System.Windows.Forms.Padding(4);
            this.BtnBrow.Name = "BtnBrow";
            this.BtnBrow.Size = new System.Drawing.Size(159, 55);
            this.BtnBrow.TabIndex = 3;
            this.BtnBrow.Text = " 浏览 ...";
            this.BtnBrow.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.BtnBrow.UseVisualStyleBackColor = false;
            this.BtnBrow.Click += new System.EventHandler(this.BtnBrow_Click);
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.Control;
            this.ClientSize = new System.Drawing.Size(957, 615);
            this.Controls.Add(this.groupBox1);
            this.Margin = new System.Windows.Forms.Padding(4);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "MainForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "丸美可用库存报表工具";
            this.Load += new System.EventHandler(this.MainForm_Load);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button BtnBrow;
        private System.Windows.Forms.TextBox TxtDataFile;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Button BtnImport;
        private System.Windows.Forms.ListBox lstStatus;
        private System.Windows.Forms.Button BtnGenReport;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox TxtRptFile;
        private System.Windows.Forms.Label label3;
    }
}

