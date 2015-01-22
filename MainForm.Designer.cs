namespace NGWordSplitter
{
    partial class MainForm
    {
        /// <summary>
        /// 設計工具所需的變數。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清除任何使用中的資源。
        /// </summary>
        /// <param name="disposing">如果應該處置 Managed 資源則為 true，否則為 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form 設計工具產生的程式碼

        /// <summary>
        /// 此為設計工具支援所需的方法 - 請勿使用程式碼編輯器
        /// 修改這個方法的內容。
        /// </summary>
        private void InitializeComponent()
        {
            this.btnSplite = new System.Windows.Forms.Button();
            this.txtPattern = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.txtDocFN = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // btnSplite
            // 
            this.btnSplite.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.btnSplite.Location = new System.Drawing.Point(132, 126);
            this.btnSplite.Name = "btnSplite";
            this.btnSplite.Size = new System.Drawing.Size(140, 23);
            this.btnSplite.TabIndex = 0;
            this.btnSplite.Text = "分割";
            this.btnSplite.UseVisualStyleBackColor = true;
            this.btnSplite.Click += new System.EventHandler(this.btnSplite_Click);
            // 
            // txtPattern
            // 
            this.txtPattern.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtPattern.Location = new System.Drawing.Point(12, 98);
            this.txtPattern.Name = "txtPattern";
            this.txtPattern.Size = new System.Drawing.Size(260, 22);
            this.txtPattern.TabIndex = 2;
            this.txtPattern.Text = "a?(\\w*)\\s*座號：(\\d*)\\s*姓名：(\\w*)";
            // 
            // label1
            // 
            this.label1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(13, 82);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(53, 12);
            this.label1.TabIndex = 3;
            this.label1.Text = "檔名模式";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(13, 9);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(71, 12);
            this.label2.TabIndex = 3;
            this.label2.Text = "Word 檔案：";
            // 
            // txtDocFN
            // 
            this.txtDocFN.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtDocFN.DataBindings.Add(new System.Windows.Forms.Binding("Text", global::NGWordSplitter.Properties.Settings.Default, "DocFileName", true, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged));
            this.txtDocFN.Location = new System.Drawing.Point(15, 25);
            this.txtDocFN.Name = "txtDocFN";
            this.txtDocFN.Size = new System.Drawing.Size(257, 22);
            this.txtDocFN.TabIndex = 4;
            this.txtDocFN.Text = global::NGWordSplitter.Properties.Settings.Default.DocFileName;
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(284, 161);
            this.Controls.Add(this.txtDocFN);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.txtPattern);
            this.Controls.Add(this.btnSplite);
            this.Name = "MainForm";
            this.Text = "分割";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.MainForm_FormClosing);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnSplite;
        private System.Windows.Forms.TextBox txtPattern;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txtDocFN;
        private System.Windows.Forms.Label label2;
    }
}

