namespace HEGII_WH_VSTO
{
    partial class ctpOrderImport
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

        #region 组件设计器生成的代码

        /// <summary> 
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.labelWorkSheet = new System.Windows.Forms.Label();
            this.comboOldOrderType = new System.Windows.Forms.ComboBox();
            this.labelOrderType = new System.Windows.Forms.Label();
            this.buttonOldDataOK = new System.Windows.Forms.Button();
            this.labelWorkBook = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // labelWorkSheet
            // 
            this.labelWorkSheet.AutoSize = true;
            this.labelWorkSheet.Location = new System.Drawing.Point(5, 28);
            this.labelWorkSheet.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.labelWorkSheet.Name = "labelWorkSheet";
            this.labelWorkSheet.Size = new System.Drawing.Size(45, 19);
            this.labelWorkSheet.TabIndex = 16;
            this.labelWorkSheet.Text = "label1";
            // 
            // comboOldOrderType
            // 
            this.comboOldOrderType.FormattingEnabled = true;
            this.comboOldOrderType.Location = new System.Drawing.Point(74, 54);
            this.comboOldOrderType.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.comboOldOrderType.Name = "comboOldOrderType";
            this.comboOldOrderType.Size = new System.Drawing.Size(118, 27);
            this.comboOldOrderType.TabIndex = 15;
            // 
            // labelOrderType
            // 
            this.labelOrderType.AutoSize = true;
            this.labelOrderType.Location = new System.Drawing.Point(5, 57);
            this.labelOrderType.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.labelOrderType.Name = "labelOrderType";
            this.labelOrderType.Size = new System.Drawing.Size(61, 19);
            this.labelOrderType.TabIndex = 14;
            this.labelOrderType.Text = "订单类型";
            // 
            // buttonOldDataOK
            // 
            this.buttonOldDataOK.Location = new System.Drawing.Point(112, 91);
            this.buttonOldDataOK.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.buttonOldDataOK.Name = "buttonOldDataOK";
            this.buttonOldDataOK.Size = new System.Drawing.Size(80, 30);
            this.buttonOldDataOK.TabIndex = 13;
            this.buttonOldDataOK.Text = "确认";
            this.buttonOldDataOK.UseVisualStyleBackColor = true;
            this.buttonOldDataOK.Click += new System.EventHandler(this.buttonOldDataOK_Click);
            // 
            // labelWorkBook
            // 
            this.labelWorkBook.AutoSize = true;
            this.labelWorkBook.Location = new System.Drawing.Point(5, 9);
            this.labelWorkBook.Margin = new System.Windows.Forms.Padding(5, 0, 5, 0);
            this.labelWorkBook.Name = "labelWorkBook";
            this.labelWorkBook.Size = new System.Drawing.Size(45, 19);
            this.labelWorkBook.TabIndex = 12;
            this.labelWorkBook.Text = "label1";
            // 
            // ctpOrderImport
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 19F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.labelWorkSheet);
            this.Controls.Add(this.comboOldOrderType);
            this.Controls.Add(this.labelOrderType);
            this.Controls.Add(this.buttonOldDataOK);
            this.Controls.Add(this.labelWorkBook);
            this.Font = new System.Drawing.Font("微软雅黑", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.Name = "ctpOrderImport";
            this.Size = new System.Drawing.Size(196, 130);
            this.Load += new System.EventHandler(this.ctpOrderImport_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label labelWorkSheet;
        private System.Windows.Forms.ComboBox comboOldOrderType;
        private System.Windows.Forms.Label labelOrderType;
        private System.Windows.Forms.Button buttonOldDataOK;
        private System.Windows.Forms.Label labelWorkBook;
    }
}
