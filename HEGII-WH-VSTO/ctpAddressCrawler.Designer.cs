namespace HEGII_WH_VSTO
{
    partial class ctpAddressCrawler
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
            this.labelStartUrl = new System.Windows.Forms.Label();
            this.textStartUrl = new System.Windows.Forms.TextBox();
            this.buttonStart = new System.Windows.Forms.Button();
            this.labelMemo = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // labelStartUrl
            // 
            this.labelStartUrl.AutoSize = true;
            this.labelStartUrl.Font = new System.Drawing.Font("微软雅黑", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.labelStartUrl.Location = new System.Drawing.Point(4, 4);
            this.labelStartUrl.Name = "labelStartUrl";
            this.labelStartUrl.Size = new System.Drawing.Size(126, 19);
            this.labelStartUrl.TabIndex = 0;
            this.labelStartUrl.Text = "爬虫起始链接地址：";
            // 
            // textStartUrl
            // 
            this.textStartUrl.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.textStartUrl.Font = new System.Drawing.Font("微软雅黑", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.textStartUrl.Location = new System.Drawing.Point(4, 29);
            this.textStartUrl.Name = "textStartUrl";
            this.textStartUrl.Size = new System.Drawing.Size(191, 25);
            this.textStartUrl.TabIndex = 1;
            // 
            // buttonStart
            // 
            this.buttonStart.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.buttonStart.Font = new System.Drawing.Font("微软雅黑", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.buttonStart.Location = new System.Drawing.Point(120, 64);
            this.buttonStart.Name = "buttonStart";
            this.buttonStart.Size = new System.Drawing.Size(75, 30);
            this.buttonStart.TabIndex = 2;
            this.buttonStart.Text = "开始";
            this.buttonStart.UseVisualStyleBackColor = true;
            this.buttonStart.Click += new System.EventHandler(this.buttonStart_Click);
            // 
            // labelMemo
            // 
            this.labelMemo.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.labelMemo.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.labelMemo.Font = new System.Drawing.Font("微软雅黑", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.labelMemo.Location = new System.Drawing.Point(0, 134);
            this.labelMemo.Name = "labelMemo";
            this.labelMemo.Size = new System.Drawing.Size(198, 123);
            this.labelMemo.TabIndex = 3;
            this.labelMemo.Text = "进入链家首页，选择城市，复制页面地址\r\n\r\n例如：https://wh.lianjia.com\r\n\r\n请注意，地址不要以“/”结尾";
            // 
            // ctpAddressCrawler
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.labelMemo);
            this.Controls.Add(this.buttonStart);
            this.Controls.Add(this.textStartUrl);
            this.Controls.Add(this.labelStartUrl);
            this.Name = "ctpAddressCrawler";
            this.Size = new System.Drawing.Size(198, 257);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label labelStartUrl;
        private System.Windows.Forms.TextBox textStartUrl;
        private System.Windows.Forms.Button buttonStart;
        private System.Windows.Forms.Label labelMemo;
    }
}
