namespace HEGII_WH_VSTO
{
    partial class ctpOrderArrange
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
            this.buttonArrangeStart = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // buttonArrangeStart
            // 
            this.buttonArrangeStart.Location = new System.Drawing.Point(50, 32);
            this.buttonArrangeStart.Name = "buttonArrangeStart";
            this.buttonArrangeStart.Size = new System.Drawing.Size(75, 30);
            this.buttonArrangeStart.TabIndex = 0;
            this.buttonArrangeStart.Text = "开始整理";
            this.buttonArrangeStart.UseVisualStyleBackColor = true;
            this.buttonArrangeStart.Click += new System.EventHandler(this.buttonArrangeStart_Click);
            // 
            // ctpOrderArrange
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 19F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.buttonArrangeStart);
            this.Font = new System.Drawing.Font("微软雅黑", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.Name = "ctpOrderArrange";
            this.Size = new System.Drawing.Size(171, 98);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button buttonArrangeStart;
    }
}
