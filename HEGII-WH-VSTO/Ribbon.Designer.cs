namespace HEGII_WH_VSTO
{
    partial class Ribbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

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
            this.tab1 = this.Factory.CreateRibbonTab();
            this.tab2 = this.Factory.CreateRibbonTab();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.ButtonServeicOrderArrange = this.Factory.CreateRibbonButton();
            this.button1 = this.Factory.CreateRibbonButton();
            this.group3 = this.Factory.CreateRibbonGroup();
            this.buttonAddressCrawler = this.Factory.CreateRibbonButton();
            this.button3 = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.tab2.SuspendLayout();
            this.group2.SuspendLayout();
            this.group3.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
            // 
            // tab2
            // 
            this.tab2.Groups.Add(this.group2);
            this.tab2.Groups.Add(this.group3);
            this.tab2.Label = "恒洁客服";
            this.tab2.Name = "tab2";
            // 
            // group2
            // 
            this.group2.Items.Add(this.ButtonServeicOrderArrange);
            this.group2.Items.Add(this.button1);
            this.group2.Name = "group2";
            // 
            // ButtonServeicOrderArrange
            // 
            this.ButtonServeicOrderArrange.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.ButtonServeicOrderArrange.Label = "服务单整理";
            this.ButtonServeicOrderArrange.Name = "ButtonServeicOrderArrange";
            this.ButtonServeicOrderArrange.OfficeImageId = "FileSave";
            this.ButtonServeicOrderArrange.ShowImage = true;
            this.ButtonServeicOrderArrange.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ButtonServiceOrderArrange_Click);
            // 
            // button1
            // 
            this.button1.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button1.Label = "服务单整理";
            this.button1.Name = "button1";
            this.button1.OfficeImageId = "HighImportance";
            this.button1.ShowImage = true;
            // 
            // group3
            // 
            this.group3.Items.Add(this.buttonAddressCrawler);
            this.group3.Items.Add(this.button3);
            this.group3.Label = "group3";
            this.group3.Name = "group3";
            // 
            // buttonAddressCrawler
            // 
            this.buttonAddressCrawler.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.buttonAddressCrawler.Image = global::HEGII_WH_VSTO.Properties.Resources.icons8_老蜘蛛侠_80;
            this.buttonAddressCrawler.Label = "地址爬虫";
            this.buttonAddressCrawler.Name = "buttonAddressCrawler";
            this.buttonAddressCrawler.OfficeImageId = "HighImportance";
            this.buttonAddressCrawler.ShowImage = true;
            this.buttonAddressCrawler.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonAddressCrawler_Click);
            // 
            // button3
            // 
            this.button3.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button3.Label = "服务单整理";
            this.button3.Name = "button3";
            this.button3.OfficeImageId = "HighImportance";
            this.button3.ShowImage = true;
            // 
            // Ribbon
            // 
            this.Name = "Ribbon";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Tabs.Add(this.tab2);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.tab2.ResumeLayout(false);
            this.tab2.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.group3.ResumeLayout(false);
            this.group3.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        private Microsoft.Office.Tools.Ribbon.RibbonTab tab2;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ButtonServeicOrderArrange;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group3;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonAddressCrawler;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button3;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon Ribbon
        {
            get { return this.GetRibbon<Ribbon>(); }
        }
    }
}
