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
            this.groupUserInfo = this.Factory.CreateRibbonGroup();
            this.labelUser = this.Factory.CreateRibbonLabel();
            this.labelUserName = this.Factory.CreateRibbonLabel();
            this.group3 = this.Factory.CreateRibbonGroup();
            this.buttonUserLogin = this.Factory.CreateRibbonButton();
            this.buttonAddressCrawler = this.Factory.CreateRibbonButton();
            this.ButtonOrderArrange = this.Factory.CreateRibbonButton();
            this.buttonOrderImport = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.tab2.SuspendLayout();
            this.groupUserInfo.SuspendLayout();
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
            this.tab2.Groups.Add(this.groupUserInfo);
            this.tab2.Groups.Add(this.group3);
            this.tab2.Label = "恒洁客服";
            this.tab2.Name = "tab2";
            // 
            // groupUserInfo
            // 
            this.groupUserInfo.Items.Add(this.buttonUserLogin);
            this.groupUserInfo.Items.Add(this.labelUser);
            this.groupUserInfo.Items.Add(this.labelUserName);
            this.groupUserInfo.Label = "用户信息";
            this.groupUserInfo.Name = "groupUserInfo";
            // 
            // labelUser
            // 
            this.labelUser.Label = "用户名";
            this.labelUser.Name = "labelUser";
            // 
            // labelUserName
            // 
            this.labelUserName.Label = "<未登录>";
            this.labelUserName.Name = "labelUserName";
            // 
            // group3
            // 
            this.group3.Items.Add(this.buttonAddressCrawler);
            this.group3.Items.Add(this.ButtonOrderArrange);
            this.group3.Items.Add(this.buttonOrderImport);
            this.group3.Name = "group3";
            // 
            // buttonUserLogin
            // 
            this.buttonUserLogin.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.buttonUserLogin.Image = global::HEGII_WH_VSTO.Properties.Resources.login;
            this.buttonUserLogin.Label = "用户登录";
            this.buttonUserLogin.Name = "buttonUserLogin";
            this.buttonUserLogin.OfficeImageId = "HighImportance";
            this.buttonUserLogin.ShowImage = true;
            this.buttonUserLogin.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonUserLogin_Click);
            // 
            // buttonAddressCrawler
            // 
            this.buttonAddressCrawler.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.buttonAddressCrawler.Image = global::HEGII_WH_VSTO.Properties.Resources.MapMarker;
            this.buttonAddressCrawler.Label = "地址爬虫";
            this.buttonAddressCrawler.Name = "buttonAddressCrawler";
            this.buttonAddressCrawler.OfficeImageId = "HighImportance";
            this.buttonAddressCrawler.ShowImage = true;
            this.buttonAddressCrawler.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonAddressCrawler_Click);
            // 
            // ButtonOrderArrange
            // 
            this.ButtonOrderArrange.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.ButtonOrderArrange.Image = global::HEGII_WH_VSTO.Properties.Resources.sort;
            this.ButtonOrderArrange.Label = "服务单整理";
            this.ButtonOrderArrange.Name = "ButtonOrderArrange";
            this.ButtonOrderArrange.OfficeImageId = "FileSave";
            this.ButtonOrderArrange.ShowImage = true;
            this.ButtonOrderArrange.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ButtonOrderArrange_Click);
            // 
            // buttonOrderImport
            // 
            this.buttonOrderImport.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.buttonOrderImport.Image = global::HEGII_WH_VSTO.Properties.Resources.database_plus;
            this.buttonOrderImport.Label = "服务单导入";
            this.buttonOrderImport.Name = "buttonOrderImport";
            this.buttonOrderImport.ShowImage = true;
            this.buttonOrderImport.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonOrderImport_Click);
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
            this.groupUserInfo.ResumeLayout(false);
            this.groupUserInfo.PerformLayout();
            this.group3.ResumeLayout(false);
            this.group3.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        private Microsoft.Office.Tools.Ribbon.RibbonTab tab2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ButtonOrderArrange;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group3;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonUserLogin;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonAddressCrawler;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel labelUserName;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel labelUser;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupUserInfo;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonOrderImport;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon Ribbon
        {
            get { return this.GetRibbon<Ribbon>(); }
        }
    }
}
