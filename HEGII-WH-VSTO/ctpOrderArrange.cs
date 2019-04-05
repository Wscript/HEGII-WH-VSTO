using Microsoft.Office.Interop.Excel;
using System;
using System.Windows.Forms;

namespace HEGII_WH_VSTO
{
    public partial class ctpOrderArrange : UserControl
    {
        public ctpOrderArrange()
        {
            InitializeComponent();
        }

        private void buttonArrangeStart_Click(object sender, EventArgs e)
        {
            string[] TitleList = { "制订日期", "安装日期", "单号", "部门", "联系人", "联系电话",
                                   "大区", "送货地点", "安装备注", "商品名称", "安装数量" };
            string[] TitleList_Survey = {"单号","部门","报单日期","需要的测量时间","联系人",
                                         "所属路线片区","导购员","送货地点","联系电话","商品名称","备注"};
            if (CheckFileFormat(TitleList_Survey, 6))
            {
                //MessageBox.Show("FormatOK");

            }
            else
            {
                MessageBox.Show("当前Sheet格式错误，请确认需要整理的单据在当前页面！");
            }
        }

        private bool CheckFileFormat(string[] TitleList,int TitleRow)
        {
            Worksheet ActiveSheet = Globals.ThisAddIn.Application.ActiveSheet;
            bool boolNotFind = true;
            foreach (string TitleItem in TitleList)
            {
                Range rangeFindItem = ActiveSheet.Range[TitleRow.ToString() + ":" + TitleRow.ToString()].Find(TitleItem);
                if (rangeFindItem == null)
                {
                    boolNotFind = false;
                }
            }
            return (boolNotFind);
        }
    }
}
