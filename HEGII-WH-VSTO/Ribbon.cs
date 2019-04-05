using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Ribbon;
using System.IO;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace HEGII_WH_VSTO
{
    public partial class Ribbon
    {
        internal Microsoft.Office.Tools.CustomTaskPane ctpAddressCrawler;
        internal Microsoft.Office.Tools.CustomTaskPane ctpUserLogin;
        internal Microsoft.Office.Tools.CustomTaskPane ctpOrderImport;
        internal Microsoft.Office.Tools.CustomTaskPane ctpOrderArrange;

        private void Ribbon_Load(object sender, RibbonUIEventArgs e)
        {
            groupUserInfo.Label = System.Net.Dns.GetHostName().ToString();      //获取当前计算机名
        }

        private void ButtonCommissionArrange_Click(object sender, RibbonControlEventArgs e)
        {
            Workbook ActiveWorkBook = Globals.ThisAddIn.Application.ActiveWorkbook;
            bool FileVerify = false;
            foreach (Worksheet WorkSheet in ActiveWorkBook.Worksheets)
            {
                if (WorkSheet.Name == "汇总")
                {
                    FileVerify = true;
                }
            }
            if (FileVerify)
            {
                FolderBrowserDialog dialog = new FolderBrowserDialog();
                dialog.Description = "请选择文件路径";

                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    string foldPath = dialog.SelectedPath;
                    DirectoryInfo theFolder = new DirectoryInfo(foldPath);
                    FileInfo[] dirInfo = theFolder.GetFiles();
                    foreach (FileInfo file in dirInfo)
                    {
                        //MessageBox.Show(foldPath + file.ToString());
                        Excel.Application NewEXCELFile = new Excel.Application();
                        Workbook NewWorkbook = NewEXCELFile.Application.Workbooks.Open(foldPath + "\\" + file.ToString());
                        foreach (Worksheet NewWorkSheet in NewWorkbook.Worksheets)
                        {
                            foreach (Worksheet ActiveWorkSheet in ActiveWorkBook.Worksheets)
                            {
                                if (ActiveWorkSheet.Name == NewWorkSheet.Name)
                                {
                                    int i = 5;
                                    while (NewWorkSheet.Cells[i,11] != "")
                                    {
                                        MessageBox.Show(ActiveWorkSheet.UsedRange.Row.ToString());
                                    }

                                }
                            }
                        }
                    }
                }
            }
            else
            {
                MessageBox.Show("文件格式不正确！");
            }
        }

        private void ButtonOrderArrange_Click(object sender, RibbonControlEventArgs e)
        {
//            if (Globals.Ribbons.Ribbon.labelUserName.Label == "<未登录>")
   //         {
 //               MessageBox.Show("未登录！");
  //          }
 //           else
 //           {
                if (ctpOrderArrange == null)
                {
                    ctpOrderArrange = Globals.ThisAddIn.CustomTaskPanes.Add(new ctpOrderArrange(), "服务单整理");
                }
                CustomTaskPaneVisible(ctpOrderArrange);
 //           }
        }

        private void buttonAddressCrawler_Click(object sender, RibbonControlEventArgs e)
        {
            if (Globals.Ribbons.Ribbon.labelUserName.Label == "<未登录>")
            {
                MessageBox.Show("未登录！");
            }
            else
            {
                if (ctpAddressCrawler == null)
                {
                    ctpAddressCrawler = Globals.ThisAddIn.CustomTaskPanes.Add(new ctpAddressCrawler(), "起始地址");
                }
                CustomTaskPaneVisible(ctpAddressCrawler);
            }
        }

        private void buttonUserLogin_Click(object sender, RibbonControlEventArgs e)
        {
            if (Globals.Ribbons.Ribbon.buttonUserLogin.Label == "用户登录")
            {
                if (ctpUserLogin == null)
                {
                    ctpUserLogin = Globals.ThisAddIn.CustomTaskPanes.Add(new ctpUserLogin(), "用户登录");
                }
                CustomTaskPaneVisible(ctpUserLogin);
            }
            else
            {
                Globals.Ribbons.Ribbon.buttonUserLogin.Label = "用户登录";
                Globals.Ribbons.Ribbon.labelUserName.Label = "<未登录>";
            }
        }

        private void CustomTaskPaneVisible (Microsoft.Office.Tools.CustomTaskPane CustomTaskPane)
        {
            if (CustomTaskPane.Visible == true)
            {
                CustomTaskPane.Visible = false;
                CustomTaskPane = null;
            }
            else
            {
                CustomTaskPane.Visible = true;
            }
        }

        private void buttonOrderImport_Click(object sender, RibbonControlEventArgs e)
        {
            if (Globals.Ribbons.Ribbon.labelUserName.Label == "<未登录>")
            {
                MessageBox.Show("未登录！");
            }
            else
            {
                if (ctpOrderImport == null)
                {
                    ctpOrderImport = Globals.ThisAddIn.CustomTaskPanes.Add(new ctpOrderImport(), "服务单导入");
                }
                CustomTaskPaneVisible(ctpOrderImport);
            }
        }
    }
}
