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

        private void Ribbon_Load(object sender, RibbonUIEventArgs e)
        {
            labelComputerName.Label = System.Net.Dns.GetHostName().ToString();      //获取当前计算机名
        }

        private bool CheckFileFormat()
        {
            string StringRangeRow;

            Worksheet ActiveSheet = Globals.ThisAddIn.Application.ActiveSheet;
            
            if (ActiveSheet.Range["A:A"].Find("日期") == null)
            {
                return (false);
            }
            else
            {
                StringRangeRow = ActiveSheet.Range["A:A"].Find("日期").Row + ":" + ActiveSheet.Range["A:A"].Find("日期").Row;
            }

            if (ActiveSheet.Range[StringRangeRow].Find("日期") == null)
            {
                return (false);
            }
            if (ActiveSheet.Range[StringRangeRow].Find("单号") == null)
            {
                return (false);
            }
            if (ActiveSheet.Range[StringRangeRow].Find("部门") == null)
            {
                return (false);
            }
            if (ActiveSheet.Range[StringRangeRow].Find("联系人") == null)
            {
                return (false);
            }
            if (ActiveSheet.Range[StringRangeRow].Find("联系电话") == null)
            {
                return (false);
            }
            if (ActiveSheet.Range[StringRangeRow].Find("送货地点") == null)
            {
                return (false);
            }
            if (ActiveSheet.Range[StringRangeRow].Find("安装备注") == null)
            {
                return (false);
            }
            if (ActiveSheet.Range[StringRangeRow].Find("商品名称") == null)
            {
                return (false);
            }
            if (ActiveSheet.Range[StringRangeRow].Find("安装数量") == null)
            {
                return (false);
            }
            return (true);
        }

        private void InstallOrderArrange (Workbook NewWorkbook, Worksheet NewWorksheet)
        {
            NewWorksheet.Cells[1, 2] = "装/修";
            NewWorksheet.Cells[1, 3] = "销售点";
            NewWorksheet.Cells[1, 4] = "报装日期";
            NewWorksheet.Cells[1, 5] = "预约日期";
            NewWorksheet.Cells[1, 6] = "用户名称";
            NewWorksheet.Cells[1, 7] = "大范围";
            NewWorksheet.Cells[1, 8] = "销售人员";
            NewWorksheet.Cells[1, 9] = "地址";
            NewWorksheet.Cells[1, 10] = "电话";
            NewWorksheet.Cells[1, 11] = "安装产品";

            Worksheet ActiveSheet = Globals.ThisAddIn.Application.ActiveSheet;

            int i = 7, j = 2;
            while (ActiveSheet.Cells[i,1] != null)
            {
                j = j + 3;
            }

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

        private void ButtonServiceOrderArrange_Click(object sender, RibbonControlEventArgs e)
        {
            if (CheckFileFormat())
            {
                Excel.Application NewEXCELFile = new Excel.Application();
                Workbook NewWorkbook = NewEXCELFile.Application.Workbooks.Add();
                Worksheet NewWorksheet = NewWorkbook.Worksheets.Add();
                InstallOrderArrange(NewWorkbook, NewWorksheet);



                NewEXCELFile.Visible = true;
            }
            else
            {
                MessageBox.Show("文件格式不正确，请重新核对数据！");
            }
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
    }
}
