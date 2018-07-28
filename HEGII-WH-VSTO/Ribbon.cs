using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Ribbon;
using System.IO;
using System.Net;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using HtmlAgilityPack;
using System.Diagnostics;

namespace HEGII_WH_VSTO
{
    public partial class Ribbon
    {
        private void Ribbon_Load(object sender, RibbonUIEventArgs e)
        {

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
            string htmlDistrictPage;
            string stringWebsiteAddress = "https://wh.lianjia.com";       //设置网站地址,结尾不要用"/"
            string htmlStartPage = getHtmlString(stringWebsiteAddress + "/xiaoqu/");         //取起始页面的HTML代码
            HtmlNode HtmlNodeStartPage = getHtmlNode(htmlStartPage);
            HtmlNodeCollection HtmlNodeCollectionDistrict = HtmlNodeStartPage.SelectNodes("//a[contains(@title,'小区二手房')]");      //找到行政区页面链接
            foreach (HtmlNode HtmlNodeDistrict in HtmlNodeCollectionDistrict)       //遍历所有行政区
            {
                htmlDistrictPage = "";
                if (HtmlNodeDistrict.Attributes["href"].Value != "/xiaoqu/")        //跳过起始页面
                {
                    htmlDistrictPage = getHtmlString(stringWebsiteAddress + HtmlNodeDistrict.Attributes["href"].Value);      //取各行政区页面的HTML代码
                    HtmlNode HtmlNodeDistrictPage = getHtmlNode(htmlDistrictPage);
                    HtmlNode HtmlNodePageCount = HtmlNodeDistrictPage.SelectSingleNode("//div[@class=\"page-box house-lst-page-box\"]");
                    int intCommunityPageCount = int.Parse(HtmlNodePageCount.Attributes["page-data"].Value.Substring(13, HtmlNodePageCount.Attributes["page-data"].Value.IndexOf(",") - 13));
                    //取小区列表的总页数
                    for (int i = 1; i < intCommunityPageCount; i++)
                    {
                        if (i > 1)      //跳过当前页
                        {
                            htmlDistrictPage = getHtmlString(stringWebsiteAddress + HtmlNodePageCount.Attributes["page-url"].Value.Replace("{page}", i.ToString()));
                        }
                        getCommunityInfo(htmlDistrictPage);     //处理小区信息
                    }
                }
            }
        }

        public static string getHtmlString(string address)     //取指定地址页面的HTML代码
        {
            WebClient client = new WebClient();
            client.Encoding = System.Text.Encoding.GetEncoding("UTF-8");
            string HtmlString = client.DownloadString(address);
            return (HtmlString);
        }

        public static HtmlNode getHtmlNode(string htmlPage)     //用HtmlAgilityPack处理HTML代码
        {
            HtmlAgilityPack.HtmlDocument HtmlDocumentPage = new HtmlAgilityPack.HtmlDocument();
            HtmlDocumentPage.LoadHtml(htmlPage);
            HtmlNode HtmlNodePage = HtmlDocumentPage.DocumentNode;
            return (HtmlNodePage);
        }

        public static void getCommunityInfo(string htmlDistrictPage)        //读取小区的信息
        {
            HtmlNode HtmlNodeCommunityList = getHtmlNode(htmlDistrictPage);
            HtmlNodeCollection HtmlNodeCollectionCommunity = HtmlNodeCommunityList.SelectNodes("//li[@class=\"clear xiaoquListItem\"]");
            foreach (HtmlNode HtmlNodeCommunity in HtmlNodeCollectionCommunity)
            {
                string htmlCommunityItem = HtmlNodeCommunity.InnerHtml;
                HtmlNode HtmlNodeCommunityItem = getHtmlNode(htmlCommunityItem);
                HtmlNode HtmlNodeCommunityName = HtmlNodeCommunityItem.SelectSingleNode("//div[@class=\"title\"]");
                Debug.WriteLine(HtmlNodeCommunityName.ChildNodes[1].InnerText);
                HtmlNode HtmlNodeCommunityPosition = HtmlNodeCommunityItem.SelectSingleNode("//div[@class=\"positionInfo\"]");
                Debug.WriteLine(HtmlNodeCommunityPosition.ChildNodes[3].InnerText);
                Debug.WriteLine(HtmlNodeCommunityPosition.ChildNodes[5].InnerText);
                string htmlCommunityPage = getHtmlString(HtmlNodeCommunityName.ChildNodes[1].Attributes["href"].Value);
                HtmlNode HtmlNodeCommunityPage = getHtmlNode(htmlCommunityPage);
                HtmlNode HtmlNodeCommunityAddress = HtmlNodeCommunityPage.SelectSingleNode("//div[@class=\"detailDesc\"]");
                Debug.WriteLine(HtmlNodeCommunityAddress.InnerText.Substring(HtmlNodeCommunityAddress.InnerText.IndexOf(")")+1));
                HtmlNode HtmlNodeCommunityInfo = HtmlNodeCommunityPage.SelectSingleNode("//div[@class=\"xiaoquInfo\"]");
                Debug.WriteLine(HtmlNodeCommunityInfo.ChildNodes[0].ChildNodes[1].InnerText);
                Debug.WriteLine(HtmlNodeCommunityInfo.ChildNodes[1].ChildNodes[1].InnerText);
                Debug.WriteLine(HtmlNodeCommunityInfo.ChildNodes[5].ChildNodes[1].InnerText);
                Debug.WriteLine(HtmlNodeCommunityInfo.ChildNodes[6].ChildNodes[1].InnerText);
                Debug.WriteLine("");
            }

        }
    }
}
