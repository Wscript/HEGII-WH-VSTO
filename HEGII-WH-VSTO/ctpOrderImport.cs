using System;
using System.Windows.Forms;
using System.Configuration;
using System.Data.SqlClient;
using Microsoft.Office.Interop.Excel;

namespace HEGII_WH_VSTO
{
    public partial class ctpOrderImport : UserControl
    {
        public ctpOrderImport()
        {
            InitializeComponent();
        }

        private void ctpOrderImport_Load(object sender, EventArgs e)
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            string consqlserver = ConfigurationManager.ConnectionStrings["HGWHConnectionString"].ToString() + ";Password=HEGII;";
            string sql = "SELECT OrderType FROM OrderTypeList";
            SqlConnection con = new SqlConnection(consqlserver);
            SqlDataAdapter da = new SqlDataAdapter(sql, con);
            try
            {
                da.Fill(dt);
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    comboOldOrderType.Items.Add(dt.Rows[i][0].ToString());
                }
            }
            catch (Exception msg)
            {
                MessageBox.Show(msg.Message);
            }
            finally
            {
                con.Close();
                con.Dispose();
                da.Dispose();
                dt.Dispose();
            }
        }

        private bool CheckFileFormat()
        {
            Worksheet ActiveSheet = Globals.ThisAddIn.Application.ActiveSheet;

            Range rangeCustAddress = ActiveSheet.Range["1:1"].Find("地址");
            if (rangeCustAddress == null)
            {
                return (false);
            }
            Range rangeCustName = ActiveSheet.Range["1:1"].Find("用户名称");
            if (rangeCustName == null)
            {
                return (false);
            }
            Range rangeProductList = ActiveSheet.Range["1:1"].Find("安装产品");
            if (rangeProductList == null)
            {
                return (false);
            }
            Range rangeCustPhone = ActiveSheet.Range["1:1"].Find("电话");
            if (rangeCustPhone == null)
            {
                return (false);
            }
            Range rangeSalesStore = ActiveSheet.Range["1:1"].Find("销售点");
            if (rangeSalesStore == null)
            {
                return (false);
            }
            Range rangeApplyDate = ActiveSheet.Range["1:1"].Find("报装");
            if (rangeApplyDate == null)
            {
                return (false);
            }
            Range rangeReserveDate = ActiveSheet.Range["1:1"].Find("预约");
            if (rangeReserveDate == null)
            {
                return (false);
            }
            Range rangeServiceArea = ActiveSheet.Range["1:1"].Find("大范围");
            if (rangeServiceArea == null)
            {
                return (false);
            }
            return (true);
        }

        private int checkCustInfo(string strCustAddress, string strCustPhone, int intDataRow)
        {
            Worksheet ActiveSheet = Globals.ThisAddIn.Application.ActiveSheet;
            System.Data.DataTable dt = new System.Data.DataTable();
            string consqlserver = ConfigurationManager.ConnectionStrings["HGWHConnectionString"].ToString() + ";Password=HEGII;";
            string sql = "EXEC progGetCustInfo '" + ActiveSheet.Cells[intDataRow, ActiveSheet.Range["1:1"].Find("地址").Column].Value + "', '" +
                                                    ActiveSheet.Cells[intDataRow, ActiveSheet.Range["1:1"].Find("电话").Column].Value + "'";
            SqlConnection con = new SqlConnection(consqlserver);
            SqlDataAdapter da = new SqlDataAdapter(sql, con);
            try
            {
                da.Fill(dt);
                switch (dt.Rows.Count)
                {
                    case 0:
                        return (0);     //没有找到匹配的客户信息，返回0
                    case 1:
                        return ((int)dt.Rows[0]["ID"]);     //找到一条匹配的信息，返回客户信息ID
                    default:
                        return (0 - dt.Rows.Count);     //找到多条匹配的信息，以负数形式返回重复的信息数量
                }
            }
            catch (Exception msg)
            {
                MessageBox.Show(msg.Message);
                return (-1);        //发生错误返回-1
            }
            finally
            {
                con.Close();
                con.Dispose();
                da.Dispose();
                dt.Dispose();
            }
        }

        private int insertCustInfo(int intDataRow)
        {
            Worksheet ActiveSheet = Globals.ThisAddIn.Application.ActiveSheet;
            System.Data.DataTable dt = new System.Data.DataTable();
            string consqlserver = ConfigurationManager.ConnectionStrings["HGWHConnectionString"].ToString() + ";Password=HEGII;";
            string sql = "EXEC progInsertCustInfo '" + ActiveSheet.Cells[intDataRow, ActiveSheet.Range["1:1"].Find("地址").Column].Value + "', '" +
                                                       ActiveSheet.Cells[intDataRow, ActiveSheet.Range["1:1"].Find("电话").Column].Value + "'";
            if (ActiveSheet.Cells[intDataRow, ActiveSheet.Range["1:1"].Find("用户名称").Column].Value == null)
            {
                sql = sql + ", NULL";
            }
            else
            {
                sql = sql + ", '" + ActiveSheet.Cells[intDataRow, ActiveSheet.Range["1:1"].Find("用户名称").Column].Value + "'";
            }
            if (ActiveSheet.Cells[intDataRow, ActiveSheet.Range["1:1"].Find("大范围").Column].Value == null)
            {
                sql = sql + ", NULL";
            }
            else
            {
                sql = sql + ", '" + ActiveSheet.Cells[intDataRow, ActiveSheet.Range["1:1"].Find("大范围").Column].Value + "'";
            }

            SqlConnection con = new SqlConnection(consqlserver);
            SqlDataAdapter da = new SqlDataAdapter(sql, con);
            try
            {
                da.Fill(dt);
                if (dt.Rows.Count > 0)
                {
                    return ((int)dt.Rows[0][0]);
                }
                else
                {
                    return (0);
                }
            }
            catch (Exception msg)
            {
                MessageBox.Show(msg.Message);
                return (-1);
            }
            finally
            {
                con.Close();
                con.Dispose();
                da.Dispose();
                dt.Dispose();
            }
        }

        private int insertOrder(int intDataRow, int intCustInfoID)
        {
            Worksheet ActiveSheet = Globals.ThisAddIn.Application.ActiveSheet;
            System.Data.DataTable dt = new System.Data.DataTable();
            string consqlserver = ConfigurationManager.ConnectionStrings["HGWHConnectionString"].ToString() + ";Password=HEGII;";
            string sql = "EXEC progInsertOrder " + intCustInfoID + ", '历史数据导入'";
            if (ActiveSheet.Cells[intDataRow, ActiveSheet.Range["1:1"].Find("报装").Column].Value == null)
            {
                sql = sql + ", NULL";
            }
            else
            {
                if (DateTime.TryParse(ActiveSheet.Cells[intDataRow, ActiveSheet.Range["1:1"].Find("报装").Column].Value.ToString(), out DateTime dtApplyDate))
                {
                    sql = sql + ", '" + dtApplyDate.ToString() + "'";
                }
                else
                {
                    sql = sql + ", NULL";
                }
            }
            if (ActiveSheet.Cells[intDataRow, ActiveSheet.Range["1:1"].Find("预约").Column].Value == null)
            {
                sql = sql + ", NULL";
            }
            else
            {
                if (DateTime.TryParse(ActiveSheet.Cells[intDataRow, ActiveSheet.Range["1:1"].Find("预约").Column].Value.ToString(), out DateTime dtReserveDate))
                {
                    sql = sql + ", '" + dtReserveDate.ToString() + "'";
                }
                else
                {
                    sql = sql + ", NULL";
                }
            }
            if (ActiveSheet.Cells[intDataRow, ActiveSheet.Range["1:1"].Find("销售点").Column].Value == null)
            {
                sql = sql + ", NULL";
            }
            else
            {
                sql = sql + ", '" + ActiveSheet.Cells[intDataRow, ActiveSheet.Range["1:1"].Find("销售点").Column].Value.ToString() + "'";
            }
            sql = sql + ", '" + comboOldOrderType.Text + "', 0";
            if (ActiveSheet.Cells[intDataRow, ActiveSheet.Range["1:1"].Find("安装产品").Column].Value == null)
            {
                sql = sql + ", NULL";
            }
            else
            {
                sql = sql + ", '" + ActiveSheet.Cells[intDataRow, ActiveSheet.Range["1:1"].Find("安装产品").Column].Value + "'";
            }
            if (ActiveSheet.Cells[intDataRow, ActiveSheet.Range["1:1"].Find("预约").Column].Value == null)
            {
                sql = sql + ", NULL";
            }
            else
            {
                sql = sql + ", '" + ActiveSheet.Cells[intDataRow, ActiveSheet.Range["1:1"].Find("预约").Column].Value.ToString() + "'";
            }
            if (ActiveSheet.Cells[intDataRow, ActiveSheet.Range["1:1"].Find("预约").Column].Value == null)
            {
                sql = sql + ", '待上门'";
            }
            else
            {
                if (DateTime.TryParse(ActiveSheet.Cells[intDataRow, ActiveSheet.Range["1:1"].Find("预约").Column].Value.ToString(), out DateTime dtReserveDate))
                {
                    if (dtReserveDate > DateTime.Today)
                    {
                        sql = sql + ", '待上门'";
                    }
                    else
                    {
                        sql = sql + ", '完成'";
                    }
                }
                else
                {
                    if (ActiveSheet.Cells[intDataRow, ActiveSheet.Range["1:1"].Find("预约").Column].Value.IndexOf("改期") > 0)
                    {
                        sql = sql + ", '改期'";
                    }
                    else
                    {
                        sql = sql + ", '完成'";
                    }
                }
            }
            if (ActiveSheet.Cells[intDataRow, ActiveSheet.Range["1:1"].Find("安装产品").Column + 1].Value == null)
            {
                sql = sql + ", NULL";
            }
            else
            {
                sql = sql + ", '" + ActiveSheet.Cells[intDataRow, ActiveSheet.Range["1:1"].Find("安装产品").Column + 1].Value + "'";
            }

            SqlConnection con = new SqlConnection(consqlserver);
            SqlDataAdapter da = new SqlDataAdapter(sql, con);
            try
            {
                da.Fill(dt);
                if (dt.Rows.Count > 0)
                {
                    return ((int)dt.Rows[0][0]);
                }
                else
                {
                    return (0);
                }
            }
            catch (Exception msg)
            {
                MessageBox.Show(msg.Message);
                return (-1);
            }
            finally
            {
                con.Close();
                con.Dispose();
                da.Dispose();
                dt.Dispose();
            }
        }

        private int checkOrder(int intDataRow, int intCustInfoID)
        {
            Worksheet ActiveSheet = Globals.ThisAddIn.Application.ActiveSheet;
            System.Data.DataTable dt = new System.Data.DataTable();
            string consqlserver = ConfigurationManager.ConnectionStrings["HGWHConnectionString"].ToString() + ";Password=HEGII;";
            string sql = "EXEC progGetOrder " + intCustInfoID.ToString();
            if (ActiveSheet.Cells[intDataRow, ActiveSheet.Range["1:1"].Find("报装").Column].Value == null)
            {
                sql = sql + ", NULL";
            }
            else
            {
                if (DateTime.TryParse(ActiveSheet.Cells[intDataRow, ActiveSheet.Range["1:1"].Find("报装").Column].Value.ToString(), out DateTime dtApplyDate))
                {
                    sql = sql + ", '" + dtApplyDate.ToString() + "'";
                }
                else
                {
                    sql = sql + ", NULL";
                }
            }
            if (ActiveSheet.Cells[intDataRow, ActiveSheet.Range["1:1"].Find("预约").Column].Value == null)
            {
                sql = sql + ", NULL";
            }
            else
            {
                sql = sql + ", '" + ActiveSheet.Cells[intDataRow, ActiveSheet.Range["1:1"].Find("预约").Column].Value.ToString() + "'";
            }
            sql = sql + ", '" + comboOldOrderType.Text + "'";
            SqlConnection con = new SqlConnection(consqlserver);
            SqlDataAdapter da = new SqlDataAdapter(sql, con);
            try
            {
                da.Fill(dt);
                switch (dt.Rows.Count)
                {
                    case 0:
                        return (0);
                    default:
                        return (-1 - dt.Rows.Count);
                }
            }
            catch (Exception msg)
            {
                MessageBox.Show(msg.Message);
                return (-1);
            }
            finally
            {
                con.Close();
                con.Dispose();
                da.Dispose();
                dt.Dispose();
            }
        }

        private void buttonOldDataOK_Click(object sender, EventArgs e)
        {
            int intCheckCustInfo, intinsertCustInfo, intcheckOrder, intinsertOrder;
            Worksheet ActiveSheet = Globals.ThisAddIn.Application.ActiveSheet;

            if (CheckFileFormat())
            {
                for (int i = 2; i <= ActiveSheet.UsedRange.Rows.Count; i++)
                {
                    ActiveSheet.Rows[i].Select();
                    if (ActiveSheet.Cells[i, ActiveSheet.Range["1:1"].Find("地址").Column].Value == null ||
                        ActiveSheet.Cells[i, ActiveSheet.Range["1:1"].Find("电话").Column].Value == null)
                    {
                        ActiveSheet.Cells[i, 1].Value = "地址或电话为空，未导入数据！";
                        ActiveSheet.Cells[i, 1].Interior.ColorIndex = 3;
                    }
                    else
                    {
                        intCheckCustInfo = checkCustInfo(ActiveSheet.Cells[i, ActiveSheet.Range["1:1"].Find("地址").Column].Value.ToString(),
                                                             ActiveSheet.Cells[i, ActiveSheet.Range["1:1"].Find("电话").Column].Value.ToString(), i);
                        if (intCheckCustInfo == -1)
                        {
                            ActiveSheet.Cells[i, 1].Value = "程序运行出现错误,客户信息获取失败" + ", 未导入数据！";
                            ActiveSheet.Cells[i, 1].Interior.ColorIndex = 3;
                        }
                        if (intCheckCustInfo < -1)
                        {
                            ActiveSheet.Cells[i, 1].Value = "找到" + Convert.ToString(0 - intCheckCustInfo) + "条相同的客户信息" + ", 未导入数据！";
                            ActiveSheet.Cells[i, 1].Interior.ColorIndex = 3;
                        }
                        if (intCheckCustInfo == 0)
                        {
                            intinsertCustInfo = insertCustInfo(i);
                            switch (intinsertCustInfo)
                            {
                                case -1 :
                                    ActiveSheet.Cells[i, 1].Value = "程序运行出现错误,客户信息导入失败" + ", 未导入数据！";
                                    ActiveSheet.Cells[i, 1].Interior.ColorIndex = 3;
                                    break;
                                case 0 :
                                    ActiveSheet.Cells[i, 1].Value = "客户信息导入失败" + ", 未导入数据！";
                                    ActiveSheet.Cells[i, 1].Interior.ColorIndex = 3;
                                    break;
                                default:
                                    intCheckCustInfo = intinsertCustInfo;
                                    break;
                            }
                        }
                        if (intCheckCustInfo > 0)
                        {
                            intcheckOrder = checkOrder(i, intCheckCustInfo);
                            switch (intcheckOrder)
                            {
                                case -1 :
                                    ActiveSheet.Cells[i, 1].Value = "程序运行出现错误,订单信息获取失败" + ", 未导入数据！";
                                    ActiveSheet.Cells[i, 1].Interior.ColorIndex = 3;
                                    break;
                                case 0 :
                                    intinsertOrder = insertOrder(i, intCheckCustInfo);
                                    switch (intinsertOrder)
                                    {
                                        case -1 :
                                            ActiveSheet.Cells[i, 1].Value = "程序运行出现错误,订单信息导入失败" + ", 未导入数据！";
                                            ActiveSheet.Cells[i, 1].Interior.ColorIndex = 3;
                                            break;
                                        case 0:
                                            ActiveSheet.Cells[i, 1].Value = "订单信息导入失败" + ", 未导入数据！";
                                            ActiveSheet.Cells[i, 1].Interior.ColorIndex = 3;
                                            break;
                                        default:
                                            ActiveSheet.Cells[i, 1].Value = "已导入订单信息！";
                                            ActiveSheet.Cells[i, 1].Interior.ColorIndex = 4;
                                            break;
                                    }
                                    break;
                                default:
                                    ActiveSheet.Cells[i, 1].Value = "找到" + Convert.ToString(0 - intcheckOrder - 1) + "条相同的订单信息" + "，未导入数据！";
                                    ActiveSheet.Cells[i, 1].Interior.ColorIndex = 4;
                                    break;
                            }
                        }
                    }
                }
            }
            else
            {
                MessageBox.Show("文件格式不正确，请核对！");
            }
        }
    }
}
