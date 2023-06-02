using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.Configuration;
using System.Data;
using System.Data.SqlClient;

public partial class SOS_Rack : System.Web.UI.Page
{
    const int RackHno = 42; //機櫃高度常數

    protected void Page_PreRenderComplete(object sender, EventArgs e)   //放Page_Load無法顯示
    {
        if (!IsPostBack)
        {
                     
        }

        Rack_Main();    //產生機櫃空間示意圖入口  
    }

    protected void Sel_Changed(object sender, EventArgs e)  //選取條件改變
    {
        Rack_Main();    //產生機櫃空間示意圖入口           
    }

    protected void OpenExecWindow(string strJavascript)    //選取後就另開視窗
    {
        if (ScriptManager.GetCurrent(Page) == null)
        {
            Page.ClientScript.RegisterStartupScript(Page.GetType(), "DevEdit", strJavascript, true);
        }
        else
        {
            ScriptManager.RegisterStartupScript(Page, Page.GetType(), "DevEdit", strJavascript, true);
        }
    }

    protected string GetValue(string SQL)   //取得單一資料
    {
        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
        Conn.Open();
        SqlCommand cmd = new SqlCommand(SQL, Conn);
        SqlDataReader dr = null;
        dr = cmd.ExecuteReader();
        string cfg = ""; if (dr.Read()) cfg = dr[0].ToString();
        cmd.Cancel(); cmd.Dispose(); dr.Close(); Conn.Close(); Conn.Dispose();

        return (cfg);
    }
    //------------------------------------------------------------------------------------------------------------------------------------------------------------機櫃空間分區配置表
    protected void Rack_Main()    //產生機櫃空間示意圖入口
    {
        tblRack.Rows.Clear();
        tblRack.Dispose();
        Rack_Head();    //產生機櫃標題
        Rack_Space();   //產生機櫃空間
        Rack_Data();    //產生機櫃設備資訊
    }

    protected void Rack_Head()    //產生機櫃標題
    {
        TableRow row = new TableRow(); TableCell cell0 = new TableCell();
        tblRack.Rows.Add(row);  //產生欄
        cell0.Text = "高度";
        cell0.HorizontalAlign = HorizontalAlign.Center;
        cell0.Font.Bold = true;
        cell0.BackColor = System.Drawing.Color.LightGray;
        row.Cells.Add(cell0);
        
        string Area = SelRack.SelectedValue;
        string SQL = "select [定位編號],[定位名稱] from [定位設定] where [定位方式]='機櫃'";
        string AreaSQL = "";
        switch (Area)
        {
            case "D0": AreaSQL = "([定位名稱] like 'D0_機櫃' or [定位名稱]='D10機櫃')"; break;
            case "D1": AreaSQL = "([定位名稱] like 'D1_機櫃' and [定位名稱]<>'D10機櫃')"; break;
            case "FR": AreaSQL = "([定位名稱] like 'F__機櫃' or [定位名稱] like 'R__機櫃')"; break;
            case "GI": AreaSQL = "([定位名稱] like 'G__機櫃' or [定位名稱] like 'I__機櫃')"; break;
            default: AreaSQL = "[定位名稱] like '" + Area + "_機櫃'"; break;
        }
        ViewState["AreaSQL"] = AreaSQL; //選擇那個機櫃分區

        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
        Conn.Open();
        SqlCommand cmd = new SqlCommand(SQL + " and " + AreaSQL + " order by [定位名稱]", Conn);
        SqlDataReader dr = null;
        dr = cmd.ExecuteReader();
        while (dr.Read())   //標題：機櫃名稱
        {
            TableCell cell = new TableCell();
            cell.Text = dr[1].ToString();
            cell.HorizontalAlign = HorizontalAlign.Center;
            cell.Font.Bold = true;
            cell.BackColor = System.Drawing.Color.LightGray;
            tblRack.Rows[0].Cells.Add(cell);
        }
        dr.Close();
        
        cmd.Cancel(); cmd.Dispose(); dr.Close(); Conn.Close(); Conn.Dispose();
    }

    protected void Rack_Space()    //產生機櫃空間
    {   
        for (int i = RackHno; i >= 0; i--)    //高度由高至低
        {
            TableRow row = new TableRow(); TableCell cell0 = new TableCell();
            tblRack.Rows.Add(row);  //產生欄
            //------------------產生高度編號
            cell0.Text = i.ToString();
            cell0.HorizontalAlign = HorizontalAlign.Center;
            cell0.Font.Bold = true;
            cell0.BackColor = System.Drawing.Color.LightGray;
            row.Cells.Add(cell0);
            //------------------產生機櫃空格
            for (int j = 1; j < tblRack.Rows[0].Cells.Count; j++)
            {
                TableCell cell = new TableCell();
                cell.Font.Size = FontUnit.Smaller;
                cell.Text = "&nbsp;";
                row.Cells.Add(cell);
            }
        }
    }

    protected void Rack_Data()    //產生機櫃設備資訊
    {
        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
        Conn.Open();
        SqlCommand cmd = new SqlCommand("", Conn);
        SqlDataReader dr = null;
        
        int RowNo = 0, CellNo = 0;
        int[,] RackA = new int[tblRack.Rows[0].Cells.Count + 1, RackHno + 2];  //機櫃空間陣列，儲存設備數
        string Unit = SelUnit.SelectedValue, Mt = SelMt.SelectedValue;
        string CheckYear = GetValue("select max(清查年度) from [機器清查]");

        string AssetsSQL = ",(select distinct [Mark] from [Config] where ([Kind]='資產清冊' or [Kind]='常用資產') and [Item]=[資產編號]) as [資產價值]";
        string ColorSQL = "";
        switch (SelColor.SelectedValue)
        {
            case "關機值": ColorSQL = ",(select distinct [Config] from [Config] where [Kind]='關機順序' and [Item]=[關機順序]) as [關機值]"; break;
            case "狀態值": ColorSQL = ",(select distinct [Config] from [Config] where [Kind]='設備狀態' and [Item]=[設備狀態]) as [狀態值]"; break;
            case "清查值": ColorSQL = ",(select distinct [Config] from [Config] where [Kind]='清查狀態' and [Item]=[清查狀態]) as [清查值]"; break;
            case "接電數": ColorSQL = ",[接電迴路],(select count(*)+1 from [接電] where [下游編號]=[設備編號]) as [接電數]"; break;
            case "接網數": ColorSQL = ",[接網迴路],(select count(*)+1 from [接網] where [下游編號]=[設備編號]) as [接網數]"; break;
        }
        
        string MtSQL = "";         
        if (ChkMt.Checked) MtSQL = " and ([設備維護人員] in (SELECT '" + Mt + "' Union SELECT [Kind] FROM [Config] where [Item]='" + Mt + "')"
             + " or [作業維護人員] in (SELECT '" + Mt + "' Union SELECT [Kind] FROM [Config] where [Item]='" + Mt + "')"
              + " or [保管人員] in (SELECT '" + Mt + "' Union SELECT [Kind] FROM [Config] where [Item]='" + Mt + "'))";
        else if (Unit != "") MtSQL = " and ([設備維護人員] in (SELECT [成員] FROM [View_組織架構] where [課別]='" + Unit + "')"
             + " or [作業維護人員] in (SELECT [成員] FROM [View_組織架構] where [課別]='" + Unit + "')"
              + " or [保管人員] in (SELECT [成員] FROM [View_組織架構] where [課別]='" + Unit + "'))";

        string SQL = "select *" + AssetsSQL + ColorSQL + " from [View_通用設備] where [定位編號] in (select [定位編號] from [定位設定] where [定位方式]='機櫃' and " + ViewState["AreaSQL"].ToString() + ")" + MtSQL;
        if (SelColor.SelectedValue == "清查值") SQL = "select *" + AssetsSQL + ColorSQL + ",[View_通用設備].[設備編號] as [設備編號],[機器清查].[備註說明] as [清查說明]" 
            + " from [View_通用設備] Left Outer Join [機器清查] ON [清查年度]='" + CheckYear + "' and [機器清查].[設備編號]=[View_通用設備].[設備編號]"
            + " where [View_通用設備].[定位編號] in (select [定位編號] from [定位設定] where [定位方式]='機櫃' and " + ViewState["AreaSQL"].ToString() + ")" + MtSQL;
        
        cmd.CommandText = SQL;
        dr = cmd.ExecuteReader();
        while (dr.Read())   //標題：機櫃名稱
        {
            RowNo = RackHno + 1 - int.Parse(dr["高度"].ToString());
            CellNo = GetRackIdx(dr["定位名稱"].ToString());
            RackA[CellNo, RowNo] = RackA[CellNo, RowNo] + 1;    //設備數+1

            if (tblRack.Rows[RowNo].Cells[CellNo].Text == "&nbsp;") //第一筆資料
            {
                tblRack.Rows[RowNo].Cells[CellNo].Text = "<span onclick=window.open('../Device/DevEdit.aspx?DevNo=" + dr["設備編號"].ToString() + "','_blank') style='cursor:pointer'>" + dr["設備名稱"].ToString() + "</span>";
            }
            else    //第二筆以上資料
            {
                tblRack.Rows[RowNo].Cells[CellNo].Text = "<span onclick=window.open('../Device/DevEdit.aspx?DevNo=" + dr["設備編號"].ToString() + "','_blank') style='cursor:pointer'>" + dr["設備名稱"].ToString() + " * " + RackA[CellNo, RowNo].ToString() + "</span>";
            }
            tblRack.Rows[RowNo].Cells[CellNo].ToolTip = tblRack.Rows[RowNo].Cells[CellNo].ToolTip + dr["設備編號"].ToString() + "." + dr["設備名稱"].ToString() + "(" + dr["設備狀態"].ToString() + ") " + dr["關機順序"].ToString() + " " + dr["設備維護人員"].ToString()
                    + " $" + dr["資產價值"].ToString() + " ↑" + dr["空間大小"].ToString();
            switch(SelColor.SelectedValue)
            {
                case "清查值":
                    {
                        if (dr["清查狀態"].ToString() == "") tblRack.Rows[RowNo].Cells[CellNo].ToolTip = tblRack.Rows[RowNo].Cells[CellNo].ToolTip + " 尚未列入機器清查" + "\n";
                        else tblRack.Rows[RowNo].Cells[CellNo].ToolTip = tblRack.Rows[RowNo].Cells[CellNo].ToolTip + " 機器清查：" + dr["清查人員"].ToString() + "(" + dr["清查狀態"].ToString() + ") " + dr["清查結果"].ToString() + " " + dr["清查說明"].ToString() + "\n";

                        if ((dr["清查結果"].ToString() + dr["清查說明"].ToString()).Trim() != "") tblRack.Rows[RowNo].Cells[CellNo].Font.Bold = true;
                        break;
                    }
                case "接電數": tblRack.Rows[RowNo].Cells[CellNo].ToolTip = tblRack.Rows[RowNo].Cells[CellNo].ToolTip + " " + dr["接電迴路"].ToString() + "\n"; break;
                case "接網數": tblRack.Rows[RowNo].Cells[CellNo].ToolTip = tblRack.Rows[RowNo].Cells[CellNo].ToolTip + " " + dr["接網迴路"].ToString() + "\n"; break;
                default: tblRack.Rows[RowNo].Cells[CellNo].ToolTip = tblRack.Rows[RowNo].Cells[CellNo].ToolTip + "\n"; break;
            }        

            SetColor(CellNo, RowNo, int.Parse(dr["空間大小"].ToString()), dr[SelColor.SelectedValue].ToString());
        }
        dr.Close();
        //-----------------------------------------
        cmd.Cancel(); cmd.Dispose(); dr.Close(); Conn.Close(); Conn.Dispose();
    }

    protected int GetRackIdx(string Rack)   //取得機櫃索引
    {
        int RackIdx = 0;
        for (int i = 1; i < tblRack.Rows[0].Cells.Count; i++) if (tblRack.Rows[0].Cells[i].Text == Rack) RackIdx = i;

        return (RackIdx);
    }

    protected void SetColor(int CellNo, int RowNo, int space, string Val)   //設定資產價值顏色
    {
        if (space == 0) space = 1;  //空間大小0與1都視為1
        string color = "";

		int spaceMax = (RowNo > space) ? space : RowNo; //防止陣列索引超出範圍
        for (int i = 0; i < spaceMax; i++)
        {
            color = tblRack.Rows[RowNo - i].Cells[CellNo].BackColor.Name;
            
            switch (Val)
            {
                case "":
                case "0":
                    {
                        if (color != "LightPink" & color != "Yellow" & color != "Green" & color != "Gray") tblRack.Rows[RowNo - i].Cells[CellNo].BackColor = System.Drawing.Color.LightGray;
                        break;
                    }                
                case "1":
                    {
                        if (color != "LightPink" & color != "Yellow" & color != "Green") tblRack.Rows[RowNo - i].Cells[CellNo].BackColor = System.Drawing.Color.White;
                        break;
                    }
                case "2":
                    {
                        if (color != "LightPink" & color != "Yellow") tblRack.Rows[RowNo - i].Cells[CellNo].BackColor = System.Drawing.Color.LightGreen;
                        break;
                    }
                case "3":
                    {
                        if (color != "LightPink") tblRack.Rows[RowNo - i].Cells[CellNo].BackColor = System.Drawing.Color.Yellow;
                        break;
                    }
                default:
                    {
                        tblRack.Rows[RowNo - i].Cells[CellNo].BackColor = System.Drawing.Color.LightPink;
                        break;
                    }
            }
        }
    }
}