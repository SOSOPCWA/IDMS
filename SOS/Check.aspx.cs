using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.Configuration;
using System.Data;
using System.Data.SqlClient;

public partial class SOS_Check : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        if (!IsPostBack)
        {
            if(Request["Del"]=="Y") 
            {
                string strChkYear = Request["ChkYear"].ToString();
                ExecDbSQL("delete from [機器清查] where [清查年度]='" + strChkYear + "'");
                Literal Msg = new Literal();
                Msg.Text = "<script>alert('已刪除 " + strChkYear + " 機器清查清單全部資料！');window.open('Check.aspx','_self');</script>";
                Page.Controls.Add(Msg);
            }

            AddChkYear(); //產生清查年度選項
        }
    }

    protected void AddChkYear() //產生清查年度選項
    {
        string ChkYear = ""; if (Session["ChkYear"] != null) ChkYear = Session["ChkYear"].ToString();
        
        ListItem ItemX = new ListItem(); ItemX.Text = "(當月新增)"; ItemX.Value = ""; SelChkYear.Items.Add(ItemX);

        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
        Conn.Open();
        SqlCommand cmd = new SqlCommand("select distinct [清查年度] from [機器清查] order by [清查年度] desc", Conn);
        SqlDataReader dr = null;

        int i = 0;
        dr = cmd.ExecuteReader();
        while (dr.Read())
        {
            ListItem Item = new ListItem(); Item.Text = dr[0].ToString(); Item.Value = dr[0].ToString();
            if (ChkYear == "" & i == 0 | ChkYear != "" & Item.Value == ChkYear) Item.Selected = true;   //預設為最近清查年度
            SelChkYear.Items.Add(Item);
            i++;
        }

        cmd.Cancel(); dr.Close(); Conn.Close(); Conn.Dispose();        
    }

    protected void Page_PreRenderComplete(object sender, EventArgs e)   //放Page_Load無法顯示
    {
        if (ViewState["SQL"] != null) SqlDataSourceSearch.SelectCommand = ViewState["SQL"].ToString();

        if (!IsPostBack)
        {
            PanelChange();
        }
    }

    protected string GetPosSQL(string PosNo)
    {
        string SQL = "select [定位編號] from [定位設定]";

        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
        Conn.Open();
        SqlCommand cmd = new SqlCommand("select * from [定位設定] where [定位編號]=" + PosNo, Conn);
        SqlDataReader dr = null;
        dr = cmd.ExecuteReader();
        if (dr.Read()) SQL = SQL + " where [坐標X]>=" + dr["坐標X1"].ToString() + " and [坐標X]<=" + dr["坐標X2"].ToString()
            + " and [坐標Y]>=" + dr["坐標Y1"].ToString() + " and [坐標Y]<=" + dr["坐標Y2"].ToString();
        
        cmd.Cancel(); dr.Close(); Conn.Close(); Conn.Dispose();        
        return (SQL);
    }

    protected void BtnSearch_Click(object sender, EventArgs e)
    {
        string SQL = "SELECT * FROM [View_機器清查] WHERE 1=1", ChksSQL = "";
        if (ChkPage.Checked) GridView1.AllowPaging = true;
        else GridView1.AllowPaging = false;

        if (ChkYear.Checked)
        {
            SQL = SQL + " and [清查年度]='" + SelChkYear.SelectedValue + "'";
            Session["ChkYear"] = SelChkYear.SelectedValue;
        }

        if (ChkType.Checked) SQL = SQL + " and [設備型態]='" + SelType.SelectedValue + "'";
        if (ChkKind.Checked) SQL = SQL + " and [設備種類]='" + SelKind.SelectedValue + "'";
        if (ChkPlace.Checked) SQL = SQL + " and ([定位編號] in (select [定位編號] from [定位設定] where [區域名稱]='" + SelPlace.SelectedValue + "')"
            + " or [設備定位編號] in (select [定位編號] from [定位設定] where [區域名稱]='" + SelPlace.SelectedValue + "'))";
        if (ChkPointer.Checked) SQL = SQL + " and ([定位編號]=" + SelPointer.SelectedValue + " or [定位編號] in (" + GetPosSQL(SelPointer.SelectedValue) + ")"
            + "or [設備定位編號]=" + SelPointer.SelectedValue + " or [設備定位編號] in (" + GetPosSQL(SelPointer.SelectedValue) + "))";

        if (ChkCircuit.Checked) SQL = SQL + " and ([設備編號]=" + SelCircuit.SelectedValue + GetPowerTreeSQL(int.Parse(SelCircuit.SelectedValue)) + ")";
        if (ChkPowerNum.Checked)
        {
            switch (SelPowerNum.SelectedValue)
            {
                case "0": SQL = SQL + " and [設備編號] not in (select [下游編號] from [接電])"; break;
                case "1": SQL = SQL + " and (select count(*) as PowerNum  from [接電] where [下游編號]=[設備編號] group by [下游編號])=1"; break;
                case "2": SQL = SQL + " and (select count(*) as PowerNum  from [接電] where [下游編號]=[設備編號] group by [下游編號])=2"; break;
                case "*": SQL = SQL + " and (select count(*) as PowerNum  from [接電] where [下游編號]=[設備編號] group by [下游編號])>2"; break;
                default: SQL = SQL + " and [設備編號] not in (select [下游編號] from [接電])"; break;
            }
        }

        if (ChkHwUnit.Checked) SQL = SQL + " and ([維護人員] in (select [成員] from [View_組織架構] where [課別]='" + SelHwUnit.SelectedValue + "')"
            + " or [設備維護人員] in (select [成員] from [View_組織架構] where [課別]='" + SelHwUnit.SelectedValue + "'))";
        if (ChkHw.Checked) SQL = SQL + " and ([維護人員]='" + SelHw.SelectedValue + "' or [維護人員] in (select Kind from Config where Item='" + SelHw.SelectedValue + "')"
            + " or [設備維護人員]='" + SelHw.SelectedValue + "' or [設備維護人員] in (select Kind from Config where Item='" + SelHw.SelectedValue + "'))";
        //Response.Write(SQL);
        //Response.End();

        if (ChkKeepUnit.Checked) SQL = SQL + " and [保管人員] in (select [成員] from [View_組織架構] where [課別]='" + SelKeepUnit.SelectedValue + "')";
        if (ChkKeeper.Checked) SQL = SQL + " and [保管人員]='" + SelKeeper.SelectedValue + "'";

        if (Chker.Checked) SQL = SQL + " and [清查人員]='" + SelChker.SelectedValue + "'";
        if (ChkUpdater.Checked) SQL = SQL + " and [更新人員]='" + SelUpdater.SelectedValue + "'";

        if (ChkStatus.Checked) SQL = SQL + " and [清查狀態]='" + SelStatus.SelectedValue + "'";
        if (ChkDevStatus.Checked) SQL = SQL + " and [設備狀態]='" + SelDevStatus.SelectedValue + "'";
        if (ChkAdd.Checked) SQL = SQL + " and [設備編號]<0";
        if (ChkSQL.Checked & TextSQL.Text != "") SQL = SQL + " and (" + TextSQL.Text + ")";

        if (ChkCheck.Checked)
        {
            for (int i = 0; i < ChkChecks.Items.Count; i++) if (ChkChecks.Items[i].Selected) ChksSQL = ChksSQL + " or [清查結果] like '%" + ChkChecks.Items[i].Value + "%'";
            SQL = SQL + " and (1<>1" + ChksSQL + ")";
        }

        if (ChkKey.Checked)
        {
            string[] KeyA = TextKey.Text.Trim().Split(',');

            for (int i = 0; i < KeyA.GetLength(0); i++)
            {
                SQL = SQL + " and ([清查年度] like '%" + KeyA[i] + "%'"
                   + " or [區域名稱]+[定位名稱] like '%" + KeyA[i] + "%'"
                   + " or [維護人員] like '%" + KeyA[i] + "%'"
                   + " or [清查結果] like '%" + KeyA[i] + "%'"
                   + " or [備註說明] like '%" + KeyA[i] + "%'"
                   + " or [清查狀態] like '%" + KeyA[i] + "%'"
                   + " or [清查人員] like '%" + KeyA[i] + "%'"
                   + " or [更新人員] like '%" + KeyA[i] + "%'"
                   + " or [設備型態] like '%" + KeyA[i] + "%'"
                   + " or [設備種類] like '%" + KeyA[i] + "%'"
                   + " or [設備名稱] like '%" + KeyA[i] + "%'"
                   + " or [設備用途] like '%" + KeyA[i] + "%'"
                   + " or [財產編號] like '%" + KeyA[i] + "%'"
                   + " or [財產名稱] like '%" + KeyA[i] + "%'"
                   + " or [財產別名] like '%" + KeyA[i] + "%'"
                   + " or [廠牌] like '%" + KeyA[i] + "%'"
                   + " or [型式] like '%" + KeyA[i] + "%'"
                   + " or [數量單位] like '%" + KeyA[i] + "%'"
                   + " or [保管人員] like '%" + KeyA[i] + "%'"
                   + " or [財產保管人員] like '%" + KeyA[i] + "%'"
                   + " or [取得來源] like '%" + KeyA[i] + "%'"
                   + " or [規格] like '%" + KeyA[i] + "%'"
                   + " or [硬體序號] like '%" + KeyA[i] + "%'"
                   + " or [iLoIP] like '%" + KeyA[i] + "%'"
                   + " or [維護廠商] like '%" + KeyA[i] + "%'"
                   + " or [設備維護人員] like '%" + KeyA[i] + "%'"
                   + " or [設備狀態] like '%" + KeyA[i] + "%'"
                   + " or [建立人員] like '%" + KeyA[i] + "%'"
                   + " or [修改人員] like '%" + KeyA[i] + "%'"
                   + " or [設備備註說明] like '%" + KeyA[i] + "%'"
                   + " or [設備定位方式] like '%" + KeyA[i] + "%'"
                   + " or [設備放置地點] like '%" + KeyA[i] + "%'"
                   + ")";
            }
        }
        int pos = SQL.IndexOf("FROM");
        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
        Conn.Open();
        SqlCommand cmd = new SqlCommand("SELECT sum([價值]) from [View_機器清查] where [設備編號] in (select [設備編號] " + SQL.Substring(pos) + ")", Conn);
        //Response.Write("SELECT sum([價值]) from [實體設備] where [設備編號] in (select [設備編號] " + SQL.Substring(pos) + ")");
        //Response.End();
        SqlDataReader dr = null;

        dr = cmd.ExecuteReader();
        try { if (dr.Read()) TextTotal.Text = String.Format("{0:C0}", dr[0]); }
        catch { }
        dr.Close();

        cmd.CommandText = "SELECT sum([額定電流]) " + SQL.Substring(pos) + " and ([設備種類]<>'電源' and [設備種類]<>'PDC' and [設備種類]<>'配電盤' and [設備種類]<>'迴路')";
        dr = cmd.ExecuteReader();
        try { if (dr.Read()) TextSumCurrent.Text = String.Format("{0:F}", dr[0]); }
        catch { }
        dr.Close();

        cmd.CommandText = "SELECT count(*) " + SQL.Substring(pos);
        dr = cmd.ExecuteReader();
        TextCount.Text = "0";
        try { if (dr.Read()) TextCount.Text = dr[0].ToString(); }
        catch { }

        cmd.Cancel(); dr.Close(); Conn.Close(); Conn.Dispose();

        string ViewCols = SelShow.SelectedValue; //所有欲顯示欄位名稱
        SQL = "Select " + ViewCols + " " + SQL.Substring(pos) + " order by [設備放置地點],[高度] desc";
        SqlDataSourceSearch.SelectCommand = SQL;
        GridView1.DataBind();
        ViewState["SQL"] = SQL;
        //Response.Write(SQL);
    }

    protected string GetPowerTreeSQL(int DevNo)
    {
        string SQL = "";

        SqlCommand sqlQuery = new SqlCommand();
        sqlQuery.CommandText = "select [下游編號] from [接電] where [上游編號]=" + DevNo;
        DataSet ds = RunQuery(sqlQuery);

        if (ds.Tables.Count > 0)
        {
            foreach (DataRow row in ds.Tables[0].Rows)
            {
                SQL = SQL + " or [設備編號]=" + row[0].ToString() + GetPowerTreeSQL(int.Parse(row[0].ToString()));
            }
        }
        sqlQuery.Cancel(); ds.Dispose();

        return (SQL);
    }

    protected DataSet RunQuery(SqlCommand sqlQuery) //讀取節點資訊
    {
        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
        SqlDataAdapter dbAdapter = new SqlDataAdapter();
        dbAdapter.SelectCommand = sqlQuery;
        sqlQuery.Connection = Conn;
        DataSet QueryDataSet = new DataSet();

        try
        {
            dbAdapter.Fill(QueryDataSet);
        }
        catch
        {
            Response.Write(sqlQuery.CommandText);
            Response.End();
        }
        dbAdapter.Dispose(); Conn.Close(); Conn.Dispose();
        return (QueryDataSet);
    }

    protected void GridView1_SelectedIndexChanging(object sender, GridViewSelectEventArgs e)
    {
        //OpenExecWindow("window.open('../Device/DevEdit.aspx?DevNo=" + GridView1.DataKeys[e.NewSelectedIndex].Value.ToString() + "','_self');");        
    }

    protected void GridView1_RowCommand(object sender, CommandEventArgs e)
    {
        if (e.CommandName == "ChkOK" | e.CommandName == "ChkXX" | e.CommandName == "ChkDev")
        {
            int Idx = int.Parse(e.CommandArgument.ToString());
            string DevNo = GridView1.DataKeys[Idx].Value.ToString();
            string ChkYear = GridView1.Rows[Idx].Cells[4].Text;
            string Chker = Session["UserName"].ToString();
            string ChkDT = DateTime.Now.ToString("yyyy/MM/dd HH:mm");


            switch (e.CommandName)
            {
                case "ChkOK":
                    SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
                    Conn.Open();
                    SqlCommand cmd = new SqlCommand("select * from [機器清查] where [清查年度]='" + ChkYear + "' and [設備編號]=" + DevNo, Conn);
                    SqlDataReader dr = null;

                    Literal Msg = new Literal();

                    dr = cmd.ExecuteReader();
                    if (dr.Read())
                    {
                        if (dr["清查狀態"].ToString() == "待清查")
                        {
                            ExecDbSQL("Update [機器清查] set [清查狀態]='已結案',"
                                + "[清查人員]='" + Chker + "',[清查時間]='" + ChkDT + "',[更新人員]='" + Chker + "',[更新時間]='" + ChkDT
                                + "' where [清查年度]='" + ChkYear + "' and [設備編號]=" + DevNo);
                        }
                        else
                        {
                            ExecDbSQL("Update [機器清查] set [清查狀態]='已結案',"
                                + "[更新人員]='" + Chker + "',[更新時間]='" + ChkDT
                                + "' where [清查年度]='" + ChkYear + "' and [設備編號]=" + DevNo);
                        }
                        Msg.Text = "<script>alert('" + ChkYear + " " + DevNo + "已設為清查符合！');</script>";
                    }
                    else
                    {
                        Msg.Text = "<script>alert('查無資料(" + ChkYear + " " + DevNo + ")！');</script>";
                    }
                    cmd.Cancel(); dr.Close(); Conn.Close(); Conn.Dispose();

                    Page.Controls.Add(Msg);
                    break;
                case "ChkXX":
                    OpenExecWindow("window.open('ChkEdit.aspx?ChkYear=" + ChkYear + "&DevNo=" + DevNo + "','_blank');");
                    break;
                case "ChkDev":
                    OpenExecWindow("window.open('../Device/DevEdit.aspx?DevNo=" + DevNo + "','_blank');"); 
                    break;
            }
        }
    }

    protected void OpenExecWindow(string strJavascript)    //選取後就另開視窗
    {
        if (ScriptManager.GetCurrent(Page) == null)
        {
            Page.ClientScript.RegisterStartupScript(Page.GetType(), "LifeLog", strJavascript, true);
        }
        else
        {
            ScriptManager.RegisterStartupScript(Page, Page.GetType(), "LifeLog", strJavascript, true);
        }
    }

    protected void SelPlace_SelectedIndexChanged(object sender, EventArgs e)
    {
        SelPointer.DataBind();
    }

    protected void SelSQL_SelectedIndexChanged(object sender, EventArgs e)
    {
        TextSQL.Text = SelSQL.SelectedValue;
    }

    protected void SelKeepUnit_SelectedIndexChanged(object sender, EventArgs e)
    {
        SelKeeper.DataBind();
    }

    protected void BtnExcel_Click(object sender, EventArgs e)
    {
        if (GridView1.Rows.Count > 0)
        {
            int pos = ViewState["SQL"].ToString().IndexOf("FROM");
            string ViewCols = SelShow.SelectedValue; //所有欲顯示欄位名稱

            GridView gvExport = new GridView();
            gvExport.DataSource = SqlDataSourceSearch;
            SqlDataSourceSearch.SelectCommand = "Select " + ViewCols + ViewState["SQL"].ToString().Substring(pos);
            gvExport.DataBind();

            string strExportFilename = "IDMS";

            Response.Clear();
            Response.ClearContent();
            Response.Write("<meta http-equiv=Content-Type content=text/html;charset=utf-8>");
            Response.AddHeader("content-disposition", "attachment;filename=" + strExportFilename + ".xls");
            Response.Cache.SetCacheability(HttpCacheability.NoCache);
            Response.ContentType = "application/vnd.xls";
            Response.Charset = "utf-8";

            System.IO.StringWriter stringWrite = new System.IO.StringWriter();
            System.Web.UI.HtmlTextWriter htmlWrite = new HtmlTextWriter(stringWrite);
            gvExport.RenderControl(htmlWrite);
            Response.Write(stringWrite.ToString().Replace("<div>", "").Replace("</div>", ""));
            Response.End();
        }
    }

    protected void SelPanel_SelectedIndexChanged(object sender, EventArgs e)
    {
        PanelChange();
    }

    protected void PanelChange()
    {
        SqlDataSourceCircuit.SelectCommand = "SELECT [設備編號],[設備名稱] FROM [實體設備],[接電] WHERE [實體設備].[設備編號]=[接電].[下游編號] and  [上游編號]=" + SelPanel.SelectedValue + " ORDER BY [設備種類],[設備名稱]";
        SelCircuit.DataBind();
    }

    protected string GetDbValue(string SQL) //取得某資料庫的值
    {
        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
        Conn.Open();
        SqlCommand cmd = new SqlCommand(SQL, Conn);
        SqlDataReader dr = null;
        dr = cmd.ExecuteReader();
        if (dr.Read()) return (dr[0].ToString());
        else return ("");
        cmd.Cancel(); dr.Close(); Conn.Close(); Conn.Dispose();
    }

    protected void ExecDbSQL(string SQL) //執行資料庫異動
    {
        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
        Conn.Open();
        SqlCommand cmd = new SqlCommand(SQL, Conn);
        cmd.ExecuteNonQuery();
        cmd.Cancel(); Conn.Close(); Conn.Dispose();
    }

    protected Boolean RightCheck() //是否有權修改資料
    {
        string UserName = Session["UserName"].ToString();
        if (InGroup(UserName, "環境小組")) return (true);
        else return (false);
    }
    protected Boolean InGroup(string ChkName, string ChkUnit) //檢查ChkName是否為ChkUnit成員或本身
    {
        if (ChkName == ChkUnit) return (true);  //是否同名
        else if (ChkUnit == "") return (false);	//檢查單位必填
        else //是否為成員UN (課別與小組同義)
        {
            SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
            Conn.Open();
            SqlCommand cmd = new SqlCommand("select * from Config where (Kind='" + ChkUnit + "') and Item='" + ChkName + "'", Conn);
            SqlDataReader dr = null;
            dr = cmd.ExecuteReader();
            if (dr.Read()) return (true);
            cmd.Cancel(); dr.Close(); Conn.Close(); Conn.Dispose();
        }

        return (false);
    }

    protected void BtnChkList_Click(object sender, EventArgs e) //產生或同步清查清單
    {
        Literal Msg = new Literal();

        if (!RightCheck()) Msg.Text = "<script>alert('您不是環境小組的成員，沒有產生或同步清查清單的權限！');</script>";        
        else
        {
            string strChkYear = SelChkYear.SelectedValue; if (strChkYear == "") strChkYear = DateTime.Now.ToString("yyyy-MM");
            string Checker = Session["UserName"].ToString();
            string ChkDT = DateTime.Now.ToString("yyyy/MM/dd HH:mm");            

            if (SelChkYear.SelectedValue != "" & SelChkYear.SelectedValue != TextConfirm.Text | SelChkYear.SelectedValue == "" & TextConfirm.Text != strChkYear) Msg.Text = "<script>alert('確認碼錯誤！');</script>";
            else
            {
                //將最近新增的實體設備資料加到清單
                SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
                Conn.Open();                
                SqlCommand cmd = new SqlCommand("select [設備編號],[定位編號],[維護人員] from [實體設備]"
                    + " where ([設備型態]='系統設備' or [設備型態]='網路設備' or [設備型態]='週邊設備' or [設備型態]='環境設備')"
                    + " and [定位編號] in (select [定位編號] from [定位設定] where [區域名稱] in (select [Item] from [Config] where [Kind]='門禁管制區'))" 
                    + " and [設備編號] not in (select [設備編號] from [機器清查] where [清查年度]='" + strChkYear + "')", Conn);
                SqlDataReader dr = null;
                dr = cmd.ExecuteReader();
                while (dr.Read())
                {
                    ExecDbSQL("insert into [機器清查] values('" + strChkYear + "'," + dr[0].ToString() + ",0,'','','','待清查','" + Checker + "','" + ChkDT + "','" + Checker + "','" + ChkDT + "')");
                }
                cmd.Cancel(); dr.Close(); Conn.Close(); Conn.Dispose();
                //刪除實體設備無資料(新增資料(DevNo<0)除外)或已非門禁管制區的清查資料
                ExecDbSQL("delete from [機器清查] where [清查年度]='" + strChkYear + "' and [設備編號]>0 and [設備編號] not in " 
                    + "(select [設備編號] from [View_設備管理] where [區域名稱] in " 
                    + "(select [Item] from [Config] where [Kind]='門禁管制區'))");

                Msg.Text = "<script>alert('已產生或同步 " + strChkYear + " 機器清查清單！');window.open('Check.aspx','_self');</script>";
            }
        }
        Page.Controls.Add(Msg);
    }

    protected void BtnChkDel_Click(object sender, EventArgs e) //刪除清查清單
    {
        Literal Msg = new Literal();
        string strChkYear = SelChkYear.SelectedValue;

        if (!RightCheck()) Msg.Text = "<script>alert('您不是環境小組的成員，沒有刪除清查清單的權限！');</script>";
        else
        {
            if (TextConfirm.Text != strChkYear) Msg.Text = "<script>alert('確認碼錯誤！');</script>";
            else Msg.Text = "<script>if(confirm('您確定要刪除" + strChkYear + "機器清查清單全部資料嗎？')) window.open('Check.aspx?Del=Y&ChkYear=" + strChkYear + "','_self');</script>";
        }
        
        Page.Controls.Add(Msg);
    }

    protected void BtnAddChk_Click(object sender, EventArgs e)  //新增清查資料
    {
        OpenExecWindow("window.open('ChkEdit.aspx?ChkYear=" + SelChkYear.SelectedValue + "','_blank');");
    }
}