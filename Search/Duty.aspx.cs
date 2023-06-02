using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.Configuration;
using System.Data;
using System.Data.SqlClient;

public partial class Search_Duty : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        if (!IsPostBack)
        {
            
        }
    }

    protected void Page_PreRenderComplete(object sender, EventArgs e)   //放Page_Load無法顯示
    {
        if (ViewState["DevSQL"] != null) SqlDataSourceDev.SelectCommand = ViewState["DevSQL"].ToString();
        if (ViewState["ApSQL"] != null) SqlDataSourceAp.SelectCommand = ViewState["ApSQL"].ToString();
        if (ViewState["SysSQL"] != null) SqlDataSourceSys.SelectCommand = ViewState["SysSQL"].ToString();
        if (ViewState["ChkSQL"] != null) SqlDataSourceChk.SelectCommand = ViewState["ChkSQL"].ToString();

        if (!IsPostBack)
        {
            string UnitName = Session["UnitName"].ToString(); if (UnitName == null) UnitName = "";
            string UserName = Session["UserName"].ToString(); if (UserName == null) UserName = "";
            
            for (int i = 0; i < SelUnit.Items.Count; i++) if (SelUnit.Items[i].Value == UnitName) SelUnit.SelectedIndex = i; 
            SelMan.DataBind();
            for (int i = 0; i < SelMan.Items.Count; i++) if (SelMan.Items[i].Value == UserName) SelMan.SelectedIndex = i; 
            
            BtnSearch_Click(null, null);
        }
    }

    protected void BtnSearch_Click(object sender, EventArgs e)
    {
        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
        Conn.Open();
        SqlCommand cmd = new SqlCommand("", Conn);
        SqlDataReader dr = null;

        string DevSQL = "SELECT * FROM [View_設備管理] WHERE 1=1";
        string ApSQL = "SELECT * FROM [View_作業主機] WHERE 1=1";
        string SysSQL = "SELECT * FROM [View_系統資源] WHERE 1=1";
        string ChkSQL = "SELECT * FROM [資料查核] WHERE 1=1";
        //---------------------------------------------------------分頁
        if (ChkPage.Checked)
        {
            GridViewDev.AllowPaging = true;
            GridViewAp.AllowPaging = true;
            GridViewSys.AllowPaging = true;
            GridViewChk.AllowPaging = true;
        }
        else
        {
            GridViewDev.AllowPaging = false;
            GridViewAp.AllowPaging = false;
            GridViewSys.AllowPaging = false;
            GridViewChk.AllowPaging = false;
        }
        //---------------------------------------------------------人員
        string UserName = SelMan.SelectedValue;
        string MtSQL = "[維護人員] = '" + UserName + "' or [維護人員] in (select [Kind] from [Config] where [Item]='" + UserName + "')";
        string KtSQL = "[設備保管人員] = '" + UserName + "' or [財產保管人員] = '" + UserName + "'";
        string AtSQL = "[相關人員] = '" + UserName + "' or [相關人員] in (select [Kind] from [Config] where [Item]='" + UserName + "')";

        DevSQL = DevSQL + " and (" + MtSQL + " or " + KtSQL + ")";
        ApSQL  = ApSQL  + " and (" + MtSQL + ")";
        SysSQL = SysSQL + " and (" + MtSQL + ")";
        ChkSQL = ChkSQL + " and (" + MtSQL + " or " + AtSQL + ")";
        //---------------------------------------------------------整合SQL之再加工
        int DevPos = DevSQL.IndexOf("FROM");
        int ApPos = ApSQL.IndexOf("FROM");
        int SysPos = SysSQL.IndexOf("FROM");
        int ChkPos = ChkSQL.IndexOf("FROM");
        //---------------------------------------------------------總筆數(統計起始)   
        cmd.CommandText = "SELECT count(*) " + DevSQL.Substring(DevPos);
        dr = cmd.ExecuteReader();
        TextDevCount.Text = "0";
        try { if (dr.Read()) TextDevCount.Text = dr[0].ToString(); }
        catch { }
        dr.Close();

        cmd.CommandText = "SELECT count(*) " + ApSQL.Substring(ApPos);
        dr = cmd.ExecuteReader();
        TextApCount.Text = "0";
        try { if (dr.Read()) TextApCount.Text = dr[0].ToString(); }
        catch { }
        dr.Close();

        cmd.CommandText = "SELECT count(*) " + SysSQL.Substring(SysPos);
        dr = cmd.ExecuteReader();
        TextSysCount.Text = "0";
        try { if (dr.Read()) TextSysCount.Text = dr[0].ToString(); }
        catch { }
        dr.Close();

        cmd.CommandText = "SELECT count(*) " + ChkSQL.Substring(ChkPos);
        dr = cmd.ExecuteReader();
        TextChkCount.Text = "0";
        try { if (dr.Read()) TextChkCount.Text = dr[0].ToString(); }
        catch { }
        dr.Close();
        //---------------------------------------------------------查詢顯示(統計結束)
        DevSQL = "Select * " + DevSQL.Substring(DevPos);
        SqlDataSourceDev.SelectCommand = DevSQL;
        GridViewDev.DataBind();        
        ViewState["DevSQL"] = DevSQL;

        ApSQL = "Select * " + ApSQL.Substring(ApPos);
        SqlDataSourceAp.SelectCommand = ApSQL;
        GridViewAp.DataBind();
        ViewState["ApSQL"] = ApSQL;

        SysSQL = "Select * " + SysSQL.Substring(SysPos);
        SqlDataSourceSys.SelectCommand = SysSQL;
        GridViewSys.DataBind();
        ViewState["SysSQL"] = SysSQL;

        ChkSQL = "Select * " + ChkSQL.Substring(ChkPos);
        SqlDataSourceChk.SelectCommand = ChkSQL;
        GridViewChk.DataBind();
        lblGroup.Text = GetValues("select [Kind] from [Config] where [Item]='" + SelMan.SelectedValue + "'");
        ViewState["ChkSQL"] = ChkSQL; 
    }

    protected void BtnRepeat_Click(object sender, EventArgs e)
    {
        OpenExecWindow("window.open('../Life/Repeat.aspx?Mt=" + SelMan.SelectedValue + "','_blank');");
    }
    protected void BtnSw_Click(object sender, EventArgs e)
    {
        OpenExecWindow("window.open('../Software/SwUser.aspx?Mt=" + SelMan.SelectedValue + "','_blank');");
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
    protected string GetValues(string SQL)   //取得多重資料
    {
        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
        Conn.Open();
        SqlCommand cmd = new SqlCommand(SQL, Conn);
        SqlDataReader dr = null;
        dr = cmd.ExecuteReader();
        string cfg = ""; while (dr.Read()) cfg = cfg + dr[0].ToString() + "、";

        cmd.Cancel(); cmd.Dispose(); dr.Close(); Conn.Close(); Conn.Dispose();

        return (FormatChilds(cfg));
    }
    protected string FormatChilds(string cfg)    //格式化下游字串
    {
        if (cfg.Substring(cfg.Length - 1) == "、") cfg = cfg.Substring(0, cfg.Length - 1);
        if (cfg.Substring(0, 1) == "、") cfg = cfg.Substring(1);

        return (cfg);
    }  
    protected void AddMsg(string strMsg)
    {
        Literal Msg = new Literal();
        Msg.Text = strMsg;
        Page.Controls.Add(Msg);
    }

    protected void GridViewDev_RowCommand(object sender, CommandEventArgs e)
    {
        if (e.CommandName == "編號")
        {
            int Idx = int.Parse(e.CommandArgument.ToString());
            string PkNo = GridViewDev.DataKeys[Idx].Value.ToString(); 
            OpenExecWindow("window.open('../Device/DevEdit.aspx?DevNo=" + PkNo + "','_blank');");
        }
    }
    protected void GridViewAp_RowCommand(object sender, CommandEventArgs e)
    {
        if (e.CommandName == "編號")
        {
            int Idx = int.Parse(e.CommandArgument.ToString());
            string PkNo = GridViewAp.DataKeys[Idx].Value.ToString();
            OpenExecWindow("window.open('../Ap/ApEdit.aspx?ApNo=" + PkNo + "','_blank');");
        }
    }
    protected void GridViewSys_RowCommand(object sender, CommandEventArgs e)
    {
        if (e.CommandName == "編號")
        {
            int Idx = int.Parse(e.CommandArgument.ToString());
            string PkNo = GridViewSys.DataKeys[Idx].Value.ToString();
            OpenExecWindow("window.open('../Ap/SysEdit.aspx?SysNo=" + PkNo + "','_blank');");
        }
    }
    protected void GridViewChk_SelectedIndexChanging(object sender, GridViewSelectEventArgs e)    //選取後就另開編輯視窗或帶入設備編號
    {
        string tblName = GridViewChk.Rows[e.NewSelectedIndex].Cells[1].Text;
        string PKno = GridViewChk.Rows[e.NewSelectedIndex].Cells[2].Text;
        string url = "";

        switch (tblName)
        {
            case "實體設備": url = "../Device/DevEdit.aspx?DevNo=" + PKno; break;
            case "作業主機": url = "../AP/ApEdit.aspx?ApNo=" + PKno; break;
            case "接網迴路": url = "../Device/TreeEdit.aspx?DevNo=" + PKno; break;
            case "系統資源": url = "../AP/SysEdit.aspx?SysNo=" + PKno; break;
            case "軟體主檔": url = "../SoftWare/SwEdit.aspx?SwNo=" + PKno; break;
            case "軟體授權": url = "../SoftWare/AskEdit.aspx?AskNo=" + PKno; break;
            case "秘總財產": url = ""; break;
            default: url = ""; break;
        }

        if (url != "") OpenExecWindow("window.open('" + url + "','_blank');");
        else
        {
            Literal Msg = new Literal();
            Msg.Text = "<script>alert('" + "未提供編輯介面!');</script>";
            if (Msg.Text != "") Page.Controls.Add(Msg);
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

    protected void BtnDevExcel_Click(object sender, EventArgs e)
    {
        if (GridViewDev.Rows.Count > 0)
        {
            GridView gvExport = new GridView();
            gvExport.DataSource = SqlDataSourceDev;
            SqlDataSourceDev.SelectCommand = ViewState["DevSQL"].ToString();
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
    protected void BtnApExcel_Click(object sender, EventArgs e)
    {
        if (GridViewAp.Rows.Count > 0)
        {
            GridView gvExport = new GridView();
            gvExport.DataSource = SqlDataSourceAp;
            SqlDataSourceAp.SelectCommand = ViewState["ApSQL"].ToString();
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
    protected void BtnSysExcel_Click(object sender, EventArgs e)
    {
        if (GridViewSys.Rows.Count > 0)
        {
            GridView gvExport = new GridView();
            gvExport.DataSource = SqlDataSourceSys;
            SqlDataSourceSys.SelectCommand = ViewState["SysSQL"].ToString();
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
    protected void BtnChkExcel_Click(object sender, EventArgs e)
    {
        if (GridViewChk.Rows.Count > 0)
        {
            GridView gvExport = new GridView();
            gvExport.DataSource = SqlDataSourceChk;
            SqlDataSourceChk.SelectCommand = ViewState["ChkSQL"].ToString();
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
    
}