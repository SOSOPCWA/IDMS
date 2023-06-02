using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.Configuration;
using System.Data;
using System.Data.SqlClient;

public partial class Ap_Sys : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        if (!IsPostBack)
        {
            
        }
    }

    protected void Page_PreRenderComplete(object sender, EventArgs e)   //放Page_Load無法顯示
    {
        if (ViewState["SQL"] != null) SqlDataSource1.SelectCommand = ViewState["SQL"].ToString();

        if (!IsPostBack)
        {
            if (Request["Search"] != null)  //外部關鍵字查詢
            {
                if (Request["Key"] != null)
                {
                    ChkKey.Checked = true;
                    TextKey.Text = Request["Key"].ToString();
                    BtnSearch_Click(null, null);
                }
            }
        }
    }

    protected void BtnSearch_Click(object sender, EventArgs e)
    {
        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
        Conn.Open();
        SqlCommand cmd = new SqlCommand("", Conn);
        SqlDataReader dr = null;

        string SQL = "SELECT * FROM [View_系統資源] WHERE 1=1";
        //---------------------------------------------------------分頁
        if (ChkPage.Checked) GridView1.AllowPaging = true;
        else GridView1.AllowPaging = false;
        //---------------------------------------------------------種類
        if (ChkKind.Checked) SQL = SQL + " and [資源種類]='" + SelKind.SelectedValue + "'";
        //---------------------------------------------------------人員
        string UnitSQL = "select distinct [成員] from [View_組織架構] where [課別]='" + SelUnit.SelectedValue + "'";
        if (ChkUnit.Checked) SQL = SQL + " and [維護人員] in (" + UnitSQL + ")";
        if (ChkGroup.Checked) SQL = SQL + " and [維護人員]='" + SelGroup.SelectedValue + "'";
        if (ChkMan.Checked) SQL = SQL + " and ([維護人員]='" + SelMan.SelectedValue + "' or [維護人員] in (select Kind from Config where Item='" + SelMan.SelectedValue + "'))";
        //---------------------------------------------------------資產編號
        if (ChkAssets.Checked) SQL = SQL + " and [資產編號]='" + SelAssets.SelectedValue + "'";
        //---------------------------------------------------------系統        
        if (ChkSysNo.Checked) SQL = SQL + " and [資源編號] in (" + SelSysNo.SelectedValue + "," + GetSysTree(SelSysNo.SelectedValue) + ")";
        //---------------------------------------------------------SQL、關鍵字
        if (ChkSQL.Checked & TextSQL.Text != "") SQL = SQL + " and (" + TextSQL.Text + ")";

        if (ChkKey.Checked)
        {
            string[] KeyA = TextKey.Text.Trim().Split(',');

            for (int i = 0; i < KeyA.GetLength(0); i++)
            {
                SQL = SQL + " and ([資源種類] like '%" + KeyA[i] + "%'"
                   + " or [資源名稱] like '%" + KeyA[i] + "%'"
                   + " or [資源功能] like '%" + KeyA[i] + "%'"
                   + " or [功能主機] like '%" + KeyA[i] + "%'"
                   + " or [維護人員] like '%" + KeyA[i] + "%'"
                   + " or [維護人員] in (select [Kind] from [Config] where [Item] like '%" + KeyA[i] + "%')"
                   + " or [備註說明] like '%" + KeyA[i] + "%'"
                   + ")";
            }
        }
        //---------------------------------------------------------整合SQL之再加工
        int pos = SQL.IndexOf("FROM");
        //---------------------------------------------------------總筆數(統計起始)   
        cmd.CommandText = "SELECT count(*) " + SQL.Substring(pos);
        dr = cmd.ExecuteReader();
        TextCount.Text = "0";
        try { if (dr.Read()) TextCount.Text = dr[0].ToString(); }
        catch { }
        dr.Close();
        //---------------------------------------------------------查詢顯示(統計結束)
        SQL = "Select * " + SQL.Substring(pos);
        SqlDataSource1.SelectCommand = SQL;
        GridView1.DataBind();        
        ViewState["SQL"] = SQL;
        //Response.Write(SQL);    
    }

    protected string GetSysTree(string SysNo)
    {
        string SysNos="";

        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
        Conn.Open();
        SqlCommand cmd = new SqlCommand("select [資源編號] from [系統資源] where [所屬編號]=" + SysNo, Conn);
        SqlDataReader dr = null;
        dr = cmd.ExecuteReader();

        while (dr.Read()) SysNos = SysNos + dr[0].ToString() + "," + GetSysTree(dr[0].ToString());

        cmd.Cancel(); cmd.Dispose(); dr.Close(); Conn.Close(); Conn.Dispose();

        return (FormatChilds(SysNos));
    }

    protected string FormatChilds(string PkNos)    //格式化下游字串
    {
        string cfg = PkNos.Replace(",0,", ",").Replace(",,", ",");
        if (cfg == "" | cfg == ",") cfg = "0";
        if (cfg.Substring(cfg.Length - 1) == ",") cfg = cfg.Substring(0, cfg.Length - 1);
        if (cfg.Substring(0, 1) == ",") cfg = cfg.Substring(1);

        return (cfg);
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

    protected void GridView1_RowCommand(object sender, CommandEventArgs e)
    {
        if (e.CommandName == "編號")
        {
            int Idx = int.Parse(e.CommandArgument.ToString());
            string SysNo = GridView1.DataKeys[Idx].Value.ToString(); 
            OpenExecWindow("window.open('SysEdit.aspx?SysNo=" + SysNo + "','_blank');");
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

    protected void BtnExcel_Click(object sender, EventArgs e)
    {
        if (GridView1.Rows.Count > 0)
        {
            GridView gvExport = new GridView();
            gvExport.DataSource = SqlDataSource1;
            SqlDataSource1.SelectCommand = ViewState["SQL"].ToString();
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

    protected void AddMsg(string strMsg)
    {
        Literal Msg = new Literal();
        Msg.Text = strMsg;
        Page.Controls.Add(Msg);
    }
}