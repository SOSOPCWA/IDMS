using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.Configuration;
using System.Data;
using System.Data.SqlClient;

public partial class Life_Check : System.Web.UI.Page
{
    const string DefSQL = "select * from [資料查核] where 1=1";
    const string OrderSQL = " order by [表格名稱],[主鍵編號]";

    protected void Page_Load(object sender, EventArgs e)
    {
        if (!IsPostBack)
        {
            ViewState["SQL"] = DefSQL + " and ([表格名稱]<>'軟體主檔' and [表格名稱]<>'軟體授權')" + OrderSQL;
            TextCount.Text = GetValue("SELECT count(*) from [資料查核] where [表格名稱]<>'軟體主檔' and [表格名稱]<>'軟體授權'");
        }
    }

    protected void Page_PreRenderComplete(object sender, EventArgs e)   //放Page_Load無法顯示
    {
        SqlDataSourceCheck.SelectCommand = ViewState["SQL"].ToString(); //查詢後，SQL要保持住，否則會用預設值
    }

    protected void SelUnit_SelectedIndexChanged(object sender, EventArgs e)
    {
        SelMt.DataBind();
    }

    protected void Button1_Click(object sender, EventArgs e)
    {
        if (ChkPage.Checked) GridView1.AllowPaging = true;
        else GridView1.AllowPaging = false;
        
        string SQL = DefSQL;

        switch (SelTbl.SelectedValue)
        {
            case "軟體除外": SQL = SQL + " and ([表格名稱]<>'軟體主檔' and [表格名稱]<>'軟體授權')"; break;
            case "設備+作業": SQL = SQL + " and ([表格名稱]='實體設備' or [表格名稱]='作業主機')"; break;            
            case "軟體管理": SQL = SQL + " and ([表格名稱]='軟體主檔' or [表格名稱]='軟體授權')"; break;
            case "(全部)": SQL = SQL + " and 2=2"; break;
            default: SQL = SQL + " and [表格名稱]='" + SelTbl.SelectedValue + "'"; break;
        }

        if (SelCheck.SelectedValue != "") SQL = SQL + " and [查核結果] like '%" + SelCheck.SelectedValue + "%'";

        string UnitSQL = "select [成員] from [View_組織架構] where [課別]='" + SelUnit.SelectedValue + "' or [成員] in (select [Item] from [Config] where [Kind]='" + SelUnit.SelectedValue + "')";
        if (ChkUnit.Checked) SQL = SQL + " and ([維護人員] in (" + UnitSQL + ") or [相關人員] in (" + UnitSQL + "))";
        if (ChkMt.Checked) SQL = SQL + " and ([維護人員]='" + SelMt.SelectedValue + "' or [維護人員] in (select Kind from Config where Item='" + SelMt.SelectedValue + "')"
            + " or [相關人員]='" + SelMt.SelectedValue + "' or [相關人員] in (select Kind from Config where Item='" + SelMt.SelectedValue + "'))";

        if (ChkKey.Checked & TextKey.Text != "")
        {
            string[] KeyA = TextKey.Text.Trim().Split(',');
            for (int i = 0; i < KeyA.GetLength(0); i++) SQL = SQL + " and ([資料內容] like '%" + KeyA[i] + "%' or [維護人員] like '%" + KeyA[i] + "%' or [相關人員] like '%" + KeyA[i] + "%' or [查核結果] like '%" + KeyA[i] + "%')";
        }

        //Response.Write(SQL);
        //Response.End();
        ViewState["SQL"] = SQL;
        SqlDataSourceCheck.SelectCommand = SQL + OrderSQL;
        GridView1.DataBind();

        TextCount.Text = GetValue("SELECT count(*) " + SQL.Substring(SQL.IndexOf("from")));
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

    protected void ExecDbSQL(string SQL) //執行資料庫異動
    {
        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
        Conn.Open();
        SqlCommand cmd = new SqlCommand(SQL, Conn);
        //Response.Write(cmd.CommandText);
        //Response.End();
        cmd.ExecuteNonQuery();
        cmd.Cancel(); cmd.Dispose(); Conn.Close(); Conn.Dispose();
    }

    protected void Button2_Click(object sender, EventArgs e)
    {
        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
        Conn.Open();
        SqlCommand cmd = new SqlCommand("select * from [View_資料查核] where [查核結果]<>''", Conn);
        cmd.CommandTimeout = 1000;
        SqlDataReader dr = null;

        ExecDbSQL("delete from [資料查核]");
        dr = cmd.ExecuteReader();
        while (dr.Read()) ExecDbSQL("insert into [資料查核] values('" + dr[0].ToString() + "'," + dr[1].ToString() + ",'" + dr[2].ToString() + "','" 
                + dr[3].ToString() + "','" + dr[4].ToString() + "','" + dr[5].ToString() + "')");

        dr.Close(); cmd.Cancel(); cmd.Dispose(); Conn.Close(); Conn.Dispose();

        GridView1.DataBind();

        Literal Msg = new Literal();
        Msg.Text = "<script>alert('暫存查核資料已更新!');</script>";
        if (Msg.Text != "") Page.Controls.Add(Msg);
    }

    protected void GridView1_SelectedIndexChanging(object sender, GridViewSelectEventArgs e)    //選取後就另開編輯視窗或帶入設備編號
    {
        string tblName = GridView1.Rows[e.NewSelectedIndex].Cells[1].Text;
        string PKno = GridView1.Rows[e.NewSelectedIndex].Cells[2].Text;
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
            Msg.Text = "<script>alert('" + tblName + "未提供編輯介面!');</script>";
            if (Msg.Text != "") Page.Controls.Add(Msg);
        }

    }

    protected void OpenExecWindow(string strJavascript)    //選取後就另開視窗
    {
        if (ScriptManager.GetCurrent(Page) == null)
        {
            Page.ClientScript.RegisterStartupScript(Page.GetType(), "Check", strJavascript, true);
        }
        else
        {
            ScriptManager.RegisterStartupScript(Page, Page.GetType(), "Check", strJavascript, true);
        }
    }

    protected void SelCheck_SelectedIndexChanged(object sender, EventArgs e)  //取得來源
    {
        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
        Conn.Open();
        SqlCommand cmd = new SqlCommand("select [Memo] from [Config] where [Kind]='查核項目' and [Item]='" + SelCheck.SelectedValue + "'", Conn);
        SqlDataReader dr = null;

        dr = cmd.ExecuteReader();
        lblCheck.Text = "";
        if (dr.Read()) lblCheck.Text = dr[0].ToString();

        dr.Close(); cmd.Cancel(); cmd.Dispose(); Conn.Close(); Conn.Dispose();
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
}