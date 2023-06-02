using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.Configuration;
using System.Data;
using System.Data.SqlClient;

public partial class Software_SwName : System.Web.UI.Page
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
            SelSwName_SelectedIndexChanged(null, null);
        }
    }

    protected void Button1_Click(object sender, EventArgs e)
    {
        string SQL = "SELECT " + SelShow.SelectedValue + " FROM [View_軟體管理] WHERE 1=1";
        if (ChkPage.Checked) GridView1.AllowPaging = true;
        else GridView1.AllowPaging = false;

        SQL = SQL + " and [軟體編號]=" + SelSwName.SelectedValue;      
        
        int pos = SQL.IndexOf("FROM");
        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
        Conn.Open();
        SqlCommand cmd = new SqlCommand("SELECT count(*) " + SQL.Substring(pos), Conn);
        SqlDataReader dr = cmd.ExecuteReader();
                
        TextCount.Text = "0";
        try { if (dr.Read()) TextCount.Text = dr[0].ToString(); }
        catch { }

        cmd.Cancel(); cmd.Dispose(); dr.Close(); Conn.Close(); Conn.Dispose();

        //Response.Write(SQL);
        //Response.End();
        SqlDataSource1.SelectCommand = SQL;
        GridView1.DataBind();
        ViewState["SQL"] = SQL;
    }

    protected void LinkKey_Click(object sender, EventArgs e)
    {
        string SQL = "Select [軟體編號],'('+[表單編號]+') '+[軟體名稱] as [軟體名稱] from [軟體主檔] where [軟體狀態]='可使用'";
        if (TextKey.Text == "") SqlDataSourceSwName.SelectCommand = SQL + " order by [軟體名稱]";
        else SqlDataSourceSwName.SelectCommand = SQL + " and [軟體名稱] like '%" + TextKey.Text + "%' order by [軟體名稱]";
    }

    protected void LinkEdit_Click(object sender, EventArgs e)
    {
        OpenExecWindow("window.open('SwEdit.aspx?SwNo=" + SelSwName.SelectedValue + "','_blank');");
    }

    protected void SelSwName_SelectedIndexChanged(object sender, EventArgs e)   //顯示系統功能之說明
    {
        string txt = "'表單編號：'+[表單編號] +'<br />"
            + "軟體名稱：'+[軟體名稱] +'<br />"
            + "購買版本：'+[購買版本] +'<br />"
            + "授權方式：'+[授權方式] +'<br />"
            + "授權數量：'+CONVERT(varchar,[授權數量]) +'<br />"
            + "軟體功能：'+[軟體功能] +'<br />"
            + "購買日期：'+CONVERT(char(10),[購買日期],111) +'<br />"
            + "提供單位：'+[提供單位] +'<br />"
            + "圖書編號：'+[圖書編號] +'<br />"
            + "軟體狀態：'+[軟體狀態]";
        lblSW.Text = GetValue("select " + txt + " from [軟體主檔] where [軟體編號]=" + SelSwName.SelectedValue);
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
        OpenExecWindow("window.open('AskEdit.aspx?AskNo=" + GridView1.DataKeys[e.NewSelectedIndex].Value.ToString() + "','_blank');");
    }

    protected void OpenExecWindow(string strJavascript)    //選取後就另開視窗
    {
        if (ScriptManager.GetCurrent(Page) == null)
        {
            Page.ClientScript.RegisterStartupScript(Page.GetType(), "SwSearch", strJavascript, true);
        }
        else
        {
            ScriptManager.RegisterStartupScript(Page, Page.GetType(), "SwSearch", strJavascript, true);
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
}