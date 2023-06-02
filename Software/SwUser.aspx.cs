using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.Configuration;
using System.Data;
using System.Data.SqlClient;

public partial class Software_SwUser : System.Web.UI.Page
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
            string Unit = "", Mt = "";
            if (Request["Mt"] != null)  //外部人員帶入查詢
            {
                Mt = Request["Mt"].ToString();
                Unit = GetValue("IDMS", "select [課別] from [View_員工] where [成員]='" + Mt + "'");
                for (int i = 0; i < SelUnitLcs.Items.Count; i++)
                {
                    if (SelUnitLcs.Items[i].Value != Unit) SelUnitLcs.Items[i].Selected = false;
                    else SelUnitLcs.Items[i].Selected = true;
                }
            }
            SelUserLcs.DataBind();
            for (int i = 0; i < SelUserLcs.Items.Count; i++)
            {
                if (SelUserLcs.Items[i].Value != Mt) SelUserLcs.Items[i].Selected = false;
                else
                {
                    SelUserLcs.Items[i].Selected = true;
                    ChkUserLcs.Checked = true;
                    Button1_Click(null, null);
                }
            }
        }
    }       

    protected void Button1_Click(object sender, EventArgs e)
    {
        string SQL = "SELECT " + SelShow.SelectedValue + " FROM [View_軟體管理] WHERE 1=1";
        if (ChkPage.Checked) GridView1.AllowPaging = true;
        else GridView1.AllowPaging = false;

        if (ChkUnitLcs.Checked) SQL = SQL + " and [授權單位]='" + SelUnitLcs.SelectedValue + "'";
        if (ChkUserLcs.Checked) SQL = SQL + " and [授權人員]='" + SelUserLcs.SelectedValue + "'";        
        
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

    protected string GetValue(string DB, string SQL)   //取得單一資料
    {
        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings[DB + "ConnectionString"].ConnectionString);
        Conn.Open();
        SqlCommand cmd = new SqlCommand(SQL, Conn);
        SqlDataReader dr = null;
        dr = cmd.ExecuteReader();
        string cfg = ""; if (dr.Read()) cfg = dr[0].ToString();
        cmd.Cancel(); cmd.Dispose(); dr.Close(); Conn.Close(); Conn.Dispose();

        return (cfg);
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