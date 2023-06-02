using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.Configuration;
using System.Data;
using System.Data.SqlClient;

public partial class Software_SwStatus : System.Web.UI.Page
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
            Button1_Click(null, null);
        }
    }

    protected void Button1_Click(object sender, EventArgs e)
    {
        if (ChkPage.Checked) GridView1.AllowPaging = true;
        else GridView1.AllowPaging = false;

        string SQL = "SELECT " + SelShow.SelectedValue + " FROM [軟體主檔] where [軟體狀態]='可使用'";
        if (TextKey.Text != "") SQL = SQL + " and [軟體名稱] like '%" + TextKey.Text + "%'";

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

    protected void GridView1_SelectedIndexChanging(object sender, GridViewSelectEventArgs e)
    {
        OpenExecWindow("window.open('SwEdit.aspx?SwNo=" + GridView1.DataKeys[e.NewSelectedIndex].Value.ToString() + "','_blank');");
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