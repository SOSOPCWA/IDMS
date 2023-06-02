using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.Configuration;
using System.Data;
using System.Data.SqlClient;

public partial class Search_Discard : System.Web.UI.Page
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
        OpenSearch();
    }

    protected void Button1_Click(object sender, EventArgs e)
    {
        OpenSearch();
    }

    protected void OpenSearch()
    {
        if (ChkPage.Checked) GridView1.AllowPaging = true;
        else GridView1.AllowPaging = false;
        
        string SQL = "SELECT * FROM [實體設備],[定位設定] WHERE [實體設備].[定位編號]=[定位設定].[定位編號] and [設備狀態]='已報廢'";

        switch (SelDiscard.SelectedValue) 
        {
            case "No"  : SQL = SQL + " and [財產編號A]+' '+[財產編號B] not in (select [財產編號A]+' '+[財產編號B] from [財產主檔] where [無效註記]='N')"; break;
            case "Yes" : SQL = SQL + " and [財產編號A]+' '+[財產編號B] in (select [財產編號A]+' '+[財產編號B] from [財產主檔] where [無效註記]='N')";break;
        }
        SQL = SQL + " and " + GetKeySQL().ToString();
        
        int pos = SQL.IndexOf("FROM");
        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
        Conn.Open();
        SqlCommand cmd = new SqlCommand("SELECT sum([價值]) from [實體設備] where [設備編號] in (select [設備編號] " + SQL.Substring(pos) + ")", Conn);
        //Response.Write("SELECT sum([價值]) from [實體設備] where [設備編號] in (select [設備編號] " + SQL.Substring(pos) + ")");
        //Response.End();
        SqlDataReader dr = null;
        dr = cmd.ExecuteReader();
        try { if (dr.Read()) TextTotal.Text = String.Format("{0:C0}", dr[0]); }
        catch { }
        cmd.Cancel(); cmd.Dispose(); dr.Close(); Conn.Close(); Conn.Dispose();

        string ViewCols = SelShow.SelectedValue; //所有欲顯示欄位名稱
        SQL = "Select " + ViewCols + SQL.Substring(pos);
        SqlDataSource1.SelectCommand = SQL;
        GridView1.DataBind();
        ViewState["SQL"] = SQL;
    }

    protected void GridView1_SelectedIndexChanging(object sender, GridViewSelectEventArgs e)
    {
        OpenExecWindow("window.open('../Device/DevEdit.aspx?DevNo=" + GridView1.DataKeys[e.NewSelectedIndex].Value.ToString() + "','_blank');");
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
            int pos = ViewState["SQL"].ToString().IndexOf("FROM");
            string ViewCols = SelShow.SelectedValue; //所有欲顯示欄位名稱

            GridView gvExport = new GridView();
            gvExport.DataSource = SqlDataSource1;
            SqlDataSource1.SelectCommand = "Select " + ViewCols + ViewState["SQL"].ToString().Substring(pos);
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

    protected string GetKeySQL()
    {
        string[] KeyA = TextKey.Text.Trim().Split(',');
        string SQL = "";

        for (int i = 0; i < KeyA.GetLength(0); i++)
        {
            if (i > 0) SQL = SQL + " and ";

            SQL = SQL + "([設備名稱] like '%" + KeyA[i] + "%'"
            + " or [設備用途] like '%" + KeyA[i] + "%'"
            + " or [財產編號A] like '%" + KeyA[i] + "%'"
            + " or [財產編號B] like '%" + KeyA[i] + "%'"
            + " or [財產名稱] like '%" + KeyA[i] + "%'"
            + " or [財產別名] like '%" + KeyA[i] + "%'"
            + " or [取得來源] like '%" + KeyA[i] + "%'"
            + " or [規格] like '%" + KeyA[i] + "%'"
            + " or [型式] like '%" + KeyA[i] + "%'"
            + " or [硬體序號] like '%" + KeyA[i] + "%'"
            + " or [iLoIP] like '%" + KeyA[i] + "%'"
            + " or [備註說明] like '%" + KeyA[i] + "%'"
            + " or [設備種類] like '%" + KeyA[i] + "%'"
            + " or [財產編號A] like '%" + KeyA[i] + "%'"
            + " or [財產編號B] like '%" + KeyA[i] + "%'"
            + " or [廠牌] like '%" + KeyA[i] + "%'"
            + " or [數量單位] like '%" + KeyA[i] + "%'"
            + " or [保管人員] like '%" + KeyA[i] + "%'"
            + " or [區域名稱]+[定位名稱] like '%" + KeyA[i] + "%'"
            + " or [維護廠商] like '%" + KeyA[i] + "%'"
            + " or [維護廠商] in (select [Item] from [Config] where [Kind]='維護廠商' and [Config] like '%" + KeyA[i] + "%')"
            + " or [維護人員] like '%" + KeyA[i] + "%'"
            + " or [維護人員] in (select [Kind] from [Config] where [Item] like '%" + KeyA[i] + "%')"
            + " or [設備狀態] like '%" + KeyA[i] + "%'"
            + " or [建立人員] like '%" + KeyA[i] + "%'"
            + " or [修改人員] like '%" + KeyA[i] + "%'"
            + " or [設備編號] in (select [設備編號] from [作業主機] where "
            + " [主機名稱] like '%" + KeyA[i]
            + "%' or [IP] like '%" + KeyA[i]
            + "%' or [緊急程度] like '%" + KeyA[i] + "%')"
            + ")";
        }

        return (SQL);
    }
}