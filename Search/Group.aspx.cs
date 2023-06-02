using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.Configuration;
using System.Data;
using System.Data.SqlClient;

public partial class Search_Group : System.Web.UI.Page
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
        string SQL = "select [Kind] as [群組],[Item] as [維護小組],[Config] as [課別]," 
            + " (select [Item] + ',' from [Config] WHERE [Kind]=[Config1].[Item] order by [Mark] FOR XML PATH('')) AS [成員]" 
            + " from [Config] as [Config1] where [Kind] in (select [Item] from [Config] where [Kind]='維護群組')";

        if (SelGroup.SelectedValue != "") SQL = SQL + " and [Kind]='" + SelGroup.SelectedValue + "'";
        SQL = SQL + " order by [群組],[維護小組]";

        SqlDataSource1.SelectCommand = SQL;
        GridView1.DataBind();
        ViewState["SQL"] = SQL;
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