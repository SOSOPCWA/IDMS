using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.Configuration;
using System.Data;
using System.Data.SqlClient;

public partial class Ap_Report : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        if (!IsPostBack)
        {
            lblDT.Text = (DateTime.Now).ToString("yyyy/MM/dd HH:mm");
        }

        //以下不能放在!IsPostBack裡面，否則不能排序
        string DevNos, ApNos, SysNos, ChkSQL="";

        lblSearch.Text = Session["Search"].ToString();
        DevNos = Session["DevNos"].ToString(); if (DevNos == "") DevNos = "0";
        ApNos = Session["ApNos"].ToString(); if (ApNos == "") ApNos = "0";
        SysNos = Session["SysNos"].ToString(); if (SysNos == "") SysNos = "0";
        if (Session["ChkSQL"] != null) ChkSQL = Session["ChkSQL"].ToString();

        SqlDataSourceDev.SelectCommand = "select [設備編號],[設備名稱],[設備種類],[設備用途],[放置地點],[設備狀態],[保管人員],[維護人員] from [View_設備管理]"
            + " where [設備編號] in (" + DevNos + ")" + ChkSQL + " order by [設備型態],[設備種類],[設備名稱]";
        
        SqlDataSourceAp.SelectCommand = "select [作業編號],[主機名稱],[系統名稱],[主要功能],[作業平台],[IP],[緊急程度],[作業狀態],[維護人員] from [View_作業主機]"
            + " where [作業編號] in (" + ApNos + ")";
        if(ChkSQL.IndexOf("[設備狀態]='已上線'")>0) SqlDataSourceAp.SelectCommand=SqlDataSourceAp.SelectCommand + " and [作業狀態]='已上線'";
        SqlDataSourceAp.SelectCommand = SqlDataSourceAp.SelectCommand + " order by [系統名稱],[主機名稱]";

        SqlDataSourceSys.SelectCommand = "select [資源編號],[資源名稱],[資源種類],[資源功能],[功能主機],[上游系統],[下游系統],[維護人員],[備註說明] from [View_系統資源]"
            + " where [資源種類]<>'分類' and [資源編號] in (" + SysNos + ")" + " order by [系統全名]";

        GridViewDev.DataBind();
        GridViewAp.DataBind();
        GridViewSys.DataBind();
    }

    protected void BtnDevExcel_Click(object sender, EventArgs e)
    {
        Excel_Out(GridViewDev, SqlDataSourceDev, "IDMS(實體設備)");
    }
    protected void BtnApExcel_Click(object sender, EventArgs e)
    {
        Excel_Out(GridViewAp, SqlDataSourceAp, "IDMS(作業主機)");
    }
    protected void BtnSysExcel_Click(object sender, EventArgs e)
    {
        Excel_Out(GridViewSys, SqlDataSourceSys, "IDMS(系統資源)");
    }

    protected void Excel_Out(GridView gd, SqlDataSource ds, string file)
    {
        if (gd.Rows.Count > 0)
        {
            GridView gvExport = new GridView();
            gvExport.DataSource = ds;
            gvExport.DataBind();

            string strExportFilename = file;

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