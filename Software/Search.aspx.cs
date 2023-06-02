using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.Configuration;
using System.Data;
using System.Data.SqlClient;

public partial class Software_Search : System.Web.UI.Page
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
            
        }
    }

    protected void Button1_Click(object sender, EventArgs e)
    {
        string SQL = "SELECT * FROM [View_軟體管理] WHERE 1=1";
        if (ChkPage.Checked) GridView1.AllowPaging = true;
        else GridView1.AllowPaging = false;

        if (ChkUnitSw.Checked) SQL = SQL + " and [軟體單位]='" + SelUnitSw.SelectedValue + "'";
        if (ChkUserSw.Checked) SQL = SQL + " and [軟體人員]='" + SelUserSw.SelectedValue + "'";
        if (ChkUnitLcs.Checked) SQL = SQL + " and [授權單位]='" + SelUnitLcs.SelectedValue + "'";
        if (ChkUserLcs.Checked) SQL = SQL + " and [授權人員]='" + SelUserLcs.SelectedValue + "'";
        if (ChkUnitUse.Checked) SQL = SQL + " and [使用單位]='" + SelUnitUse.SelectedValue + "'";
        if (ChkUserUse.Checked) SQL = SQL + " and [使用人員]='" + SelUserUse.SelectedValue + "'";

        if (ChkLcsBuy.Checked) SQL = SQL + " and [購買授權]='" + SelLcsBuy.SelectedValue + "'";
        if (ChkLcsAsk.Checked) SQL = SQL + " and [申請授權]='" + SelLcsAsk.SelectedValue + "'";
        if (ChkSwStatus.Checked) SQL = SQL + " and [軟體狀態]='" + SelSwStatus.SelectedValue + "'";
        if (ChkAskStatus.Checked) SQL = SQL + " and [授權狀態]='" + SelAskStatus.SelectedValue + "'";
        if (ChkDevStatus.Checked) SQL = SQL + " and [設備狀態]='" + SelDevStatus.SelectedValue + "'";
        if (ChkApStatus.Checked) SQL = SQL + " and [作業狀態]='" + SelApStatus.SelectedValue + "'";

        string strNotEQ = "";
        if (ChkApNo.Checked) strNotEQ = strNotEQ + " or " + "[授權作業編號]<>[作業編號] or ([授權作業編號]<>0 and [作業編號] is null)";                
        if (ChkUnit.Checked) strNotEQ = strNotEQ + " or " + "[授權單位]<>[使用單位] or ([作業編號]>0 and [授權單位]<>'' and [使用單位] is null)";
        if (ChkUser.Checked) strNotEQ = strNotEQ + " or " + "[授權人員]<>[使用人員] and ([授權人員] not in (Select [Item] from [Config] where [Kind]=[使用人員])) or ([作業編號]>0 and [授權人員]<>'' and [使用人員] is null)";
        if (ChkHost.Checked) strNotEQ = strNotEQ + " or " + "[授權主機]<>[使用主機] or ([授權主機]<>'' and [使用主機] is null)";
        if (ChkIP.Checked) strNotEQ = strNotEQ + " or " + "[授權IP]<>[使用IP] or ([授權IP]<>'' and [使用IP] is null)";
        if (ChkProp.Checked) strNotEQ = strNotEQ + " or " + "[授權財編A]<>[使用財編A] or [授權財編B]<>[使用財編B] or ([授權財編A]<>'' and [使用財編A] is null)";
        if (ChkBrand.Checked) strNotEQ = strNotEQ + " or " + "[授權廠牌]<>[使用廠牌] or ([授權廠牌]<>'' and [使用廠牌] is null)";
        if (ChkStyle.Checked) strNotEQ = strNotEQ + " or " + "[授權型式]<>[使用型式] or ([授權型式]<>'' and [使用型式] is null)";
        if (strNotEQ != "") SQL = SQL + " and (" + strNotEQ.Substring(3) + ")";

        if (ChkSQL.Checked & TextSQL.Text != "") SQL = SQL + " and (" + TextSQL.Text + ")";

        string strKey = "";
        if (ChkSwKey.Checked & TextKey.Text != "")   //軟體主檔
        {
            string[] KeyA = TextKey.Text.Trim().Split(',');

            for (int i = 0; i < KeyA.GetLength(0); i++)
            {
                strKey = strKey + " or ([表單編號] like '%" + KeyA[i] + "%'"
                   + " or [軟體單位] like '%" + KeyA[i] + "%'"
                   + " or [軟體人員] like '%" + KeyA[i] + "%'"
                   + " or [軟體名稱] like '%" + KeyA[i] + "%'"
                   + " or [購買版本] like '%" + KeyA[i] + "%'"
                   + " or [可用版本] like '%" + KeyA[i] + "%'"
                   + " or [購買授權] like '%" + KeyA[i] + "%'"
                   + " or [授權說明] like '%" + KeyA[i] + "%'"
                   + " or [軟體功能] like '%" + KeyA[i] + "%'"
                   + " or [適用廠牌] like '%" + KeyA[i] + "%'"
                   + " or [適用型式] like '%" + KeyA[i] + "%'"
                   + " or [購買序號] like '%" + KeyA[i] + "%'"
                   + " or [降級序號] like '%" + KeyA[i] + "%'"
                   + " or [期限說明] like '%" + KeyA[i] + "%'"
                   + " or [軟體來源] like '%" + KeyA[i] + "%'"
                   + " or [價格說明] like '%" + KeyA[i] + "%'"
                   + " or [提供單位] like '%" + KeyA[i] + "%'"
                   + " or [存放媒體] like '%" + KeyA[i] + "%'"
                   + " or [圖書編號] like '%" + KeyA[i] + "%'"
                   + " or [軟體附件] like '%" + KeyA[i] + "%'"
                   + " or [軟體狀態] like '%" + KeyA[i] + "%'"
                   + " or [軟體減損原因] like '%" + KeyA[i] + "%'"
                   + " or [軟體減損方式] like '%" + KeyA[i] + "%'"                   
                   + " or [軟體建立人員] like '%" + KeyA[i] + "%'"
                   + " or [軟體修改人員] like '%" + KeyA[i] + "%'"
                   + " or [軟體備註說明] like '%" + KeyA[i] + "%'"
                   + ")";
            }
        }
        if (ChkAskKey.Checked & TextKey.Text != "")  //軟體授權
        {
            string[] KeyB = TextKey.Text.Trim().Split(',');

            for (int i = 0; i < KeyB.GetLength(0); i++)
            {
                strKey = strKey + " or ([授權識別] like '%" + KeyB[i] + "%'"
                   + " or [申請版本] like '%" + KeyB[i] + "%'"
                   + " or [申請授權] like '%" + KeyB[i] + "%'"
                   + " or [申請序號] like '%" + KeyB[i] + "%'"
                   + " or [授權單位] like '%" + KeyB[i] + "%'"
                   + " or [授權人員] like '%" + KeyB[i] + "%'"
                   + " or [授權主機] like '%" + KeyB[i] + "%'"
                   + " or [授權IP] like '%" + KeyB[i] + "%'"
                   + " or [授權財編A] like '%" + KeyB[i] + "%'"
                   + " or [授權財編B] like '%" + KeyB[i] + "%'"
                   + " or [授權廠牌] like '%" + KeyB[i] + "%'"
                   + " or [授權型式] like '%" + KeyB[i] + "%'"
                   + " or [授權附件] like '%" + KeyB[i] + "%'"
                   + " or [授權狀態] like '%" + KeyB[i] + "%'"
                   + " or [授權減損原因] like '%" + KeyB[i] + "%'"
                   + " or [授權減損方式] like '%" + KeyB[i] + "%'"
                   + " or [授權填表人員] like '%" + KeyB[i] + "%'"
                   + " or [授權建立人員] like '%" + KeyB[i] + "%'"
                   + " or [授權修改人員] like '%" + KeyB[i] + "%'"
                   + " or [授權備註說明] like '%" + KeyB[i] + "%'"
                   + " or [申請編號] like '%" + KeyB[i] + "%'"
                   + " or [申請事項] like '%" + KeyB[i] + "%'"
                   + ")";
            }
        }
        if (ChkApKey.Checked & TextKey.Text != "")   //作業主機
        {
            string[] KeyC = TextKey.Text.Trim().Split(',');

            for (int i = 0; i < KeyC.GetLength(0); i++)
            {
                strKey = strKey + " or ([使用主機] like '%" + KeyC[i] + "%'"
                   + " or [系統全名] like '%" + KeyC[i] + "%'"
                   + " or [主要功能] like '%" + KeyC[i] + "%'"
                   + " or [作業平台] like '%" + KeyC[i] + "%'"
                   + " or [核心版本] like '%" + KeyC[i] + "%'"
                   + " or [使用IP] like '%" + KeyC[i] + "%'"
                   + " or [監控IP] like '%" + KeyC[i] + "%'"
                   + " or [緊急程度] like '%" + KeyC[i] + "%'"
                   + " or [作業狀態] like '%" + KeyC[i] + "%'"
                   + " or [作業維護人員] like '%" + KeyC[i] + "%'"
                   + " or [作業建立人員] like '%" + KeyC[i] + "%'"
                   + " or [作業修改人員] like '%" + KeyC[i] + "%'"
                   + " or [作業備註說明] like '%" + KeyC[i] + "%'"
                   + ")";
            }
        }
        if (ChkDevKey.Checked & TextKey.Text != "")  //實體設備
        {
            string[] KeyD = TextKey.Text.Trim().Split(',');

            for (int i = 0; i < KeyD.GetLength(0); i++)
            {
                strKey = strKey + " or ([放置地點] like '%" + KeyD[i] + "%'"
                   + " or [設備型態] like '%" + KeyD[i] + "%'"
                   + " or [設備種類] like '%" + KeyD[i] + "%'"
                   + " or [設備名稱] like '%" + KeyD[i] + "%'"
                   + " or [設備用途] like '%" + KeyD[i] + "%'"
                   + " or [使用財編A] like '%" + KeyD[i] + "%'"
                   + " or [使用財編B] like '%" + KeyD[i] + "%'"
                   + " or [財產名稱] like '%" + KeyD[i] + "%'"
                   + " or [財產別名] like '%" + KeyD[i] + "%'"
                   + " or [使用廠牌] like '%" + KeyD[i] + "%'"
                   + " or [使用型式] like '%" + KeyD[i] + "%'"
                   + " or [數量單位] like '%" + KeyD[i] + "%'"
                   + " or [保管人員] like '%" + KeyD[i] + "%'"
                   + " or [財產保管人員] like '%" + KeyD[i] + "%'"
                   + " or [取得來源] like '%" + KeyD[i] + "%'"
                   + " or [規格] like '%" + KeyD[i] + "%'"
                   + " or [硬體序號] like '%" + KeyD[i] + "%'"
                   + " or [iLoIP] like '%" + KeyD[i] + "%'"
                   + " or [維護廠商] like '%" + KeyD[i] + "%'"
                   + " or [設備維護人員] like '%" + KeyD[i] + "%'"
                   + " or [設備狀態] like '%" + KeyD[i] + "%'"
                   + " or [關機順序] like '%" + KeyD[i] + "%'"
                   + " or [設備建立人員] like '%" + KeyD[i] + "%'"
                   + " or [設備修改人員] like '%" + KeyD[i] + "%'"
                   + " or [設備備註說明] like '%" + KeyD[i] + "%'"
                   + ")";
            }
        }
        if (strKey != "") SQL = SQL + " and (" + strKey.Substring(3) + ")";

        //Response.Write(SQL);
        //Response.End();

        int pos = SQL.IndexOf("FROM");
        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
        Conn.Open();
        SqlCommand cmd = new SqlCommand("SELECT sum([軟體單價]) " + SQL.Substring(pos) + " and [授權狀態]<>'已結案'", Conn);
        SqlDataReader dr = null;

        dr = cmd.ExecuteReader();
        try { if (dr.Read()) TextTotal.Text = String.Format("{0:C0}", dr[0]); }
        catch { }
        dr.Close();

        cmd.CommandText = "SELECT count(*) " + SQL.Substring(pos);
        dr = cmd.ExecuteReader();
        TextCount.Text = "0";
        try { if (dr.Read()) TextCount.Text = dr[0].ToString(); }
        catch { }

        cmd.Cancel(); cmd.Dispose(); dr.Close(); Conn.Close(); Conn.Dispose();

        string ViewCols = SelShow.SelectedValue; //所有欲顯示欄位名稱
        SQL = "Select " + ViewCols + SQL.Substring(pos);
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
        if (GridView1.DataKeys[e.NewSelectedIndex].Value.ToString() != "")
            OpenExecWindow("window.open('AskEdit.aspx?AskNo=" + GridView1.DataKeys[e.NewSelectedIndex].Value.ToString() + "','_blank');");
        else
            OpenExecWindow("window.open('SwEdit.aspx?SwNo=" + GridView1.Rows[e.NewSelectedIndex].Cells[2].Text + "','_blank');");
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

    protected void SelSQL_SelectedIndexChanged(object sender, EventArgs e)
    {
        TextSQL.Text = SelSQL.SelectedValue;
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

    protected Boolean RightCheck() //是否有權修改資料
    {
        if (InGroup(Session["UserName"].ToString(), "軟體小組") | InGroup(Session["UserName"].ToString(), "SSM小組")) return (true);
        else return (false);
    }

    protected Boolean InGroup(string ChkName, string ChkUnit) //檢查ChkName是否為ChkUnit成員或本身
    {
        Boolean TF = false;

        if (ChkName == ChkUnit) TF = true;  //是否同名
        else if (ChkUnit == "") TF = false;	//檢查單位必填
        else //是否為成員UN (課別與小組同義)
        {
            SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
            Conn.Open();
            SqlCommand cmd = new SqlCommand("select * from Config where (Kind='" + ChkUnit + "') and Item='" + ChkName + "'", Conn);
            SqlDataReader dr = null;
            dr = cmd.ExecuteReader();
            if (dr.Read()) TF = true;
            cmd.Cancel(); cmd.Dispose(); dr.Close(); Conn.Close(); Conn.Dispose();
        }

        return (TF);
    }
}