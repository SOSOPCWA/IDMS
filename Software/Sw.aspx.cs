using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.Configuration;
using System.Data;
using System.Data.SqlClient;

public partial class SoftWare_Sw : System.Web.UI.Page
{   //-------------起始頁面---------------------------------------------------------------------
    protected void Page_Load(object sender, EventArgs e)
    {
        if (!IsPostBack)   //第一次開網頁時
        {
            TreeView1.Nodes[0].Text = "<font color='blue' size='large'><b>軟體保管</b></font>";
            TreeView1.Nodes[0].Value = "軟體保管";

            //樹狀查詢後，SQL要保持住，否則會用預設值
            string ViewCols = GetSearchCols("軟體保管"); //所有欲顯示欄位名稱
            ViewState["SQL"] = "SELECT " + ViewCols + " FROM [軟體主檔] WHERE 1=1";
            ViewState["Key"] = ""; ViewState["TreeSQL"] = ViewState["SQL"];

            if (Request["Search"] != null)  //外部查詢
            {
                int pos = Session["SwSQL"].ToString().IndexOf("FROM");
                Session["SwSQL"] = "SELECT " + ViewCols + " " + Session["SwSQL"].ToString().Substring(pos); //ViewState["SQL"];
            }
            else Session["SwSQL"] = ViewState["SQL"];
        }
        else //非第一次開網頁(網頁回傳時)
        {
            //是否要重建樹狀結構之判斷；兩邊均須ToString()，比對才能相等
            if (ViewState["SelNode1"].ToString() != SelNode1.SelectedValue.ToString() | ViewState["SelNode2"].ToString() != SelNode2.SelectedValue.ToString())
            {
                TreeView1.Nodes[0].ChildNodes.Clear();  //清除樹狀結構 
                TreeNodeAdd(TreeView1.Nodes[0]);   //產生樹狀結構
            }
        }
        ViewState["SelNode1"] = SelNode1.SelectedValue;
        ViewState["SelNode2"] = SelNode2.SelectedValue;
    }

    protected string GetSearchCols(string SearchWhat) //取得顯示欄位清單
    {
        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
        Conn.Open();
        SqlCommand cmd = new SqlCommand("SELECT Memo from Config where Kind='軟體查詢' and Item='" + SearchWhat + "'", Conn);
        SqlDataReader dr = null;
        dr = cmd.ExecuteReader();
        string cfg = "*";
        try { if (dr.Read()) cfg=dr[0].ToString(); }
        catch { cfg="*"; }
        cmd.Cancel(); cmd.Dispose(); dr.Close(); Conn.Close(); Conn.Dispose();

        return (cfg);
    }

    protected void Page_PreRenderComplete(object sender, EventArgs e)   //放Page_Load無法顯示
    {
        ExecSearch();   //執行查詢
    }
    //------------產生樹狀結構------------------------------------------------------------------------------
    protected void TreeView1_TreeNodePopulate(object sender, TreeNodeEventArgs e)  //當使用者按一下開啟節點時
    {
        TreeNodeAdd(e.Node);
    }

    protected string GetTreeSQL(TreeNode node, string SelItem, Boolean IsParent)  //取得產生第一或第二節點的SQL
    {
        string DefSQL = " from [軟體主檔]";
        if (IsParent) DefSQL = DefSQL + " where " + GetNodeSQL(SelNode1.SelectedItem, node);
        string SQL = "select distinct [" + SelItem + "]" + DefSQL;

        switch (SelItem)
        {
            case "表單年度":
                {
                    SQL = "select distinct substring([表單編號],1,4) As [表單年度] from [軟體主檔] order by [表單年度]";
                    break;
                }
            case "適用機型":
                {
                    SQL = "select distinct Rtrim(Ltrim([適用廠牌]))+' '+Rtrim(Ltrim([適用型式])) as [適用機型]" + DefSQL + " order by [適用機型]";
                    break;
                }
            default:
                {
                    SQL = SQL + " order by [" + SelItem + "]";
                    break;
                }
        }
        return (SQL);
    }

    protected string GetNodeSQL(ListItem item, TreeNode node)   //取得選取節點之查詢SQL
    {
        string SQL = "[" + item.Text + "]='" + node.Value + "'";
        switch (item.Text)
        {
            case "表單年度":
                {
                    if (node.Value.Trim() == "") SQL = "[表單編號]=''";
                    else SQL = "[表單編號] like '" + node.Value + "%'";
                    break;
                }
            case "適用機型":
                {
                    SQL = "Rtrim(Ltrim([適用廠牌]))+' '+Rtrim(Ltrim([適用型式]))='" + node.Value + "'";
                    break;
                }
        }
        return (SQL);
    }

    protected void TreeNodeAdd(TreeNode node)   //從某節點開始往下產生樹狀結構
    {
        if (node.ChildNodes.Count == 0)   //子節點展開過了沒？沒有，則需要產生，否則不必重複產生
        {
            switch (node.Depth)   //使用者所按節點深度
            {
                case 0:    //若在根節點，則產生第一子節點
                    if (SelNode2.SelectedValue != "") RootNodeAdd(node, GetTreeSQL(node, SelNode1.SelectedValue, false));
                    else ChildNodeAdd(node, GetTreeSQL(node, SelNode1.SelectedValue, false));
                    break;
                case 1:   //若在第一子節點，則產生第二子節點
                    if (SelNode2.SelectedValue != "")
                    {
                        ChildNodeAdd(node, GetTreeSQL(node, SelNode2.SelectedValue, true));
                    }
                    break;
            }
        }
    }

    protected void RootNodeAdd(TreeNode node, string SQL)    //產生有子節點之目錄節點
    {
        SqlCommand sqlQuery = new SqlCommand(SQL);
        DataSet ds = RunQuery(sqlQuery);

        if (ds.Tables.Count > 0)
        {
            foreach (DataRow row in ds.Tables[0].Rows)
            {
                TreeNode NewNode = new TreeNode(row[0].ToString(), row[0].ToString());  //TreeNode(值，顯示)
                NewNode.PopulateOnDemand = true;    //還有子節點，使用者按一下會觸發TreeNodePopulate事件
                NewNode.SelectAction = TreeNodeSelectAction.Expand;   //觸發TreeNodeExpanded
                NewNode.Expanded = false;   //不展開節點
                node.ChildNodes.Add(NewNode);
            }
        }

        sqlQuery.Cancel(); ds.Dispose();
    }

    protected void ChildNodeAdd(TreeNode node, string SQL)    //產生無子節點之樹葉節點
    {
        SqlCommand sqlQuery = new SqlCommand();
        sqlQuery.CommandText = SQL;
        DataSet ds = RunQuery(sqlQuery);

        if (ds.Tables.Count > 0)
        {
            foreach (DataRow row in ds.Tables[0].Rows)
            {
                TreeNode NewNode = new TreeNode(row[0].ToString(), row[0].ToString());
                NewNode.PopulateOnDemand = false;   //已無子節點，不再展開                   
                node.ChildNodes.Add(NewNode);
            }
        }
        sqlQuery.Cancel(); ds.Dispose();
    }

    protected DataSet RunQuery(SqlCommand sqlQuery) //讀取節點資訊
    {
        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
        SqlDataAdapter dbAdapter = new SqlDataAdapter();
        dbAdapter.SelectCommand = sqlQuery;
        sqlQuery.Connection = Conn;
        DataSet QueryDataSet = new DataSet();
        //Response.Write(sqlQuery.CommandText);
        //Response.End();
        dbAdapter.Fill(QueryDataSet);
        dbAdapter.Dispose(); Conn.Close(); Conn.Dispose();
        return (QueryDataSet);
    }
    //------------點選節點已產生查詢結果
    protected void TreeView1_SelectedNodeChanged(object sender, EventArgs e)
    {
        string SQL = "";

        if (TreeView1.SelectedNode == TreeView1.Nodes[0]) ViewState["TreeSQL"] = ViewState["SQL"];    //排除根目錄
        else
        {
            if (SelNode2.SelectedValue == "" | SelNode1.SelectedValue == SelNode2.SelectedValue)    //單一選項(含重複選項，排除根目錄)
            {
                if (TreeView1.SelectedNode != TreeView1.Nodes[0]) SQL = GetNodeSQL(SelNode1.SelectedItem, TreeView1.SelectedNode);
            }
            else //雙選項
            {
                SQL = GetNodeSQL(SelNode1.SelectedItem, TreeView1.SelectedNode.Parent) + " and " + GetNodeSQL(SelNode2.SelectedItem, TreeView1.SelectedNode);
            }

            ViewState["TreeSQL"] = ViewState["SQL"] + " and " + SQL;  //樹狀查詢後，SQL要保持住，否則會用預設值
        }

        ViewState["Key"] = "Tree"; ExecSearch();   //執行查詢
    }  

    protected void GridView1_SelectedIndexChanging(object sender, GridViewSelectEventArgs e)    //選取後就另開視窗
    {
        if (Request["SelMode"] != null)
        {
            Literal Msg = new Literal();
            Msg.Text = "<script>opener.document.form1.TextSwNo.value=" + GridView1.DataKeys[e.NewSelectedIndex].Value.ToString()
                + ";opener.document.getElementById('LblSwName').innerHTML='" + GridView1.Rows[e.NewSelectedIndex].Cells[3].Text
                + "';window.close();</script>";
            Page.Controls.Add(Msg);
        }
        else
        {
            //if (GridView1.SelectedIndex != e.NewSelectedIndex) OpenExecWindow("window.open('SwEdit.aspx?SwNo=" + GridView1.DataKeys[e.NewSelectedIndex].Value.ToString() + "','_blank');");
            //else { GridView1.SelectedIndex = -1; e.NewSelectedIndex = -1; }
            OpenExecWindow("window.open('SwEdit.aspx?SwNo=" + GridView1.DataKeys[e.NewSelectedIndex].Value.ToString() + "','_blank');");
        }
    }

    protected void OpenExecWindow(string strJavascript)    //選取後就另開視窗
    {
        if (ScriptManager.GetCurrent(Page) == null)
        {
            Page.ClientScript.RegisterStartupScript(Page.GetType(), "SwEdit", strJavascript, true);
        }
        else
        {
            ScriptManager.RegisterStartupScript(Page, Page.GetType(), "SwEdit", strJavascript, true);
        }
    }
    protected void BtnNew_Click(object sender, EventArgs e)
    {
        OpenExecWindow("window.open('SwEdit.aspx','_blank');");
    }

    protected void BtnSearch_Click(object sender, EventArgs e)
    {
        ViewState["Key"] = "Key"; ExecSearch();   //執行查詢
    }

    protected string GetKeySQL()
    {
        string[] KeyA = TextKey.Text.Trim().Split(',');
        string SQL = "";

        for (int i = 0; i < KeyA.GetLength(0); i++)
        {
            if (i > 0) SQL = SQL + " and ";

            switch (SelKey.SelectedValue)
            {
                case "":
                    SQL = SQL + "([表單編號] like '%" + KeyA[i] + "%'"
                        + " or [軟體單位] like '%" + KeyA[i] + "%'"
                        + " or [軟體人員] like '%" + KeyA[i] + "%'"
                        + " or [軟體名稱] like '%" + KeyA[i] + "%'"
                        + " or [購買版本] like '%" + KeyA[i] + "%'"
                        + " or [可用版本] like '%" + KeyA[i] + "%'"
                        + " or [授權方式] like '%" + KeyA[i] + "%'"
                        + " or [授權說明] like '%" + KeyA[i] + "%'"
                        + " or [軟體來源] like '%" + KeyA[i] + "%'"
                        + " or [軟體功能] like '%" + KeyA[i] + "%'"
                        + " or [適用廠牌] like '%" + KeyA[i] + "%'"
                        + " or [適用型式] like '%" + KeyA[i] + "%'"
                        + " or [購買序號] like '%" + KeyA[i] + "%'"
                        + " or [降級序號] like '%" + KeyA[i] + "%'"
                        + " or [期限說明] like '%" + KeyA[i] + "%'"
                        + " or [價格說明] like '%" + KeyA[i] + "%'"
                        + " or [提供單位] like '%" + KeyA[i] + "%'"
                        + " or [存放媒體] like '%" + KeyA[i] + "%'"
                        + " or [圖書編號] like '%" + KeyA[i] + "%'"
                        + " or [軟體附件] like '%" + KeyA[i] + "%'"
                        + " or [軟體狀態] like '%" + KeyA[i] + "%'"
                        + " or [減損原因] like '%" + KeyA[i] + "%'"
                        + " or [減損方式] like '%" + KeyA[i] + "%'"
                        + " or [建立人員] like '%" + KeyA[i] + "%'"
                        + " or [修改人員] like '%" + KeyA[i] + "%'"
                        + " or [備註說明] like '%" + KeyA[i] + "%'"
                        + ")";
                    break;
                case "軟體編號":
                    SQL = SQL + "([" + SelKey.SelectedValue + "] =" + KeyA[i] + ")";
                    break;
                default:
                    SQL = SQL + "([" + SelKey.SelectedValue + "] like '%" + KeyA[i] + "%')";
                    break;
            }
        }
        return (SQL);
    }

    protected string GetExecSQL()
    {
        switch (ViewState["Key"].ToString())
        {
            case "Search":
                {
                    int pos = Session["SwSQL"].ToString().IndexOf("FROM");
                    Session["SwSQL"] = "SELECT " + SelShow.SelectedValue + " " + Session["SwSQL"].ToString().Substring(pos) + " and " + GetKeySQL().ToString();
                    break;
                }
            case "All":
                {
                    int pos = ViewState["SQL"].ToString().IndexOf("FROM");
                    Session["SwSQL"] = "SELECT " + SelShow.SelectedValue + " " + ViewState["SQL"].ToString().Substring(pos) + " and " + GetKeySQL().ToString();
                    break;
                }
        }

        return (Session["SwSQL"].ToString());
    }

    protected void ExecSearch()   //執行查詢
    {
        //------------------------------------------------------------------------------查詢結果
        string SQL = ""; int pos = 0;
        if (ViewState["Key"].ToString() == "Key")
        {
            pos = ViewState["TreeSQL"].ToString().IndexOf("FROM");
            Session["SwSQL"] = "SELECT " + SelShow.SelectedValue + " " + ViewState["TreeSQL"].ToString().Substring(pos) + " and " + GetKeySQL().ToString();
        }
        else if (ViewState["Key"].ToString() == "Tree")
        {
            pos = ViewState["TreeSQL"].ToString().IndexOf("FROM");
            Session["SwSQL"] = "SELECT " + SelShow.SelectedValue + " " + ViewState["TreeSQL"].ToString().Substring(pos);
        }
        else
        {
            pos = Session["SwSQL"].ToString().IndexOf("FROM");
            Session["SwSQL"] = "SELECT " + SelShow.SelectedValue + " " + Session["SwSQL"].ToString().Substring(pos);
        }

        SQL = Session["SwSQL"].ToString();
        SqlDataSource1.SelectCommand = SQL + " order by [表單編號]";
        if (ChkPage.Checked) GridView1.AllowPaging = true;
        else GridView1.AllowPaging = false;
        //------------------------------------------------------------------------------統計資訊   
        pos = SQL.IndexOf("FROM");

        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
        Conn.Open();
        SqlCommand cmd = new SqlCommand("SELECT sum([軟體總價]) " + SQL.Substring(pos), Conn);
        SqlDataReader dr = null;
        dr = cmd.ExecuteReader();
        TextTotal.Text = "0";
        try { if (dr.Read()) TextTotal.Text = ((decimal)dr[0]).ToString("c0"); }
        catch { }
        dr.Close();

        cmd.CommandText = "SELECT count(*) " + SQL.Substring(pos);
        dr = cmd.ExecuteReader();
        TextCount.Text = "0";
        try { if (dr.Read()) TextCount.Text = dr[0].ToString(); }
        catch { }

        cmd.Cancel(); cmd.Dispose(); dr.Close(); Conn.Close(); Conn.Dispose();
    }

    protected void BtnExcel_Click(object sender, EventArgs e)
    {
        if (GridView1.Rows.Count > 0)
        {
            int pos = Session["SwSQL"].ToString().IndexOf("FROM");
            string ViewCols = SelShow.SelectedValue; //所有欲顯示欄位名稱

            GridView gvExport = new GridView();
            gvExport.DataSource = SqlDataSource1;
            SqlDataSource1.SelectCommand = "Select " + ViewCols + Session["SwSQL"].ToString().Substring(pos);
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