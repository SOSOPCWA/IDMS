using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.Configuration;
using System.Data;
using System.Data.SqlClient;

public partial class Property_Property : System.Web.UI.Page
{   //-------------起始頁面---------------------------------------------------------------------
    protected void Page_Load(object sender, EventArgs e)
    {
        if (!IsPostBack)   //第一次開網頁時
        {
            TreeView1.Nodes[0].Text = "<font color='blue' size='large'><b>秘總財產</b></font>";
            TreeView1.Nodes[0].Value = "秘總財產";

            //樹狀查詢後，SQL要保持住，否則會用預設值
            ViewState["SQL"] = "SELECT * FROM [財產主檔] WHERE 1=1";
            ViewState["Key"] = ""; ViewState["TreeSQL"] = ViewState["SQL"];
            if (Request["PropA"] == null) Session["PropSQL"] = ViewState["SQL"] + " and [無效註記]<>'N'";
            else Session["PropSQL"] = "SELECT * FROM [財產主檔] WHERE [財產編號A]='" + Request["PropA"] + "' and [財產編號B]='" + Request["PropB"] + "'";
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
        string DefSQL = " from [財產主檔]";
        if (IsParent) DefSQL = DefSQL + " where " + GetNodeSQL(SelNode1.SelectedItem, node);
        string SQL = "select distinct [" + SelItem + "]" + DefSQL;

        switch (SelItem)
        {
            case "廠牌型式":
                {
                    SQL = "select distinct [廠牌]+' '+[型式] as [廠牌型式]" + DefSQL + " order by [廠牌型式]";
                    break;
                }
            case "保管課別":
                {
                    SQL = "select distinct [課別] from [View_組織架構] where [成員] in (select distinct [保管人員]" + DefSQL + ") order by [課別]";
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
            case "財產大類":
            case "使用年限":
                {
                    SQL = "[" + item.Value + "]=" + node.Value; 
                    break;
                }
            case "廠牌型式":
                {
                    SQL = "[廠牌]+' '+[型式]='" + node.Value + "'";
                    break;
                }
            case "保管課別":
                {
                    SQL = "[保管人員] in (select distinct [成員] from [View_組織架構] where [課別]='" + node.Value + "')";
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
        string PropA = GridView1.Rows[e.NewSelectedIndex].Cells[1].Text, PropB = GridView1.Rows[e.NewSelectedIndex].Cells[2].Text;
        string DevNo = "";
        
        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
        Conn.Open();
        SqlCommand cmd = new SqlCommand("SELECT [設備編號] AS CountX FROM [實體設備] WHERE [財產編號A]='" + PropA + "' and [財產編號B]='" + PropB + "'", Conn);
        SqlDataReader dr = null;
        
        dr = cmd.ExecuteReader();

        if (dr.Read())  //財產編號可能重複，不可直接帶出編輯畫面
        {
            DevNo = dr[0].ToString();
            if (dr.Read()) OpenExecWindow("window.open('../Device/Device.aspx?PropA=" + PropA + "&PropB=" + PropB + "','_self');");
            else OpenExecWindow("window.open('../Device/DevEdit.aspx?DevNo=" + DevNo + "','_blank');");
        }
        else //查無此財產編號
        {
            //Response.Write(dr.Read().ToString());
            
            Literal Msg = new Literal();
            Msg.Text = "<script>if(confirm('[實體設備]查無此財產編號(" + PropA + " " + PropB + ")之資料！是否要帶出設備編輯介面以新增資料？')) window.open('../Device/DevEdit.aspx?PropA=" + PropA + "&PropB=" + PropB + "','_blank');</script>";
            Page.Controls.Add(Msg);
        }

        cmd.Cancel(); cmd.Dispose(); dr.Close(); Conn.Close(); Conn.Dispose();
    }

    protected void OpenExecWindow(string strJavascript)    //選取後就另開視窗
    {
        if (ScriptManager.GetCurrent(Page) == null)
        {
            Page.ClientScript.RegisterStartupScript(Page.GetType(), "PropEdit", strJavascript, true);
        }
        else
        {
            ScriptManager.RegisterStartupScript(Page, Page.GetType(), "PropEdit", strJavascript, true);
        }
    }

    protected void BtnSearch_Click(object sender, EventArgs e)
    {
        ViewState["Key"] = "Key"; ExecSearch();   //執行查詢
    }

    protected void BtnBuy_Click(object sender, EventArgs e)  //以購買日期查詢
    {
        ViewState["Key"] = "Buy"; ExecSearch();   //執行查詢
    }

    protected string GetKeySQL()
    {
        string[] KeyA = TextKey.Text.Trim().Split(' ');
        string SQL="";

        for (int i = 0; i < KeyA.GetLength(0); i++)
        {
            if (i > 0) SQL = SQL + " and ";
            SQL = SQL + "([財產編號A] like '%" + KeyA[i] + "%'"
                + " or [財產編號B] like '%" + KeyA[i] + "%'"
                + " or [財產別名] like '%" + KeyA[i] + "%'"
                + " or [廠牌] like '%" + KeyA[i] + "%'"
                + " or [型式] like '%" + KeyA[i] + "%'"
                + " or [計量單位] like '%" + KeyA[i] + "%'"
                + " or [保管人員] like '%" + KeyA[i] + "%'"
                + " or [財產附屬] like '%" + KeyA[i] + "%'"
                + ")";
        }
        return (SQL);
    }    

    protected void BtnExcel_Click(object sender, EventArgs e)
    {
        if (GridView1.Rows.Count > 0)
        {
            int pos = Session["PropSQL"].ToString().IndexOf("FROM");
            string ViewCols = "財產編號A,財產編號B,財產大類,財產別名,廠牌,型式,數量,計量單位,價值,CONVERT(char(10),[購買日期],111) as [購買日期],使用年限,保管人員,無效註記,財產附屬 "; //所有欲顯示欄位名稱

            //Response.Write("Select " + ViewCols + Session["PropSQL"].ToString().Substring(pos));
            //Response.End();
            
            GridView gvExport = new GridView();
            gvExport.DataSource = SqlDataSource1;
            SqlDataSource1.SelectCommand = "Select " + ViewCols + Session["PropSQL"].ToString().Substring(pos) + " order by [購買日期] desc";
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

    protected void ExecSearch()   //執行查詢
    {
        //------------------------------------------------------------------------------查詢結果
        string SQL = ""; int pos = 0;
        if (ViewState["Key"].ToString() == "Key" | ViewState["Key"].ToString() == "Buy")
        {
            if (ViewState["Key"].ToString() == "Buy") Session["PropSQL"] = ViewState["TreeSQL"].ToString() + " and [購買日期] like '" + TextKey.Text.Replace("/", "-") + "%'";
            else Session["PropSQL"] = ViewState["TreeSQL"].ToString() + " and " + GetKeySQL().ToString();

            if (SelBuy.SelectedValue != "" & Request["PropA"] == null) Session["PropSQL"] = Session["PropSQL"].ToString() + " and [購買日期]>='" + DateTime.Now.AddDays(-int.Parse(SelBuy.SelectedValue)).ToString("yyyy/MM/dd") + "'";

            if (SelIDMS.SelectedValue == "已登") Session["PropSQL"] = Session["PropSQL"].ToString() + " and [財產編號A]+' '+[財產編號B] in (select [財產編號A]+' '+[財產編號B] from [實體設備])";
            else if (SelIDMS.SelectedValue == "未登") Session["PropSQL"] = Session["PropSQL"].ToString() + " and [財產編號A]+' '+[財產編號B] not in (select [財產編號A]+' '+[財產編號B] from [實體設備])";

            if (!ChkOffice.Checked) Session["PropSQL"] = Session["PropSQL"].ToString() + " and [財產大類]<>6";            
            if (!ChkMark.Checked) Session["PropSQL"] = Session["PropSQL"].ToString() + " and [無效註記]<>'N'";
        }
        else if (ViewState["Key"].ToString() == "Tree") Session["PropSQL"] = ViewState["TreeSQL"];
        else Session["PropSQL"] = Session["PropSQL"];

        SQL = Session["PropSQL"].ToString();
        SqlDataSource1.SelectCommand = SQL + " order by [購買日期] desc";
        GridView1.AllowPaging = ChkPage.Checked;
        //------------------------------------------------------------------------------統計資訊     
        pos = SQL.IndexOf("FROM");

        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
        Conn.Open();
        SqlCommand cmd = new SqlCommand("SELECT sum([價值]) " + SQL.Substring(pos), Conn);
        SqlDataReader dr = null;

        dr = cmd.ExecuteReader();
        TextTotal.Text = "0";
        try { if (dr.Read()) TextTotal.Text = String.Format("{0:C0}", dr[0]); }
        catch { }
        dr.Close();

        cmd.CommandText = "SELECT count(*) " + SQL.Substring(pos);
        dr = cmd.ExecuteReader();
        TextCount.Text = "0";
        try { if (dr.Read()) TextCount.Text = dr[0].ToString(); }
        catch { }
        cmd.Cancel(); cmd.Dispose(); dr.Close(); Conn.Close(); Conn.Dispose();
    }
}