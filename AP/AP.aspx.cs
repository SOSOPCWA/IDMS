using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.Configuration;
using System.Data;
using System.Data.SqlClient;

public partial class AP_AP : System.Web.UI.Page
{   //-------------起始頁面---------------------------------------------------------------------
    protected void Page_Load(object sender, EventArgs e)
    {
        if (!IsPostBack)   //第一次開網頁時
        {
            TreeView1.Nodes[0].Text = "<font color='blue' size='large'><b>作業主機</b></font>";
            TreeView1.Nodes[0].Value = "作業主機";

            ViewState["SQL"] = "SELECT * FROM [View_作業主機] WHERE [設備編號] not in (select [設備編號] from [View_設備管理] where [設備型態]='網路設備')";
            ViewState["Key"] = ""; ViewState["TreeSQL"] = ViewState["SQL"];

            if (Request["Search"] != null)  //外部查詢
            {
                int pos = Session["ApSQL"].ToString().IndexOf("FROM");
                Session["ApSQL"] = "SELECT * " + Session["ApSQL"].ToString().Substring(pos); //ViewState["SQL"];
            }
            else Session["ApSQL"] = ViewState["SQL"];            
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
        string DefSQL = " from [作業主機]";
        if (IsParent) DefSQL = DefSQL + " where " + GetNodeSQL(SelNode1.SelectedItem, node);
        string SQL = "select distinct [" + SelItem + "]" + DefSQL;

        switch (SelItem)
        {
            //case "維護群組":
            //    {
            //        if (IsParent) SQL = "select distinct [維護人員]" + DefSQL + " and [維護人員] in (select [成員] from [View_組織架構] where [性質]='群組') order by [維護人員]";
            //        else SQL = "select distinct [維護人員]" + DefSQL + " where [維護人員] in (select [成員] from [View_組織架構] where [性質]='群組') order by [維護人員]";
            //        break;
            //    }
            case "維護科別":
                {
                    SQL = "select distinct [課別] from [View_組織架構] where [成員] in (select distinct [維護人員]" + DefSQL + ") order by [課別]";
                    break;
                }
            case "系統全名":
                {
                    SQL = "SELECT [系統全名] FROM [View_系統資源] where [資源種類]='系統' order by [系統全名]";
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
            case "維護人員":
                {
                    SQL = "([維護人員]='" + node.Value + "' or [維護人員] in (select distinct Kind from Config where Item='" + node.Value + "'))";
                    break;
                }
            //case "維護群組":
            //    {
            //        SQL = "[維護人員]='" + node.Value + "'";
            //        break;
            //    }
            case "維護課別":
                {
                    SQL = "[維護人員] in (select distinct [成員] from [View_組織架構] where [課別]='" + node.Value + "')";
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
            Msg.Text = "<script>opener.document.form1.TextApNo.value=" + GridView1.DataKeys[e.NewSelectedIndex].Value.ToString()
                + ";opener.document.getElementById('LblApNo').innerHTML='" + GridView1.Rows[e.NewSelectedIndex].Cells[2].Text
                + "';window.close();</script>";
            Page.Controls.Add(Msg);
        }
        else
        {
            OpenExecWindow("window.open('ApEdit.aspx?ApNo=" + GridView1.DataKeys[e.NewSelectedIndex].Value.ToString() + "','_blank');");
        }
    }

    protected void OpenExecWindow(string strJavascript)    //選取後就另開視窗
    {
        if (ScriptManager.GetCurrent(Page) == null)
        {
            Page.ClientScript.RegisterStartupScript(Page.GetType(), "ApEdit", strJavascript, true);
        }
        else
        {
            ScriptManager.RegisterStartupScript(Page, Page.GetType(), "ApEdit", strJavascript, true);
        }
    }
    protected void BtnNew_Click(object sender, EventArgs e)
    {
        OpenExecWindow("window.open('ApEdit.aspx','_blank');");
    }

    protected void BtnSearch_Click(object sender, EventArgs e)
    {
        ViewState["Key"] = "Key"; ExecSearch();   //執行查詢
    }

    protected void BtnPKNo_Click(object sender, EventArgs e)  //以主鍵編號查詢
    {
        Literal Msg = new Literal();
        int n;
        if (int.TryParse(TextKey.Text.ToString(), out n))
        {
            string ApNo = TextKey.Text; if (ApNo == null) ApNo = "0";
            SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
            Conn.Open();
            SqlCommand cmd = new SqlCommand("select * from [作業主機] where [作業編號]=" + ApNo, Conn);
            SqlDataReader dr = null;
            dr = cmd.ExecuteReader();
            if (dr.Read()) OpenExecWindow("window.open('ApEdit.aspx?ApNo=" + ApNo + "','_blank');");
            else
            {
                Msg.Text = "<script>alert('查無作業編號為" + ApNo + "的系統作業！');</script>";
                Page.Controls.Add(Msg);
            }
            cmd.Cancel(); cmd.Dispose(); dr.Close(); Conn.Close(); Conn.Dispose();
        }
        else
        {
            Msg.Text = "<script>alert('請檢查輸入的作業編號是否為數字！');</script>";
            Page.Controls.Add(Msg);
        }
    }

    protected string GetKeySQL()
    {
        string[] KeyA = TextKey.Text.Trim().Split(',');
        string SQL="";

        for (int i = 0; i < KeyA.GetLength(0); i++)
        {
            if (i > 0) SQL = SQL + " and ";
            SQL = SQL + "([主機名稱] like '%" + KeyA[i] + "%'"
                + " or [系統全名] like '%" + KeyA[i] + "%'"
                + " or [主要功能] like '%" + KeyA[i] + "%'"
                + " or [作業平台] like '%" + KeyA[i] + "%'"
                + " or [核心版本] like '%" + KeyA[i] + "%'"
                + " or [IP] like '%" + KeyA[i] + "%'"
                + " or [監控IP] like '%" + KeyA[i] + "%'"
                + " or [備註說明] like '%" + KeyA[i] + "%'"
                + " or [緊急程度] like '%" + KeyA[i] + "%'"
                + " or [作業狀態] like '%" + KeyA[i] + "%'"
                + " or [維護人員] like '%" + KeyA[i] + "%'"
                + " or [維護人員] in (select [Kind] from [Config] where [Item] like '%" + KeyA[i] + "%')"
                + " or [建立人員] like '%" + KeyA[i] + "%'"
                + " or [修改人員] like '%" + KeyA[i] + "%'"
                + ")";
        }
        return (SQL);
    }

    protected void ExecSearch()   //執行查詢
    {
        //------------------------------------------------------------------------------查詢結果
        string SQL = ""; int pos = 0;
        if (ViewState["Key"].ToString() == "Key")
        {
            pos = ViewState["TreeSQL"].ToString().IndexOf("FROM");
            Session["ApSQL"] = "SELECT * " + ViewState["TreeSQL"].ToString().Substring(pos) + " and " + GetKeySQL().ToString();
        }
        else if (ViewState["Key"].ToString() == "Tree")
        {
            pos = ViewState["TreeSQL"].ToString().IndexOf("FROM");
            Session["ApSQL"] = "SELECT * " + ViewState["TreeSQL"].ToString().Substring(pos);
        }
        else
        {
            pos = Session["ApSQL"].ToString().IndexOf("FROM");
            Session["ApSQL"] = "SELECT * " + Session["ApSQL"].ToString().Substring(pos);
        }

        SQL = Session["ApSQL"].ToString();
        SqlDataSource1.SelectCommand = SQL + " order by [主機名稱]";
        GridView1.AllowPaging = ChkPage.Checked;
        //------------------------------------------------------------------------------統計資訊   
        pos = SQL.IndexOf("FROM");

        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
        Conn.Open();
        SqlCommand cmd = new SqlCommand("SELECT count(*) " + SQL.Substring(pos), Conn);
        SqlDataReader dr = null;

        dr = cmd.ExecuteReader();
        TextCount.Text = "0";
        try { if (dr.Read()) TextCount.Text = dr[0].ToString(); }
        catch { }
        dr.Close();

        cmd.Cancel(); cmd.Dispose(); dr.Close(); Conn.Close(); Conn.Dispose();
    }
}