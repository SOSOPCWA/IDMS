using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.Configuration;
using System.Data;
using System.Data.SqlClient;

public partial class Device_Device : System.Web.UI.Page
{   //-------------起始頁面---------------------------------------------------------------------
    protected void Page_Load(object sender, EventArgs e)
    {
        if (!IsPostBack)   //第一次開網頁時
        {
            ViewState["DevType"] = Request["DevType"];
            if (Request["DevType"] == null) ViewState["DevType"] = "系統設備";

            TreeView1.Nodes[0].Text = "<font color='blue' size='large'><b>" + ViewState["DevType"].ToString() + "</b></font>";
            TreeView1.Nodes[0].Value = ViewState["DevType"].ToString();

            string ViewCols = GetSearchCols("通用設備"); //所有欲顯示欄位名稱
            //注意大寫FROM不可改小寫，map.aspx.cs有IndexOf
            ViewState["SQL"] = "SELECT " + ViewCols + " FROM [View_設備管理] WHERE [設備型態]='" + ViewState["DevType"].ToString() + "'";
            ViewState["Key"]="";  ViewState["TreeSQL"] = ViewState["SQL"];
            if (Request["PropA"] != null) Session["DevSQL"] = "SELECT " + ViewCols + " FROM [View_設備管理] WHERE [財產編號A]='" + Request["PropA"].ToString() + "' and [財產編號B]='" + Request["PropB"].ToString() + "'";
            else
            {
                if (Request["Search"] != null) //Tree及Map外部查詢
                {
                    if (Request["Search"] == "Map") //Map外部查詢
                    {
                        Session["DevSQL"] = "SELECT " + ViewCols + " FROM [View_設備管理] where [坐標X]=" + Request["X"] + " and [坐標Y]=" + Request["Y"];
                    }
                    else //PowerTree外部查詢
                    {
                        int pos = Session["DevSQL"].ToString().IndexOf("FROM");
                        Session["DevSQL"] = "SELECT " + ViewCols + " " + Session["DevSQL"].ToString().Substring(pos); //ViewState["SQL"];
                    }
                }
                else Session["DevSQL"] = ViewState["SQL"];
            }            
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

    protected string GetSearchCols(string SearchWhat) //取得顯示欄位清單
    {
        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
        Conn.Open();
        SqlCommand cmd = new SqlCommand("SELECT Memo from Config where Kind='設備查詢' and Item='" + SearchWhat + "'", Conn);
        SqlDataReader dr = null;
        dr = cmd.ExecuteReader();
        string cfg = "*"; 
        try { if (dr.Read()) cfg=dr[0].ToString(); }
        catch { cfg="*"; }
        cmd.Cancel(); cmd.Dispose(); dr.Close(); Conn.Close(); Conn.Dispose();

        return (cfg);
    }
    //------------產生樹狀結構------------------------------------------------------------------------------
    protected void TreeView1_TreeNodePopulate(object sender, TreeNodeEventArgs e)  //當使用者按一下開啟節點時
    {
        TreeNodeAdd(e.Node);
    }

    protected string GetTreeSQL(TreeNode node, string SelItem,Boolean IsParent)   //取得產生第一或第二節點的SQL
    {
        string DefSQL = " from [View_設備管理] where [設備型態]='" + ViewState["DevType"].ToString() + "'";
        if (IsParent) DefSQL = DefSQL + " and " + GetNodeSQL(SelNode1.SelectedItem, node);
        string SQL = "select distinct [" + SelItem + "]" + DefSQL;

        switch (SelItem)
        {
            case "廠牌":
                {
                    SQL = "select distinct [廠牌]" + DefSQL + " and [設備種類] not like '測站%' order by [廠牌]";
                    break;
                }
            case "廠牌型式":
                {
                    SQL = "select distinct [廠牌]+' '+[型式] as [廠牌型式]" + DefSQL + " and [設備種類] not like '測站%' order by [廠牌型式]";
                    break;
                }
            case "維護課別":
                {
                    SQL = "select distinct [課別] from [View_組織架構] where [成員] in (select distinct [維護人員]" + DefSQL + ") order by [課別]";
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
            case "定位名稱":
                {
                    SQL = "([定位名稱]='" + node.Value + "' or [定位名稱] in (" + GetPosSQL("[定位名稱]","[定位名稱]",node.Value) + "))";
                    break;
                }
            case "放置地點":
                {
                    SQL = "([放置地點]='" + node.Value + "' or [放置地點] in (" + GetPosSQL("[區域名稱]+[定位名稱]","[放置地點]",node.Value) + "))";
                    break;
                }
            case "用電電壓":
                {
                    SQL = "[用電電壓]=" + node.Value;
                    break;
                }
            case "廠牌型式":
                {
                    SQL = "[廠牌]+' '+[型式]='" + node.Value + "'";
                    break;
                }
            case "維護人員":
                {
                    SQL = "([維護人員]='" + node.Value + "' or [維護人員] in (select distinct Kind from Config where Item='" + node.Value + "'))";
                    break;
                }
            case "維護課別":
                {
                    SQL = "[維護人員] in (select distinct [成員] from [View_組織架構] where [課別]='" + node.Value + "')";
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

    protected string GetPosSQL(string ColName, string AsName, string PosName)
    {
        string SQL = "select " + ColName + " AS " + AsName + " from [定位設定] where [定位方式]='坐標'";

        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
        Conn.Open();
        SqlCommand cmd = new SqlCommand("select * from [定位設定] where " + ColName + "='" + PosName + "'", Conn);
        SqlDataReader dr = null;
        dr = cmd.ExecuteReader();
        if (dr.Read()) SQL = SQL + " and [坐標X]>=" + dr["坐標X1"].ToString() + " and [坐標X]<=" + dr["坐標X2"].ToString()
            + " and [坐標Y]>=" + dr["坐標Y1"].ToString() + " and [坐標Y]<=" + dr["坐標Y2"].ToString();

        cmd.Cancel(); cmd.Dispose(); dr.Close(); Conn.Close(); Conn.Dispose();
        return (SQL);
    }

    protected void TreeNodeAdd(TreeNode node)   //從某節點開始往下產生樹狀結構
    {
        if (node.ChildNodes.Count == 0)   //子節點展開過了沒？沒有，則需要產生，否則不必重複產生
        {
            switch (node.Depth)   //使用者所按節點深度
            {
                case 0:    //若在根節點，則產生第一子節點
                    if (SelNode2.SelectedValue != "") RootNodeAdd(node, GetTreeSQL(node, SelNode1.SelectedValue,false));
                    else ChildNodeAdd(node, GetTreeSQL(node, SelNode1.SelectedValue,false));
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

    protected void GridView1_SelectedIndexChanging(object sender, GridViewSelectEventArgs e)    //選取後就另開編輯視窗或帶入設備編號
    {
        if (Request["SelMode"] != null)
        {
            Literal Msg = new Literal();
            Msg.Text = "<script>opener.document.form1.TextDevNo.value=" + GridView1.DataKeys[e.NewSelectedIndex].Value.ToString()
                + ";opener.document.getElementById('LblDevName').innerHTML='" + GridView1.Rows[e.NewSelectedIndex].Cells[2].Text
                + "';opener.document.form1.SelHostType.selectedIndex=0;window.close();</script>";
            Page.Controls.Add(Msg);
        }
        else
        {
            OpenExecWindow("window.open('DevEdit.aspx?DevNo=" + GridView1.DataKeys[e.NewSelectedIndex].Value.ToString() + "','_blank');");
        }
    }

    protected void OpenExecWindow(string strJavascript)    //選取後就另開視窗
    {
        if (ScriptManager.GetCurrent(Page) == null)
        {
            Page.ClientScript.RegisterStartupScript(Page.GetType(), "DevEdit", strJavascript, true);
        }
        else
        {
            ScriptManager.RegisterStartupScript(Page, Page.GetType(), "DevEdit", strJavascript, true);
        }
    }

    protected void BtnNew_Click(object sender, EventArgs e) //新增設備，順便帶入設備種類與資產編號
    {
        int pos = 0,pos1=0,pos2=0;
        string DevType = ViewState["DevType"].ToString();
        
        string DevKind = ""; 
        pos = Session["DevSQL"].ToString().IndexOf("[設備種類]="); 
        pos1 = pos + 8; pos2 = Session["DevSQL"].ToString().IndexOf("'", pos1);
        if (pos > 0) DevKind = Session["DevSQL"].ToString().Substring(pos1, pos2 - pos1);         

        OpenExecWindow("window.open('DevEdit.aspx?DevType=" + DevType + "&DevKind=" + DevKind + "','_blank');");
    }

    protected void BtnMap_Click(object sender, EventArgs e) //設備分佈圖
    {
        //使用Session["DevSQL"]傳值
        OpenExecWindow("window.open('../Lib/map.aspx','_blank','location=no;menubar=no;resizable=yes;scrollbars=no;status=no;toolbar=no;fullscreen=yes');");
    }

    protected void BtnSearch_Click(object sender, EventArgs e)  //關鍵字查詢
    {
        ViewState["Key"] = "Key"; ExecSearch();   //執行查詢
    }

    protected void BtnPKNo_Click(object sender, EventArgs e)  //以主鍵編號查詢
    {
        Literal Msg = new Literal();
        int n;
        if (int.TryParse(TextKey.Text.ToString(), out n))
        {
            string DevNo = TextKey.Text; if (DevNo == null) DevNo = "0";
            SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
            Conn.Open();
            SqlCommand cmd = new SqlCommand("select * from [實體設備] where [設備編號]=" + DevNo, Conn);
            SqlDataReader dr = null;
            dr = cmd.ExecuteReader();
            if (dr.Read()) OpenExecWindow("window.open('DevEdit.aspx?DevNo=" + DevNo + "','_blank');");
            else
            {
                Msg.Text = "<script>alert('查無設備編號為" + DevNo + "的設備！');</script>";
                Page.Controls.Add(Msg);
            }
            cmd.Cancel(); cmd.Dispose(); dr.Close(); Conn.Close(); Conn.Dispose();
        }
        else
        {
            Msg.Text = "<script>alert('請檢查輸入的設備編號是否為數字！');</script>";
            Page.Controls.Add(Msg);
        }
    }

    protected string GetKeySQL()
    {
        string[] KeyA = TextKey.Text.Trim().Split(',');
        string SQL = "";

        for (int i = 0; i < KeyA.GetLength(0); i++)
        {
            if (i > 0) SQL = SQL + " and ";

            SQL = SQL + "([資產編號] like '%" + KeyA[i] + "%'"
            + " or [設備名稱] like '%" + KeyA[i] + "%'"
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
            //+ " or [財產保管人員] like '%" + KeyA[i] + "%'"
            + " or [維護廠商] like '%" + KeyA[i] + "%'"
            + " or [維護廠商] in (select [Item] from [Config] where [Kind]='維護廠商' and [Config] like '%" + KeyA[i] + "%')"
            + " or [維護人員] like '%" + KeyA[i] + "%'"
            + " or [維護人員] in (select [Kind] from [Config] where [Item] like '%" + KeyA[i] + "%')"
            + " or [設備狀態] like '%" + KeyA[i] + "%'"
            + " or [關機順序] like '%" + KeyA[i] + "%'"
            + " or [建立人員] like '%" + KeyA[i] + "%'"
            + " or [修改人員] like '%" + KeyA[i] + "%'"
            + " or [設備型態]='網路設備' and [設備編號] in (select [設備編號] from [作業主機] where "
            + " [主機名稱] like '%" + KeyA[i]
            + "%' or [IP] like '%" + KeyA[i]
            + "%' or [緊急程度] like '%" + KeyA[i] + "%')"
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
            Session["DevSQL"] = "SELECT " + SelShow.SelectedValue + " " + ViewState["TreeSQL"].ToString().Substring(pos) + " and " + GetKeySQL().ToString();
        }
        else if (ViewState["Key"].ToString() == "Tree")
        {
            pos = ViewState["TreeSQL"].ToString().IndexOf("FROM");
            Session["DevSQL"] = "SELECT " + SelShow.SelectedValue + " " + ViewState["TreeSQL"].ToString().Substring(pos);
        }
        else
        {
            pos = Session["DevSQL"].ToString().IndexOf("FROM");
            Session["DevSQL"] = "SELECT " + SelShow.SelectedValue + " " + Session["DevSQL"].ToString().Substring(pos);
        }

        SQL = Session["DevSQL"].ToString();
        SqlDataSource1.SelectCommand = Session["DevSQL"].ToString() +" order by [設備名稱]";
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

        cmd.CommandText = "SELECT sum([額定電流]) " + SQL.Substring(pos) + " and ([設備種類]<>'電源' and [設備種類]<>'PDC' and [設備種類]<>'配電盤' and [設備種類]<>'迴路')";
        dr = cmd.ExecuteReader();
        TextSumCurrent.Text = "0";
        try { if (dr.Read()) TextSumCurrent.Text = String.Format("{0:F}", dr[0]); }
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