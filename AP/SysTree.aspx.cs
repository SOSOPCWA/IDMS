using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.Configuration;
using System.Data;
using System.Data.SqlClient;

public partial class AP_SysTree : System.Web.UI.Page
{   //-------------起始頁面---------------------------------------------------------------------
    protected void Page_Load(object sender, EventArgs e)
    {
        if (!Page.IsPostBack)   //網頁非回傳(第一次)執行，用於初始化
        {
            if (Request["SysNo"] == "" | Request["SysNo"] == null)
            {
                TextSysNo.Text = "0";
                ReadNode(TextSysNo.Text, TreeViewDn.Nodes[0]);
            }
            else
            {
                TextSysNo.Text = Request["SysNo"];
                ReadNode(TextSysNo.Text,null);
            }

            ListSysFun();   //列出系統功能待選清單
            helpSysTree.Text = GetValue("IDMS", "select [Memo] from [Config] where [Kind]='系統資源' and [Item]='系統迴路'").Replace("\r\n", "<br />"); //讀取欄位說明
            LinkHostHelp_Click(null, null);
        }
    }

    //僅此事件在DataSource Binding之後，而Page_Load或其它事件，均在那之前
    protected void Page_PreRenderComplete(object sender, EventArgs e)   //放在此事件，改變選項才能自動設值，放Select_Change無效
    {
        if (!IsPostBack & TextSysNo.Text != "" & Request["SysNo"] != "" & Request["SysNo"] != null)
        {
            TextSysNo.Text = Request["SysNo"];
            ReadNode(TextSysNo.Text, null);
        }
    }

    protected void ListSysFun()   //列出系統功能待選清單
    {
        string SQL = "(SELECT [資源名稱],[資源編號],[資源種類],[系統全名] FROM [View_系統資源] where [資源種類]='分類' and [所屬編號]=0)"
                + " Union (SELECT '　' + [資源名稱] AS [資源名稱],[資源編號],[資源種類],[系統全名] FROM [View_系統資源] where [資源種類]='系統')"
                + " Union (SELECT '　　' + [資源名稱] AS [資源名稱],[資源編號],[資源種類],[系統全名] FROM [View_系統資源] where [資源種類]='功能') order by [系統全名]";

        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
        Conn.Open();
        SqlCommand cmd = new SqlCommand(SQL, Conn);
        SqlDataReader dr = null;
        dr = cmd.ExecuteReader();

        while (dr.Read())
        {
            ListItem item = new ListItem();
            item.Text = dr[0].ToString();
            item.Value = dr[1].ToString(); if (dr[2].ToString() == "分類") item.Value = "-" + dr[1].ToString();
            ListSysTree.Items.Add(item);
        }
    }
    //------------產生樹狀結構------------------------------------------------------------------------------
    protected void OpenTree(TreeNode node, string kind)   //當使用者按一下開啟節點時
    {
        string SQL = "";
        if (kind == "Up") SQL = "select * from [View_系統資源] where [資源編號] in (Select [上游編號] from [View_系統迴路] where [下游編號]=" + node.Value.ToString() + ") order by [維護課別],[資源名稱]";
        if (kind == "Dn") SQL = "select * from [View_系統資源] where [資源編號] in (Select [下游編號] from [View_系統迴路] where [上游編號]=" + node.Value.ToString() + ") order by [維護課別],[資源名稱]";

        if (node.ChildNodes.Count == 0) NodeAdd(node, SQL);//子節點展開過了沒？沒有，則需要產生，否則不必重複產生
    }

    protected void TreeViewDn_TreeNodePopulate(object sender, TreeNodeEventArgs e)   //當使用者按一下開啟節點時
    {
        OpenTree(e.Node, "Dn");
    }
    protected void TreeViewUp_TreeNodePopulate(object sender, TreeNodeEventArgs e)   //當使用者按一下開啟節點時
    {
        OpenTree(e.Node, "Up");
    }

    protected void NodeAdd(TreeNode node, string SQL)    //產生有子節點之目錄節點
    {
        SqlCommand sqlQuery = new SqlCommand(SQL);
        DataSet ds = RunQuery(sqlQuery);

        if (ds.Tables.Count > 0)
        {
            foreach (DataRow row in ds.Tables[0].Rows)
            {
                TreeNode NewNode = new TreeNode(row["資源名稱"].ToString(), row["資源編號"].ToString());  //TreeNode(顯示，值)

                if (HasChild(NewNode.Value, SQL))
                {
                    NewNode.PopulateOnDemand = true;
                    NewNode.SelectAction = TreeNodeSelectAction.SelectExpand;   //觸發TreeNodeExpanded
                    NewNode.Expanded = false;   //不展開節點                    
                }
                else
                {
                    NewNode.PopulateOnDemand = false; //還有子節點，使用者按一下會觸發TreeNodePopulate事件
                }

                NewNode.ToolTip = row["資源編號"].ToString() + ". " + row["資源功能"].ToString() + " (" + row["資源種類"].ToString() + ")" + " [" + FormatChilds(row["功能主機"].ToString()) + "]";
                node.ChildNodes.Add(NewNode);
            }
        }

        sqlQuery.Cancel(); ds.Dispose();
    }

    protected string FormatChilds(string PkNos)    //格式化下游字串
    {
        string cfg = PkNos.Replace(",0,", ",").Replace(",,", ",");
        if (cfg == "" | cfg == ",") cfg = "無功能主機";
        if (cfg.Substring(cfg.Length - 1) == ",") cfg = cfg.Substring(0, cfg.Length - 1);
        if (cfg.Substring(0, 1) == ",") cfg = cfg.Substring(1);

        return (cfg);
    }

    protected Boolean HasChild(string SysNo, string SQL)   //判斷是否還有子節點
    {
        Boolean ChildTF = false;
        string HasSQL = "Select * from [系統資源] where [所屬編號]=" + SysNo;
        if (SQL.IndexOf("Select [下游編號]") > 0) HasSQL = "Select * from [View_系統迴路] where [上游編號]=" + SysNo;
        if (SQL.IndexOf("Select [上游編號]") > 0) HasSQL = "Select * from [View_系統迴路] where [上游編號]<>0 and [下游編號]=" + SysNo;

        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
        Conn.Open();
        SqlCommand cmd = new SqlCommand(HasSQL, Conn);
        SqlDataReader dr = null;
        dr = cmd.ExecuteReader();

        try { if (dr.Read()) ChildTF = true; }
        catch { ChildTF = false; }

        cmd.Cancel(); cmd.Dispose(); dr.Close(); Conn.Close(); Conn.Dispose();
        return (ChildTF);
    }

    protected DataSet RunQuery(SqlCommand sqlQuery) //讀取節點資訊
    {
        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
        SqlDataAdapter dbAdapter = new SqlDataAdapter();
        dbAdapter.SelectCommand = sqlQuery;
        sqlQuery.Connection = Conn;
        DataSet QueryDataSet = new DataSet();
        dbAdapter.Fill(QueryDataSet);
        dbAdapter.Dispose(); Conn.Close(); Conn.Dispose();
        return (QueryDataSet);
    }

    //------------選取節點而顯示該節點設定值------------------------------------------------------------------------------
    protected void TreeViewDn_SelectedNodeChanged(object sender, EventArgs e)
    {
        try
        {
            ReadNode(TreeViewDn.SelectedValue, TreeViewDn.SelectedNode);
        }
        catch (Exception ex)
        {

        }
    }

    protected void TreeViewDn_TreeNodeCollapsed(object sender, TreeNodeEventArgs e)
    {
        try
        {
            ReadNode(e.Node.Value, e.Node);
        }
        catch (Exception ex)
        {

        }
    }

    protected void TreeViewDn_TreeNodeExpanded(object sender, TreeNodeEventArgs e)
    {
        try
        {
            ReadNode(e.Node.Value, e.Node);
        }
        catch (Exception ex)
        {

        }
    }

    protected void ReadNode(string SysNo, TreeNode node) //讀取節點以顯示設定值，產生上游樹狀流程，以及顯示相關主機資訊
    {
        if (node != null) node.Selected = true;

        TextSysNo.Text = SysNo;
        lblSys.Text = "";
        ListUpSysTree.Items.Clear();
        ListDnSysTree.Items.Clear();
        TreeViewUp.Nodes[0].ChildNodes.Clear();
        TreeViewUp.Nodes[0].Text = "<font color='blue' size='large'><b>系統迴路(往上)</b></font>";
        TreeViewUp.Nodes[0].ToolTip = "-1. 系統迴路根目錄";
        TreeViewUp.Nodes[0].Value = "-1";

        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
        Conn.Open();
        SqlCommand cmd = new SqlCommand("select * from [View_系統資源] where [資源編號]=" + SysNo, Conn);
        SqlDataReader dr = null;
        dr = cmd.ExecuteReader();

        if (dr.Read())
        {
            lblSys.Text = "資源名稱：" + dr["資源名稱"].ToString() + "(" + dr["資源編號"].ToString() + ")<br />"
                + "資源種類：" + dr["資源種類"].ToString() + "<br />"
                + "資源路徑：" + dr["系統全名"].ToString() + "<br />"
                + "資源功能：" + dr["資源功能"].ToString() + "<br />"
                + "維護人員：" + dr["維護人員"].ToString() + "<br />"
                + "備註說明：" + dr["備註說明"].ToString(); 
            
            TreeViewUp.Nodes[0].Text = "<font color='blue' size='large'><b>" + dr["資源名稱"].ToString() + "(往上)</b></font>";
            TreeViewUp.Nodes[0].ToolTip = SysNo + ". " + dr["資源功能"].ToString() + " (" + dr["資源種類"].ToString() + ")" + " [" + FormatChilds(dr["功能主機"].ToString()) + "]";
            TreeViewUp.Nodes[0].Value = SysNo;
            OpenTree(TreeViewUp.Nodes[0], "Up");

            ReadSysTree(TextSysNo.Text, "Up", ListUpSysTree);   //讀取系統迴路清單
            ReadSysTree(TextSysNo.Text, "Dn", ListDnSysTree);   //讀取系統迴路清單
            ReadHosts(TextSysNo.Text);   //讀取功能主機清單
        }
        cmd.Cancel(); cmd.Dispose(); dr.Close(); Conn.Close(); Conn.Dispose();
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
    protected void AddMsg(string strMsg)
    {
        Literal Msg = new Literal();
        Msg.Text = strMsg;
        Page.Controls.Add(Msg);
    }
    //-------------異動執行-----------------------------------------------------------------------------
    protected void InsLifeLog(string SQL, string PkNo,string OldMt) //寫入生命履歷
    {
        string LifeNo = GetPKNo("履歷編號", "生命履歷").ToString(); //履歷編號
        string TblName = "系統迴路";    //表格名稱
        string Mt = GetValue("IDMS", "select [維護人員] from [系統資源] where [資源編號]=" + PkNo);    //維護人員
        //string OldMt = Mt;    //原負責人
        string OldKeeper = "";    //原保管人
        string UN = Session["UserName"].ToString();   //登入的UserName：異動人員
        string LiftDT = DateTime.Now.ToString("yyyy/MM/dd HH:mm");  //異動日期
        string LifeIP = Request.ServerVariables["REMOTE_ADDR"].ToString();
        ExecDbSQL("Insert into [生命履歷] values(" + LifeNo + ",'" + TblName + "'," + PkNo + ",'" + SQL.Replace("'", "''") + "','" + Mt + "','" + OldMt + "','" + OldKeeper + "','" + UN + "','" + LiftDT + "','" + LifeIP + "')");
    }

    protected int GetPKNo(string PKfield, string PKtbl) //取得主鍵編號
    {
        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
        Conn.Open();
        SqlCommand cmd = new SqlCommand("select max(" + PKfield + ") from " + PKtbl, Conn);
        SqlDataReader dr = null;
        dr = cmd.ExecuteReader();

        int PkNo = 1;
        try { if (dr.Read()) PkNo = int.Parse(dr[0].ToString()) + 1; }
        catch { }

        cmd.Cancel(); cmd.Dispose(); dr.Close(); Conn.Close(); Conn.Dispose();

        return (PkNo);
    }

    protected void ExecDbSQL(string SQL) //執行資料庫異動
    {
        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
        Conn.Open();
        SqlCommand cmd = new SqlCommand(SQL, Conn);
        cmd.ExecuteNonQuery();
        cmd.Cancel(); cmd.Dispose(); Conn.Close(); Conn.Dispose();
    }

    protected Boolean RightCheck() //是否有權修改資料
    {
        string UserID = Session["UserID"].ToString().ToLower();
        string UserName = Session["UserName"].ToString();   //登入的UserName
        string UnitName = Session["UnitName"].ToString();   //登入的UnitName
        string Mt = GetValue("IDMS", "select [維護人員] from [系統資源] where [資源編號]=" + TextSysNo.Text);  //維護人員 
        string OldMt = Mt;    //原負責人

        if (UserID != "operator" & (InGroup(UserName, Mt) | InGroup(UserName, OldMt) | UnitName == "作業管控科" | UnitName == "系統管控科")) return (true);
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
            TF = true;
            cmd.Cancel(); cmd.Dispose(); dr.Close(); Conn.Close(); Conn.Dispose();
        }

        return (TF);
    }
    protected void LinkSys_Click(object sender, EventArgs e)  //編輯架構
    {
        if (TextSysNo.Text == "" | TextSysNo.Text == "0")
        {
            Literal Msg = new Literal();
            Msg.Text = "<script>alert('請先選取左方樹狀節點以編輯系統資源！');</script>";
            Page.Controls.Add(Msg);
        }
        else OpenExecWindow("window.open('SysEdit.aspx?SysNo=" + TextSysNo.Text + "','_self');");
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

    //系統迴路說明-------------------------------------------------------------------------------------------------------
    protected void ListSysTree_SelectedIndexChanged(object sender, EventArgs e)
    {
        lblSysTree.Text = GetSysAlias(ListSysTree.SelectedValue);
    }
    protected void ListUpSysTree_SelectedIndexChanged(object sender, EventArgs e)
    {
        lblSysTree.Text = GetSysAlias(ListUpSysTree.SelectedValue);
    }
    protected void ListDnSysTree_SelectedIndexChanged(object sender, EventArgs e)
    {
        lblSysTree.Text = GetSysAlias(ListDnSysTree.SelectedValue);
    }
    protected string GetSysAlias(string SysNo)   //取得系統相關資訊
    {
        return (GetValue("IDMS", "select '(' + CAST([資源編號] AS nvarchar) +')' + [資源名稱] + ' ' + [資源功能] + ' -- ' + [備註說明] as [系統說明] from [系統資源] where [資源編號]=" + SysNo));
    }
    //新增系統迴路-------------------------------------------------------------------------------------------------------
    protected void LinkUpSysTreeAdd_Click(object sender, EventArgs e)
    {
        SysTree_Add(TextSysNo.Text, "Up", ListUpSysTree);
    }
    protected void LinkDnSysTreeAdd_Click(object sender, EventArgs e)
    {
        SysTree_Add(TextSysNo.Text, "Dn", ListDnSysTree);
    }
    protected void SysTree_Add(string SysNo, string UpDn, ListBox ListName)  //新增系統迴路
    {
        Literal Msg = new Literal();
        string Kind = GetValue("IDMS", "select [資源種類] from [系統資源] where [資源編號]=" + TextSysNo.Text);

        if (TextSysNo.Text == "" | TextSysNo.Text == "0") Msg.Text = "<script>alert('請先選定左方樹狀結構的系統資源後再設定系統迴路！');</script>";
        else if (HasSysTree(ListSysTree.SelectedValue, ListName)) Msg.Text = "<script>alert('該系統迴路已存在，您不必再新增！');</script>";
        else if (ListSysTree.SelectedValue == SysNo) Msg.Text = "<script>alert('您不能新增自己為系統迴路！');</script>";
        else if (Kind !="系統" & Kind !="功能") Msg.Text = "<script>alert('只有系統或功能可設定上下游！');</script>";
        else
        {
            if (ListSysTree.SelectedValue != "")
            {
                if (RightCheck())
                {
                    string TreeNos = "",UpNo = "", DnNo = "", UpDnC = "";
                    if (UpDn == "Up")
                    {
                        UpNo = ListSysTree.SelectedValue; DnNo = TextSysNo.Text; UpDnC = "上游";
                        TreeNos = GetTreeNos("View_系統迴路", SysNo, "Dn") ;
                    }
                    else
                    {
                        UpNo = TextSysNo.Text; DnNo = ListSysTree.SelectedValue; UpDnC = "下游";
                        TreeNos = GetTreeNos("View_系統迴路", SysNo, "Up");
                    }                   
                    
                    if (("," + TreeNos + ",").IndexOf("," + ListSysTree.SelectedValue + ",") >= 0)
                    {
                        if (UpDn == "Up") Msg.Text = "<script>alert('不能把節點之下游樹系中任一節點再新增為上游！');</script>";
                        else Msg.Text = "<script>alert('不能把節點之上游樹系中任一節點再新增為下游！');</script>";
                    }
                    else
                    {
                        string SQL = "Insert into [系統迴路] values(" + UpNo + "," + DnNo + ")";
                        ExecDbSQL(SQL);
                        string SysName = GetValue("IDMS", "select [資源名稱] from [系統資源] where [資源編號]=" + TextSysNo.Text);
                        string OldMt = "";
                        if (UpDn == "Up") OldMt = GetValue("IDMS", "select [維護人員] from [系統資源] where [資源編號]=" + UpNo);
                        else OldMt = GetValue("IDMS", "select [維護人員] from [系統資源] where [資源編號]=" + DnNo);
                        InsLifeLog("新增 [" + SysName + "] 的" + UpDnC + "系統迴路 [" + ListSysTree.SelectedItem.Text.Trim() + "]：" + SQL, TextSysNo.Text,OldMt);
                        ReadSysTree(TextSysNo.Text, UpDn, ListName);   //讀取系統迴路清單
                        Msg.Text = "<script>alert('新增" + UpDnC + "系統迴路已完成！');</script>";
                    }
                }
                else Msg.Text = "<script>alert('您沒有新增系統迴路的權限！');</script>";
            }
            else Msg.Text = "<script>alert('請先選取您要新增的候選系統迴路！');</script>";
        }
        Page.Controls.Add(Msg);
    }

    protected Boolean HasSysTree(string SysNo, ListBox ListName) //是否有重複的系統迴路
    {
        Boolean TF = false;
        for (int i = 0; i < ListName.Items.Count; i++)
        {
            if (ListName.Items[i].Value == SysNo) TF = true;
        }

        return (TF);
    }

    protected string GetTreeNos(string tbl, string PkNo, string UpDn)   //依某節點取得所有樹系上下游編號
    {
        string SQL = "select [下游編號] from [" + tbl + "] where [上游編號]=" + PkNo;
        if (UpDn == "Up") SQL = "select [上游編號] from [" + tbl + "] where [下游編號]=" + PkNo;
        string cfg = "";

        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
        Conn.Open();
        SqlCommand cmd = new SqlCommand(SQL, Conn);
        SqlDataReader dr = null;
        dr = cmd.ExecuteReader();
        while (dr.Read()) cfg = cfg + dr[0].ToString() + "," + GetTreeNos(tbl, dr[0].ToString(), UpDn) + ",";

        cmd.Cancel(); cmd.Dispose(); dr.Close(); Conn.Close(); Conn.Dispose();

        return (FormatChilds(cfg));
    }
    //刪除系統迴路-------------------------------------------------------------------------------------------------------
    protected void LinkUpSysTreeDel_Click(object sender, EventArgs e)
    {
        SysTree_Del(TextSysNo.Text, "Up", ListUpSysTree);
    }
    protected void LinkDnSysTreeDel_Click(object sender, EventArgs e)
    {
        SysTree_Del(TextSysNo.Text, "Dn", ListDnSysTree);
    }
    protected void SysTree_Del(string SysNo, string UpDn, ListBox ListName)
    {
        Literal Msg = new Literal();

        if (TextSysNo.Text != "" & TextSysNo.Text != "0" & ListName.SelectedValue != "")
        {
            if (RightCheck())
            {
                string UpNo = TextSysNo.Text, DnNo = ListName.SelectedValue, UpDnC = "下游";
                if (UpDn == "Up")
                {
                    UpNo = ListName.SelectedValue; DnNo = TextSysNo.Text; UpDnC = "上游";
                }
                string SQL = "delete from [系統迴路] where [上游編號]=" + UpNo + " and [下游編號]=" + DnNo;

                string SysName = GetValue("IDMS", "select [資源名稱] from [系統資源] where [資源編號]=" + TextSysNo.Text);
                string OldMt = "";
                if (UpDn == "Up") OldMt = GetValue("IDMS", "select [維護人員] from [系統資源] where [資源編號]=" + UpNo);
                else OldMt = GetValue("IDMS", "select [維護人員] from [系統資源] where [資源編號]=" + DnNo);
                InsLifeLog("刪除 [" + SysName + "] 的" + UpDnC + "系統迴路 [" + ListName.SelectedItem.Text + "]：" + SQL, TextSysNo.Text,OldMt);
                ExecDbSQL(SQL);
                ReadSysTree(TextSysNo.Text, UpDn, ListName);   //讀取系統迴路清單
                Msg.Text = "<script>alert('刪除" + UpDnC + "系統迴路已完成！');</script>";
            }
            else Msg.Text = "<script>alert('您沒有刪除系統迴路的權限！');</script>";
        }
        else Msg.Text = "<script>alert('請先選取您要刪除的系統迴路！');</script>";

        Page.Controls.Add(Msg);
    }
    //讀取系統迴路清單-------------------------------------------------------------------------------------------------------
    protected void ReadSysTree(string SysNo, string UpDn, ListBox ListName)
    {
        string Source = "上游", Target = "下游";
        if (UpDn == "Up")
        {
            Source = "下游"; Target = "上游";
        }

        SqlCommand sqlQuery = new SqlCommand();
        sqlQuery.CommandText = "SELECT [資源編號],[資源名稱] FROM [系統資源] WHERE [資源編號] in "
            + " (select [" + Target + "編號] from [系統迴路] where [" + Source + "編號]=" + SysNo + ")";
        DataSet ds = RunQuery(sqlQuery);

        ListName.Items.Clear();

        if (ds.Tables.Count > 0)
        {
            foreach (DataRow row in ds.Tables[0].Rows)
            {
                ListName.Items.Add(new ListItem(row[1].ToString(), row[0].ToString()));
            }
        }

        sqlQuery.Cancel(); ds.Dispose();
    }

    //功能主機------------------------------------------------------------------------------------------------------------------
    protected void ListHosts1_SelectedIndexChanged(object sender, EventArgs e)   //顯示作業主機之說明
    {
        lblHosts.Text = GetValue("IDMS", "select '('+CAST([作業編號] AS nvarchar)+')' + [主機名稱] + ' ' + [主要功能] from [作業主機] where [作業編號]=" + ListHosts1.SelectedValue);
    }
    protected void ListHosts2_SelectedIndexChanged(object sender, EventArgs e)   //顯示作業主機之說明
    {
        lblHosts.Text = GetValue("IDMS", "select '('+CAST([作業編號] AS nvarchar)+')' + [主機名稱] + ' ' + [主要功能] from [作業主機] where [作業編號]=" + ListHosts2.SelectedValue);
    }
    protected void LinkHostHelp_Click(object sender, EventArgs e)   //功能主機之說明
    {
        lblHosts.Text = GetValue("IDMS", "select [Memo] from [Config] where [Kind]='系統資源' and [Item]='系統功能'").Replace(Convert.ToChar(13).ToString() + Convert.ToChar(10).ToString(), "<br />");
        if(sender != null) AddMsg("<script>alert('請見下方說明！');</script>");
    }
    protected void LinkHostsAdd_Click(object sender, EventArgs e)  //新增功能主機
    {
        Literal Msg = new Literal();
        string Kind = GetValue("IDMS", "select [資源種類] from [系統資源] where [資源編號]=" + TextSysNo.Text);
        string SysName = GetValue("IDMS", "select [資源名稱] from [系統資源] where [資源編號]=" + TextSysNo.Text);

        if (TextSysNo.Text == "" | TextSysNo.Text == "0") Msg.Text = "<script>alert('您尚未指定系統資源！');</script>";
        else if (HasHosts(ListHosts1.SelectedValue)) Msg.Text = "<script>alert('該功能主機已存在，您不必再新增！');</script>";
        else if (Kind != "系統" & Kind != "功能") Msg.Text = "<script>alert('只有系統或功能才可設定功能主機！');</script>";
        else
        {
            if (ListHosts1.SelectedValue != "")
            {
                if (RightCheck())
                {
                    string SQL = "Insert into [系統功能] values(" + ListHosts1.SelectedValue + "," + TextSysNo.Text + ")";
                    ExecDbSQL(SQL);
                    InsLifeLog("新增 [" + SysName + "] 的功能主機 [" + ListHosts1.SelectedItem.Text + "] ： " + SQL, TextSysNo.Text,"");
                    ReadHosts(TextSysNo.Text);   //讀取功能主機清單
                    Msg.Text = "<script>alert('新增功能主機(" + ListHosts1.SelectedItem.Text + ")已完成！您不必再重複執行存檔。');</script>";
                }
                else Msg.Text = "<script>alert('此系統(" + SysName + ")您沒有新增功能主機的權限！');</script>";
            }
            else Msg.Text = "<script>alert('請先選取您要新增的功能主機！');</script>";
        }
        Page.Controls.Add(Msg);
    }
    protected Boolean HasHosts(string ApNo) //是否有重複的功能主機
    {
        Boolean TF = false;
        for (int i = 0; i < ListHosts2.Items.Count; i++)
        {
            if (ListHosts2.Items[i].Value == ApNo) TF = true;
        }

        return (TF);
    }
    protected void LinkHostsDel_Click(object sender, EventArgs e)  //刪除功能主機
    {
        Literal Msg = new Literal();
        string SysName = GetValue("IDMS", "select [資源名稱] from [系統資源] where [資源編號]=" + TextSysNo.Text);

        if (TextSysNo.Text != "" & TextSysNo.Text != "0" & ListHosts2.SelectedValue != "")
        {
            if (RightCheck())
            {
                string SQL = "delete from [系統功能] where [作業編號]=" + ListHosts2.SelectedValue + " and [資源編號]=" + TextSysNo.Text;
                InsLifeLog("刪除 [" + SysName + "] 的功能主機 [" + ListHosts2.SelectedItem.Text + "] ： " + SQL, TextSysNo.Text,"");
                ExecDbSQL(SQL);
                Msg.Text = "<script>alert('刪除功能主機(" + ListHosts2.SelectedItem.Text + ")已完成！您不必再重複執行存檔。');</script>";
                ReadHosts(TextSysNo.Text);   //讀取功能主機清單
            }
            else Msg.Text = "<script>alert('此系統(" + SysName + ")您沒有刪除功能主機的權限！');</script>";
        }
        else Msg.Text = "<script>alert('請先選取您要刪除的功能主機！');</script>";

        Page.Controls.Add(Msg);
    }
    protected void LinkHostsEdit_Click(object sender, EventArgs e)  //編輯功能主機
    {
        if (ListHosts2.SelectedValue == "") AddMsg("<script>alert('請先選取右方您要編輯的功能主機！');</script>");
        else OpenExecWindow("window.open('ApEdit.aspx?ApNo=" + ListHosts2.SelectedValue + "','_blank');");
    }
    protected void ReadHosts(string SysNo)    //讀取功能主機清單
    {
        SqlCommand sqlQuery = new SqlCommand();
        sqlQuery.CommandText = "SELECT [作業主機].[作業編號],[作業主機].[主機名稱] FROM [作業主機],[系統功能] WHERE [作業主機].[作業編號]=[系統功能].[作業編號] and [系統功能].[資源編號] = " + SysNo + " ORDER BY [作業主機].[主機名稱]";
        DataSet ds = RunQuery(sqlQuery);

        ListHosts2.Items.Clear();

        if (ds.Tables.Count > 0)
        {
            foreach (DataRow row in ds.Tables[0].Rows)
            {
                ListHosts2.Items.Add(new ListItem(row[1].ToString(), row[0].ToString()));
            }
        }

        sqlQuery.Cancel(); ds.Dispose();
    }

    protected void BtnLife_Click(object sender, EventArgs e) //查詢生命履歷
    {
        Session["LifeSQL"] = "select * from [生命履歷] where [表格名稱]='系統迴路' and [主鍵編號]=" + TextSysNo.Text + " order by [履歷編號] desc";
        OpenExecWindow("window.open('../Life/LifeLog.aspx?Search=SysTree&Tbl=系統迴路&PK=" + TextSysNo.Text + "','_self');");
    }
}