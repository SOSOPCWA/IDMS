using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.Configuration;
using System.Data;
using System.Data.SqlClient;

public partial class AP_SysEdit : System.Web.UI.Page
{   //-------------起始頁面---------------------------------------------------------------------
    protected void Page_Load(object sender, EventArgs e)
    {
        if (!Page.IsPostBack)   //網頁非回傳(第一次)執行，用於初始化
        {
            TextSysNo.Text = Request["SysNo"];
            ReadHelp(); //讀取欄位說明
            ReadAssets(); //讀取資產編號下拉式選單  

            //填入預設值
            if (Request["SysNo"] == "" | Request["SysNo"] == null) TextSysNo.Text = "0";
            else ReadNode(Request["SysNo"].ToString());

            SqlDataSourceBelong.SelectCommand = "(SELECT [資源名稱],[資源編號],[系統全名] FROM [View_系統資源] where [資源種類]='分類' and [所屬編號]=0)"
                + " Union (SELECT '　' + [資源名稱] AS [資源名稱],[資源編號],[系統全名] FROM [View_系統資源] where [資源種類]='分類' and [所屬編號]<>0)"
                + " Union (SELECT '　' + [資源名稱] AS [資源名稱],[資源編號],[系統全名] FROM [View_系統資源] where [資源種類]='系統')" 
                + " Union (SELECT '　　' + [資源名稱] AS [資源名稱],[資源編號],[系統全名] FROM [View_系統資源] where [資源種類]='功能') order by [系統全名]";
        }
    }

    //僅此事件在DataSource Binding之後，而Page_Load或其它事件，均在那之前
    protected void Page_PreRenderComplete(object sender, EventArgs e)   //放在此事件，改變選項才能自動設值，放Select_Change無效
    {
        if (TextSysNo.Text == "")    //新增狀態
        {
            BtnEdit.Enabled = false; BtnDel.Enabled = false;
        }
        else //編輯狀態
        {
            BtnEdit.Enabled = true; BtnDel.Enabled = true;
        }      


        if (!IsPostBack)
        {
            if (TextSysNo.Text != "") ReadNode(TextSysNo.Text);  //讀取資料

            ShowHideHelp();  //顯示或隱藏欄位說明            
        }

        GetMtList();    //顯示維護人員清單 
        SelKind_SelectedIndexChanged(null, null);
        SelAssets_SelectedIndexChanged(null, null);   //顯示資產描述
    }

    protected void GetMtList()   //顯示維護人員清單
    {
        SqlCommand sqlQuery = new SqlCommand();
        sqlQuery.CommandText = "select [Item] from [Config] where [Kind]='" + SelMt.SelectedValue + "' order by [Mark]";
        DataSet ds = RunQuery(sqlQuery);

        lblMt.Text = "";
        if (ds.Tables.Count > 0)
        {
            foreach (DataRow row in ds.Tables[0].Rows)
            {
                lblMt.Text = lblMt.Text + " " + row[0].ToString();
            }
        }
        else
        {
            lblMt.Text = SelMt.SelectedValue;
        }

        sqlQuery.Cancel(); ds.Dispose();
    }

    protected void ReadAssets() //讀取資產清冊下拉式選單
    {
        SqlCommand sqlQuery = new SqlCommand();
        sqlQuery.CommandText = "select [Item],[Config] from [Config] where [Kind]='常用資產' or [Kind]='資產清冊' order by [Kind],[Item]";
        DataSet ds = RunQuery(sqlQuery);

        if (ds.Tables.Count > 0)
        {
            foreach (DataRow row in ds.Tables[0].Rows)
            {
                SelAssets.Items.Add(new ListItem(row[0].ToString() + " " + row[1].ToString(), row[0].ToString()));
            }
        }

        sqlQuery.Cancel(); ds.Dispose();
    }

    protected void SelAssets_SelectedIndexChanged(object sender, EventArgs e)   //顯示資產描述
    {
        lblAssets.Text = GetValue("IDMS", "select [Memo] from [Config] where ([Kind]='資產清冊' or [Kind]='常用資產') and [Item]='" + SelAssets.SelectedValue + "'");
    }

    protected void ReadHelp() //讀取欄位說明
    {
        SqlCommand sqlQuery = new SqlCommand();
        sqlQuery.CommandText = "select [Item],[Memo] from [Config] where [Kind]='系統資源' order by [Mark]";
        DataSet ds = RunQuery(sqlQuery);

        if (ds.Tables.Count > 0)
        {
            foreach (DataRow row in ds.Tables[0].Rows)
            {
                switch (row[0].ToString())
                {
                    case "資源編號": helpSysNo.Text = row[1].ToString().Replace("\r\n", "<br />"); break;
                    case "資源種類": helpKind.Text = row[1].ToString().Replace("\r\n", "<br />"); break;
                    case "資源名稱": helpSysName.Text = row[1].ToString().Replace("\r\n", "<br />"); break;
                    case "資源功能": helpFunc.Text = row[1].ToString().Replace("\r\n", "<br />"); break;
                    case "系統架構": helpBelong.Text = row[1].ToString().Replace("\r\n", "<br />"); break;
                    case "系統功能": helpHosts.Text = row[1].ToString().Replace("\r\n", "<br />"); break;
                    case "系統迴路": helpSysTree.Text = row[1].ToString().Replace("\r\n", "<br />"); break;                    
                    case "資產編號": helpAssetsNo.Text = row[1].ToString().Replace("\r\n", "<br />"); break;
                    case "維護人員": helpMt.Text = row[1].ToString().Replace("\r\n", "<br />"); break;
                    case "備註說明": helpMemo.Text = row[1].ToString().Replace("\r\n", "<br />"); break;
                    case "資料建立": helpCreate.Text = row[1].ToString().Replace("\r\n", "<br />"); break;
                    case "資料修改": helpUpdate.Text = row[1].ToString().Replace("\r\n", "<br />"); break;
                    case "資料查核": helpChecks.Text = row[1].ToString().Replace("\r\n", "<br />"); break;                    
                }
            }
        }

        sqlQuery.Cancel(); ds.Dispose();
    }
    //------------產生樹狀結構------------------------------------------------------------------------------
    protected void TreeView1_TreeNodePopulate(object sender, TreeNodeEventArgs e)   //當使用者按一下開啟節點時
    {
        TreeNode node = e.Node;

        if (node.ChildNodes.Count == 0)   //子節點展開過了沒？沒有，則需要產生，否則不必重複產生
        {
            NodeAdd(node, "select * from [View_系統資源] where [所屬編號]=" + e.Node.Value.ToString() + " order by [維護課別],[資源名稱]");
        }
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

                if (HasChild(int.Parse(NewNode.Value)))
                {
                    NewNode.PopulateOnDemand = true;
                    NewNode.SelectAction = TreeNodeSelectAction.Expand;   //觸發TreeNodeExpanded
                    NewNode.Expanded = false;   //不展開節點                    
                }
                else
                {
                    NewNode.PopulateOnDemand = false; //還有子節點，使用者按一下會觸發TreeNodePopulate事件
                }

                NewNode.ToolTip = row["資源編號"].ToString() + ". " + row["資源功能"].ToString();
                node.ChildNodes.Add(NewNode);
            }
        }

        sqlQuery.Cancel(); ds.Dispose();
    }

    protected Boolean HasChild(int SysNo)   //判斷是否還有子節點
    {
        Boolean ChildTF = false;
        string SQL = "Select * from [系統資源] where [所屬編號]=" + SysNo;
        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
        Conn.Open();
        SqlCommand cmd = new SqlCommand(SQL, Conn);
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
    protected void TreeView1_SelectedNodeChanged(object sender, EventArgs e)
    {
        try
        {
            NodeSelect(TreeView1.SelectedNode); //選取節點以顯示設定值
        }
        catch (Exception ex)
        {

        }
    }
    protected void TreeView1_TreeNodeCollapsed(object sender, TreeNodeEventArgs e)
    {
        try
        {
            NodeSelect(e.Node); //選取節點以顯示設定值
        }
        catch (Exception ex)
        {

        }
    }

    protected void TreeView1_TreeNodeExpanded(object sender, TreeNodeEventArgs e)
    {
        try
        {
            NodeSelect(e.Node); //選取節點以顯示設定值
        }
        catch (Exception ex)
        {

        }
    }

    protected void NodeSelect(TreeNode node) //選取節點以顯示設定值
    {
        node.Selected = true;
        ReadNode(node.Value);
    }

    protected void ReadNode(string SysNo) //讀取節點以顯示設定值
    {
        lblUpSysTree.Text = ""; lblDnSysTree.Text = "";

        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
        Conn.Open();
        SqlCommand cmd = new SqlCommand("select * from [View_系統資源] where [資源編號]=" + SysNo, Conn);
        SqlDataReader dr = null;
        dr = cmd.ExecuteReader();
        if (dr.Read())
        {
            TextSysNo.Text = SysNo;
            for (int i = 0; i < SelBelong.Items.Count; i++) if (SelBelong.Items[i].Value == dr["所屬編號"].ToString()) SelBelong.SelectedIndex = i;
            for (int i = 0; i < SelAssets.Items.Count; i++) if (SelAssets.Items[i].Value == dr["資產編號"].ToString()) SelAssets.SelectedIndex = i;

            SelKind.ClearSelection();
            for (int i = 0; i < SelKind.Items.Count; i++)
            {
                if (SelKind.Items[i].Value == dr["資源種類"].ToString()) SelKind.Items[i].Selected = true;
                else SelKind.Items[i].Selected = false;
            }

            TextSysName.Text = dr["資源名稱"].ToString();
            TextFunc.Text = dr["資源功能"].ToString();
            lblHosts.Text = dr["功能主機"].ToString(); if (lblHosts.Text == "") lblHosts.Text = "(無)";
            lblSysNames.Text = GetValue("IDMS", "select count(*) from [作業主機] where [系統編號]=" + TextSysNo.Text) + " 筆 (異動請至作業主機介面)";

            string MtUnit = GetMtUnit(dr["維護人員"].ToString());
            SelMtUnit.ClearSelection(); for (int i = 0; i < SelMtUnit.Items.Count; i++) if (SelMtUnit.Items[i].Value == MtUnit) SelMtUnit.Items[i].Selected = true;
            SelMt.DataBind();  //連帶選項需要觸發
            SelMt.ClearSelection(); for (int i = 0; i < SelMt.Items.Count; i++) if (SelMt.Items[i].Value == dr["維護人員"].ToString()) SelMt.Items[i].Selected = true;
            GetMtList();   //顯示維護人員清單

            TextMemo.Text = dr["備註說明"].ToString();
            LblCreateDT.Text = DateTime.Parse(dr["建立日期"].ToString()).ToString("yyyy/MM/dd HH:mm");
            LblCreater.Text = dr["建立人員"].ToString();
            LblUpdateDT.Text = DateTime.Parse(dr["修改日期"].ToString()).ToString("yyyy/MM/dd HH:mm");
            LblUpdater.Text = dr["修改人員"].ToString();

            lblChecks.Text = GetChecks("系統資源", SysNo);   //取得資料查核 
            
            ReadConn(SysNo,"Up", lblUpSysTree); //讀取系統迴路
            ReadConn(SysNo,"Dn", lblDnSysTree); //讀取系統迴路
        }
        cmd.Cancel(); cmd.Dispose(); dr.Close(); Conn.Close(); Conn.Dispose();        
    }

    protected void ReadConn(string SysNo, string UpDn, Label lbl)    //讀取系統迴路選單
    {
        string SQL = "";
        if (UpDn == "Up") SQL = "SELECT [資源編號],[資源名稱] FROM [系統資源],[系統迴路] WHERE [系統資源].[資源編號]=[系統迴路].[上游編號] and [系統迴路].[下游編號] = " + SysNo + " ORDER BY [資源名稱]";
        else SQL = "SELECT [資源編號],[資源名稱] FROM [系統資源],[系統迴路] WHERE [系統資源].[資源編號]=[系統迴路].[下游編號] and [系統迴路].[上游編號] = " + SysNo + " ORDER BY [資源名稱]";

        SqlCommand sqlQuery = new SqlCommand();
        sqlQuery.CommandText = SQL;
        DataSet ds = RunQuery(sqlQuery);

        lbl.Text = "";

        if (ds.Tables.Count > 0)
        {
            foreach (DataRow row in ds.Tables[0].Rows)
            {
                lbl.Text = lbl.Text + row[1].ToString() + "、";
            }
        }

        sqlQuery.Cancel(); ds.Dispose();
    }

    protected string GetMtUnit(string Mt) //讀取維護人員單位或分組名稱
    {
        //先讀資訊中心各課成員(各課優先)
        string MtUnit=GetValue("IDMS","SELECT [Kind] FROM [Config] WHERE [Kind] in (select [Item] from [Config] where [Kind]='資訊中心') and [Item]='" + Mt + "'");
        //再讀其它中心或維護群組(分組次之)
        if (MtUnit=="") MtUnit = GetValue("IDMS","SELECT [Kind] FROM [Config] WHERE [Kind] not in (select [Item] from [Config] where [Kind]='資訊中心') and [Item]='" + Mt + "'");

        return (MtUnit);
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
    //-------------異動執行-----------------------------------------------------------------------------
    protected void InsLifeLog(string SQL, string PkNo) //寫入生命履歷
    {
        string LifeNo = GetPKNo("履歷編號", "生命履歷").ToString(); //履歷編號
        string TblName = "系統資源";    //表格名稱
        string Mt = SelMt.SelectedValue;    //維護人員
        string OldMt = GetValue("IDMS", "select [維護人員] from [系統資源] where [資源編號]=" + PkNo);    //原負責人
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
        string Mt = SelMt.SelectedValue;  //維護人員 
        string OldMt = GetValue("IDMS", "select [維護人員] from [系統資源] where [資源編號]=" + TextSysNo.Text);    //原負責人

        if (UserID != "operator" & (InGroup(UserName, Mt) | InGroup(UserName, OldMt) | UnitName == "系統控制課" | UnitName == "電腦操作課")) return (true);
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

    protected string ChkDepend(string SysNo,string Belong,string DBaction,string Kind) //檢查系統資源相依性能否新增刪除修改
    {
        string res = "",tmp="";

        if (DBaction == "新增" | DBaction == "修改")
        {
            tmp = GetValue("IDMS", "select [資源編號] from [系統資源] where [資源編號]=" + Belong);
            if (Belong != "0" & tmp == "") res = "所屬編號對應不到系統資源";
            else if (Belong != "0" & Belong == SysNo) res = "所屬系統不能是自己";
            else
            {
                if (Kind == "系統")  //若系統，所屬編號所對應亦應為系統********暫先不查核
                {
                    //tmp = GetValue("IDMS", "select [資源種類] from [系統資源] where [資源編號]=" + Belong);
                    //if (Belong != "0" & tmp != "系統") res = "系統之所屬系統亦必須為系統，請先更正所屬系統之種類";
                }
            }
        }

        if (DBaction == "修改")
        {
            tmp = GetValue("IDMS", "select [主機名稱]+'('+CAST([作業編號] AS nvarchar)+')' from [作業主機] where [系統編號]=" + SysNo);
            if (tmp != "" & Kind != "系統") res = "作業主機" + tmp + "已引用為系統，不可改為非系統";
            else if (Belong == SysNo) res = "所屬系統不能是自己";
            else
            {
                tmp = GetDownFlow(SysNo, Belong);
                if (tmp != "") res = "所屬系統不可為下游(層)系統資源：" + tmp;
                else 
                {
                    tmp = GetValue("IDMS", "select [主機名稱]+'('+CAST([作業編號] AS nvarchar)+')' from [作業主機] where [作業編號] in (select [作業編號] from [系統功能] where [資源編號]=" + SysNo + ")");
                    if (tmp != "" & Kind != "系統" & Kind != "功能") res = "作業主機：" + tmp + "已引用為系統功能，不可改為非系統或功能";
                }
            }
        }

        if (DBaction == "刪除")
        {
            tmp = GetValue("IDMS", "select [資源名稱]+'('+CAST([資源編號] AS nvarchar)+')' from [系統資源] where [所屬編號]=" + SysNo);
            if (Belong != "0" & tmp != "") res = "系統資源：" + tmp + "已引用為所屬系統，請先修正";
            else
            {
                tmp = GetValue("IDMS", "select [資源名稱]+'('+CAST([資源編號] AS nvarchar)+')' from [系統資源] where [資源編號]=" + SysNo
                    + " and [資源編號] in (select [上游編號] from [系統迴路] Union select [下游編號] from [系統迴路])");
                if (tmp != "") res = "系統資源：" + tmp + "已被系統迴路引用為上下游，請先修正";
                else
                {
                    tmp = GetValue("IDMS", "select [主機名稱]+'('+CAST([作業編號] AS nvarchar)+')' from [作業主機] where [系統編號]=" + SysNo);
                    if (tmp != "") res = "作業主機：" + tmp + "已引用為系統名稱，請先修正";
                    else 
                    {
                        tmp = GetValue("IDMS", "select [主機名稱]+'('+CAST([作業編號] AS nvarchar)+')' from [作業主機] where [作業編號] in (select [作業編號] from [系統功能] where [資源編號]=" + SysNo + ")");
                        if (tmp != "") res = "作業主機：" + tmp + "已引用為系統功能，請先修正";
                    }
                }
            }
        }

        return (res);
    }

    protected string GetDownFlow(string SysNo, string Belong) //檢查是否為下游系統
    {
        string res = "";

        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
        Conn.Open();
        SqlCommand cmd = new SqlCommand("select [下游編號] FROM [View_系統迴路] WHERE [上游編號]=" + SysNo, Conn);
        SqlDataReader dr = null;
        dr = cmd.ExecuteReader();
        while (dr.Read())
        {
            if (dr[0].ToString() == Belong) res = GetValue("IDMS", "select [資源名稱]+'(' + CAST([資源編號] AS nvarchar)+')' from [系統資源] where [資源編號]=" + dr[0].ToString());
            else GetDownFlow(dr[0].ToString(), Belong);            
        }
        cmd.Cancel(); cmd.Dispose(); dr.Close(); Conn.Close(); Conn.Dispose();        
        
        return (res);
    }

    protected void BtnNew_Click(object sender, EventArgs e) //新增
    {
        Literal Msg = new Literal();        
        int n;

        if (!RightCheck()) Msg.Text = "<script>alert('您不是維護人員，沒有權限新增資料！');</script>";
        else if (TextSysName.Text == "") Msg.Text = "<script>alert('未輸入資源名稱！');</script>";
        else if (GetValue("IDMS", "select [資源編號] from [系統資源] where [所屬編號]=" + SelBelong.SelectedValue + " and [資源名稱]='" + TextSysName.Text + "'") != "") Msg.Text = "<script>alert('該資源名稱已存在，請使用另外的名稱當系統！');</script>";
        else if (TextFunc.Text == "") Msg.Text = "<script>alert('未輸入資源功能！');</script>";
        else if (SelMt.SelectedValue == "") Msg.Text = "<script>alert('未輸入維護人員！');</script>";
        else
        {
            string Depend = ChkDepend("0", SelBelong.SelectedValue, "新增", SelKind.SelectedValue);
            if (Depend != "") Msg.Text = "<script>alert('" + Depend + "！');</script>";
            else
            {
                string SysNo = GetPKNo("[資源編號]", "[系統資源]").ToString();
                string SQL = GetInsSQL(SysNo);
                ExecDbSQL(SQL);
                InsLifeLog(SQL, SysNo);
                //lblHosts.Text = SQL;
                ReadNode(SysNo);
                Msg.Text = "<script>alert('新增資料 [" + TextSysName.Text + "] 完成！\\n\\n" + GetChecks("系統資源", SysNo) + "');</script>";                
            }
        }

        Page.Controls.Add(Msg);
    }

    protected string GetChecks(string tbl, string PkNo)   //取得資料查核
    {
        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
        Conn.Open();
        SqlCommand cmd = new SqlCommand("select [查核結果] FROM [View_資料查核] WHERE [表格名稱]='" + tbl + "' and [主鍵編號]=" + PkNo, Conn);
        SqlDataReader dr = null;
        dr = cmd.ExecuteReader();
        string cfg = ""; if (dr.Read()) cfg = dr[0].ToString();
        cmd.Cancel(); cmd.Dispose(); dr.Close(); Conn.Close(); Conn.Dispose();

        return (cfg);
    }

    protected string GetInsSQL(string SysNo) //取得新增資料的語法
    {
        return ("insert into [系統資源] values("
            + SysNo + ","
            + SelBelong.SelectedValue + ",'"
            + SelAssets.SelectedValue + "','"
            + SelKind.SelectedValue + "','"
            + TextSysName.Text + "','"
            + TextFunc.Text + "','"
            + SelMt.SelectedValue + "','"
            + DateTime.Now.ToString("yyyy/MM/dd HH:mm") + "','"
            + Session["UserName"].ToString() + "','"
            + DateTime.Now.ToString("yyyy/MM/dd HH:mm") + "','"
            + Session["UserName"].ToString() + "','"
            + TextMemo.Text + "')");
    }

    protected void BtnEdit_Click(object sender, EventArgs e)    //修改
    {
        Literal Msg = new Literal();        
        int n;

        if (!RightCheck()) Msg.Text = "<script>alert('您不是維護人員，沒有權限新增資料！');</script>";
        else if (TextSysNo.Text == "") Msg.Text = "<script>alert('您尚未點選欲修改之資料！');</script>";
        else if (TextSysName.Text == "") Msg.Text = "<script>alert('未輸入資源名稱！');</script>";
        else if (GetValue("IDMS", "select [資源編號] from [系統資源] where [資源編號]<>" + TextSysNo.Text + " and [所屬編號]=" + SelBelong.SelectedValue 
            + " and [資源名稱]='" + TextSysName.Text + "'") != "") Msg.Text = "<script>alert('該資源名稱已存在，請使用另外的名稱當系統！');</script>";
        else if (TextFunc.Text == "") Msg.Text = "<script>alert('未輸入資源功能！');</script>";
        else if (SelMt.SelectedValue == "") Msg.Text = "<script>alert('未輸入維護人員！');</script>";
        else
        {
            string Depend = ChkDepend(TextSysNo.Text, SelBelong.SelectedValue, "修改", SelKind.SelectedValue);
            if (Depend != "") Msg.Text = "<script>alert('" + Depend + "！');</script>";
            else
            {
                InsLifeLog("修改系統資源(" + TextSysName.Text + ") ：： " + GetUpdate(int.Parse(TextSysNo.Text), "Life"), TextSysNo.Text);
                ExecDbSQL("Update [系統資源] set " + GetUpdate(int.Parse(TextSysNo.Text), "SQL") + " where [資源編號]=" + TextSysNo.Text);
                ReadNode(TextSysNo.Text); //選取節點以顯示設定值，必須DB異動之後才能執行
                Msg.Text = "<script>alert('更新資料 [" + TextSysName.Text + "] 完成！\\n\\n" + GetChecks("系統資源", TextSysNo.Text) + "');</script>";
            }
        }

        Page.Controls.Add(Msg);
    }

    protected string GetUpdateCol(string ColName, string Source, string Target, string Kind, string SQLorLife) //取得單一欄位修改資料的語法
    {
        string SQL = "";
        if (Kind == "date" & Target.Length != 10 | Kind == "datetime" & Target.Length != 16)
        {
            Kind = "null"; Target = "null";
        }

        if (Source != Target)
        {
            if (SQLorLife == "SQL")
            {
                switch (Kind)
                {
                    case "string":
                    case "date":
                    case "datetime": SQL = SQL + ",[" + ColName + "]='" + Target + "'"; break;
                    case "integer":
                    case "money": SQL = SQL + ",[" + ColName + "]=" + Target; break;
                    case "null": SQL = SQL + ",[" + ColName + "]=" + null; break;
                    default: SQL = SQL + ",[" + ColName + "]='" + Target + "'"; break;
                }
            }
            else if (SQLorLife == "Life")
            {
                if (ColName != "所屬編號")
                {
                    if (Source == "") Source = "(空白)";
                    if (Target == "") Target = "(空白)";
                    SQL = SQL + ",[" + ColName + "]：" + Source + " -> " + Target;
                }
                else
                {
                    string Sname=GetValue("IDMS","select [資源名稱] from [系統資源] where [資源編號]=" + Source); if(Sname=="") Sname="根目錄";
                    string Tname=GetValue("IDMS","select [資源名稱] from [系統資源] where [資源編號]=" + Target); if(Tname=="") Tname="根目錄";
                    SQL = SQL + ",[" + ColName + "]：" + Source + "(" + Sname + ") -> " + Target + "(" + Tname + ")";
                }
            }
        }
        return (SQL);
    }

    protected string GetUpdate(int SysNo, string SQLorLife) //取得修改資料的SQL語法
    {
        string SQL = "";
        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
        Conn.Open();
        SqlCommand cmd = new SqlCommand("select * from [系統資源] where [資源編號]=" + SysNo, Conn);
        SqlDataReader dr = null;
        dr = cmd.ExecuteReader();
        if (dr.Read())
        {
            SQL = GetUpdateCol("資源編號", dr["資源編號"].ToString(), TextSysNo.Text, "integer", SQLorLife)
                + GetUpdateCol("所屬編號", dr["所屬編號"].ToString(), SelBelong.SelectedValue, "integer", SQLorLife)
                + GetUpdateCol("資產編號", dr["資產編號"].ToString(), SelAssets.SelectedValue, "string", SQLorLife)
                + GetUpdateCol("資源種類", dr["資源種類"].ToString(), SelKind.SelectedValue, "string", SQLorLife)
                + GetUpdateCol("資源名稱", dr["資源名稱"].ToString(), TextSysName.Text, "string", SQLorLife)
                + GetUpdateCol("資源功能", dr["資源功能"].ToString(), TextFunc.Text, "string", SQLorLife)
                + GetUpdateCol("維護人員", dr["維護人員"].ToString(), SelMt.SelectedValue, "string", SQLorLife)
                + GetUpdateCol("修改日期", DateTime.Parse(dr["修改日期"].ToString()).ToString("yyyy/MM/dd HH:mm"), DateTime.Now.ToString("yyyy/MM/dd HH:mm"), "datetime", SQLorLife)
                + GetUpdateCol("修改人員", dr["修改人員"].ToString(), Session["UserName"].ToString(), "string", SQLorLife)
                + GetUpdateCol("備註說明", dr["備註說明"].ToString(), TextMemo.Text, "string", SQLorLife);

            if (SQL != "") SQL = SQL.Substring(1);
        }
        cmd.Cancel(); cmd.Dispose(); dr.Close(); Conn.Close(); Conn.Dispose();

        return (SQL);
    }

    protected void BtnDel_Click(object sender, EventArgs e) //刪除
    {
        Literal Msg = new Literal();
        string Depend = ChkDepend(TextSysNo.Text, SelBelong.SelectedValue, "刪除", SelKind.SelectedValue);

        if (!RightCheck()) Msg.Text = "<script>alert('您不是維護人員，沒有權限刪除資料，若有需求，請洽系統負責人！');</script>";
        else if (TextSysNo.Text == "") Msg.Text = "<script>alert('您尚未點選欲刪除之資料！');</script>";
        else if (GetValue("IDMS", "select count(*) from [作業主機] where [系統編號]=" + TextSysNo.Text) != "0") Msg.Text = "<script>alert('作業主機尚有本系統名稱，請先移除之！');</script>";
        else if (GetValue("IDMS", "select count(*) from [系統功能] where [資源編號]=" + TextSysNo.Text) != "0") Msg.Text = "<script>alert('功能主機尚有本系統功能，請先移除之！');</script>";
        else if (GetValue("IDMS", "select count(*) from [系統資源] where [所屬編號]=" + TextSysNo.Text) != "0") Msg.Text = "<script>alert('尚有所屬系統引用本系統，請先移除之！');</script>";
        else if (GetValue("IDMS", "select count(*) from [系統迴路] where [上游編號]=" + TextSysNo.Text + " or [下游編號]=" + TextSysNo.Text) != "0") Msg.Text = "<script>alert('尚有系統迴路引用本系統，請先移除之！');</script>";
        else if (Depend != "") Msg.Text = "<script>alert('" + Depend + "！');</script>";
        else
        {
            string SQL = "delete from [系統資源] where [資源編號]=" + TextSysNo.Text;
            InsLifeLog("刪除 [" + TextSysName.Text + "] ： " + SQL + "，原本SQL ： " + GetInsSQL(TextSysNo.Text),TextSysNo.Text); 
            ExecDbSQL(SQL);

            Msg.Text = "<script>alert('刪除系統資源(" + TextSysName.Text + ")完成！');</script>";
            TextSysNo.Text = "";
        }
        Page.Controls.Add(Msg);
    }

    protected void BtnBlank_Click(object sender, EventArgs e) //空白頁
    {
        OpenExecWindow("window.open('SysEdit.aspx','_self');");
    }

    protected void BtnLife_Click(object sender, EventArgs e) //查詢生命履歷
    {
        Session["LifeSQL"] = "select * from [生命履歷] where [表格名稱]='系統資源' and [主鍵編號]=" + TextSysNo.Text + " order by [履歷編號] desc";
        OpenExecWindow("window.open('../Life/LifeLog.aspx?Search=SysEdit&Tbl=系統資源&PK=" + TextSysNo.Text + "','_self');");
    }

    protected void SelKind_SelectedIndexChanged(object sender, EventArgs e)  //資源種類
    {
        lblKind.Text = GetValue("IDMS", "Select [Memo] from [Config] where [Kind]='資源種類' and [Item]='" + SelKind.SelectedValue + "'");
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

    protected void BtnSysTree_Click(object sender, EventArgs e)  //系統迴路
    {
        string SysNo = TextSysNo.Text; if (SysNo == "") SysNo = "0";
        Literal Msg = new Literal();

        if (SysNo == "0")
        {
            Msg.Text = "<script>alert('請指定系統資源節點！');</script>";
            Page.Controls.Add(Msg);
        }
        else OpenExecWindow("window.open('SysTree.aspx?SysNo=" + SysNo + "','_self');");
    }

    protected void LinkHideHelp_Click(object sender, EventArgs e)  //顯示隱藏欄位說明
    {
        if (LinkHideHelp.Text == "隱藏說明")
        {
            LinkHideHelp.Text = "顯示說明";
            Table1.Rows[0].Cells[3].Visible = false;
            for (int i = 1; i < Table1.Rows.Count-1; i++) Table1.Rows[i].Cells[2].Visible = false;
            WriteHideHelp("隱藏說明");
        }
        else
        {
            LinkHideHelp.Text = "隱藏說明";
            Table1.Rows[0].Cells[3].Visible = true;
            for (int i = 1; i < Table1.Rows.Count-1; i++) Table1.Rows[i].Cells[2].Visible = true;
            WriteHideHelp("");
        }
    }
    protected void WriteHideHelp(string txt)  //顯示隱藏欄位說明註記於DB
    {
        string UserName = Session["UserName"].ToString();
        ExecDbSQL("delete from [Config] where [Kind]='隱藏說明' and [Item]='" + UserName + "'");
        if (txt != "") ExecDbSQL("insert into [Config] values('隱藏說明','" + UserName + "','','','')");
    }

    protected void ShowHideHelp()  //顯示或隱藏欄位說明
    {
        string UserName = GetValue("IDMS", "select [Item] from [Config] where [Kind]='隱藏說明' and [Item]='" + Session["UserName"].ToString() + "'");
        if (UserName != "")
        {
            Table1.Rows[0].Cells[3].Visible = false;
            for (int i = 1; i < Table1.Rows.Count-1; i++) Table1.Rows[i].Cells[2].Visible = false;
            LinkHideHelp.Text = "顯示說明";
        }
        else
        {
            Table1.Rows[0].Cells[3].Visible = true;
            for (int i = 1; i < Table1.Rows.Count-1; i++) Table1.Rows[i].Cells[2].Visible = true;
            LinkHideHelp.Text = "隱藏說明";
        }
    }
}