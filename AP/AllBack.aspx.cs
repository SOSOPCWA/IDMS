﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.Configuration;
using System.Data;
using System.Data.SqlClient;

public partial class AP_AllBack : System.Web.UI.Page
{   //起始頁面---------------------------------------------------------------------
    protected void Page_Load(object sender, EventArgs e)
    {
        if (!Page.IsPostBack)   //網頁非回傳(第一次)執行，用於初始化
        {
            ViewState["PowerNos"] = "0"; ViewState["NetNos"] = "0"; ViewState["DevNos"] = "0"; ViewState["ApNos"] = "0"; ViewState["SysNos"] = "0";
        }
    }
    //產生樹狀結構------------------------------------------------------------------------------
    protected void NodeAdd(TreeNode node, string SQL, string ChildSQL, string kind, string NoName, string ColName, string ColFunc)    //產生有子節點之目錄節點
    {   //node：選取節點,SQL：新增節點,ChildSQL：有下層否,NoName：編號,ColName：名稱,ColFunc：功能
        SqlCommand sqlQuery = new SqlCommand(SQL);
        DataSet ds = RunQuery(sqlQuery);
        string PkNos = "", PosName = "";
        switch (kind)
        {
            case "接電迴路": 
            case "接網迴路": PkNos = ViewState["DevNos"].ToString(); break;
            case "作業主機": PkNos = ViewState["ApNos"].ToString();  break;
            case "系統迴路": PkNos = ViewState["SysNos"].ToString(); break;
        }

        if (ds.Tables.Count > 0)
        {
            foreach (DataRow row in ds.Tables[0].Rows)
            {
                PosName = row[ColName].ToString();
                if (kind == "接網迴路") if (row["定位方式"].ToString() == "坐標") PosName = GetValue("select [定位名稱] from [定位設定] where [定位方式]='分區' and "
                     + row["坐標X"].ToString() + " between [坐標X1] and [坐標X2] and " + row["坐標Y"].ToString() + " between [坐標Y1] and [坐標Y2]") + "：" + row["設備名稱"].ToString();
                TreeNode NewNode = new TreeNode(PosName, row[NoName].ToString());  //TreeNode(顯示，值)

                if (HasChild(ChildSQL + row[NoName].ToString()))    //有子節點者一律產生
                {
                    NewNode.PopulateOnDemand = true; //還有子節點，使用者按一下會觸發TreeNodePopulate事件
                    NewNode.SelectAction = TreeNodeSelectAction.Select;   //觸發TreeNodeExpanded
                    NewNode.ToolTip = row[NoName].ToString() + ". " + row[ColFunc].ToString();
                    node.ChildNodes.Add(NewNode);
                }
                else   //無子節點 => 系統迴路一律產生樹葉，其它則為受影響才產生
                {
                    NewNode.PopulateOnDemand = false;   //樹葉
                    if (kind == "系統迴路" | ("," + PkNos + ",").IndexOf("," + row[NoName].ToString() + ",") >= 0)
                    {
                        NewNode.ToolTip = row[NoName].ToString() + ". " + row[ColFunc].ToString();
                        node.ChildNodes.Add(NewNode);
                    }
                }

                //受影響節點以紅色顯示
                if (("," + PkNos + ",").IndexOf("," + row[NoName].ToString() + ",") >= 0) NewNode.Text = "<font color=red>" + NewNode.Text + "</font>";                

                switch (kind)    //接電接網提示所接迴路、作業主機提示系統功能、系統資源提示功能主機
                {
                    case "接電迴路":
                    case "接網迴路":
                        {
                            NewNode.ToolTip = NewNode.ToolTip + " " + ("[" + FormatChilds(row[kind].ToString()) + "]").Replace("[0]", "[無上游]");
                            break;
                        }
                    case "作業主機":
                        {
                            NewNode.ToolTip = NewNode.ToolTip + " " + GetSysFunc(row[NoName].ToString());
                            break;
                        }
                    case "系統迴路":
                        {
                            NewNode.ToolTip = NewNode.ToolTip + " " + ("[" + FormatChilds(row["功能主機"].ToString()) + "]").Replace("[0]","[無功能主機]");
                            break;
                        }
                }
            }
        }

        sqlQuery.Cancel(); ds.Dispose();
    }

    protected Boolean HasChild(string SQL)   //判斷是否還有子節點
    {
        Boolean ChildTF = false;

        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
        Conn.Open();
        SqlCommand cmd = new SqlCommand(SQL, Conn);
        SqlDataReader dr = null;
        dr = cmd.ExecuteReader();

        if (dr.Read()) ChildTF = true;

        cmd.Cancel(); cmd.Dispose(); dr.Close(); Conn.Close(); Conn.Dispose();
        return (ChildTF);
    }

    //當使用者按一下開啟節點時------------------------------------------------------------------------------
    protected void TreeViewPower_TreeNodePopulate(object sender, TreeNodeEventArgs e)
    {
        OpenTree(e.Node, "接電迴路");
    }
    protected void TreeViewNet_TreeNodePopulate(object sender, TreeNodeEventArgs e)
    {
        OpenTree(e.Node, "接網迴路");
    }
    protected void TreeViewAp_TreeNodePopulate(object sender, TreeNodeEventArgs e)
    {
        OpenTree(e.Node, "作業主機");
    }
    protected void TreeViewSys_TreeNodePopulate(object sender, TreeNodeEventArgs e)
    {
        OpenTree(e.Node, "系統迴路");
    }

    //點開節點時，才產生下層節點------------------------------------------------------------------------------
    protected void OpenTree(TreeNode node, string kind)
    {   //node：選取節點,kind：迴路種類
        if (node.ChildNodes.Count == 0) //子節點展開過了沒？沒有，則需要產生，否則不必重複產生
        {
            string SQL = "", ChildSQL = "";

            switch (kind)
            {
                case "接電迴路":
                    {
                        SQL = "select * from [View_設備管理] where [設備編號] in (select [下游編號] from [接電] where [上游編號]=" + node.Value + ")";
                        if (node.Depth == 0) SQL = SQL + " or [設備編號] in (select [上游編號] from [接電]) and [設備編號] not in (select [下游編號] from [接電])";
                        ChildSQL = "Select * from [接電] where [上游編號]=";

                        NodeAdd(node, SQL + " order by [設備名稱]", ChildSQL, "接電迴路", "設備編號", "設備名稱", "設備用途");
                        break;
                    }
                case "接網迴路":
                    {
                        SQL = "select *,[定位名稱]+'：'+[設備名稱] as [定位設備] from [View_設備管理] where [設備編號] in (select [下游編號] from [接網] where [上游編號]=" + node.Value + ")";
                        if (node.Depth == 0) SQL = SQL + " or [設備編號] in (select [上游編號] from [接網]) and [設備編號] not in (select [下游編號] from [接網])";
                        ChildSQL = "Select * from [接網] where [上游編號]=";

                        NodeAdd(node, SQL + " order by [放置地點],[設備名稱]", ChildSQL, "接網迴路", "設備編號", "定位設備", "設備用途");
                        break;
                    }
                case "作業主機":
                    {
                        if (node.Depth == 0)
                        {
                            SQL = "select * from [View_系統資源] where [所屬編號]=0 order by [資源名稱]";
                            ChildSQL = "Select * from [系統資源] where [所屬編號]=";
                            NodeAdd(node, SQL, ChildSQL, "系統迴路", "資源編號", "資源名稱", "資源功能");
                            break;
                        }
                        else if (node.Depth == 1)
                        {
                            SQL = "select * from [View_系統資源] where [資源種類]='系統' and ([資源編號]=" + node.Value + " or [資源編號] in (" + GetSubSys(node.Value,"'系統'") + ")) order by [資源名稱]";
                            ChildSQL = "Select * from [作業主機] where [系統編號]=";
                            NodeAdd(node, SQL, ChildSQL, "系統迴路", "資源編號", "資源名稱", "資源功能");
                            break;
                        }
                        else
                        {
                            SQL = "select * from [View_作業主機] where [系統編號] =" + node.Value + " order by [主機名稱]";
                            ChildSQL = "select * from [作業主機] where 1=2";
                            NodeAdd(node, SQL, ChildSQL, "作業主機", "作業編號", "主機名稱", "主要功能");
                        }

                        break;
                    }
                case "系統迴路":
                    {
                        SQL = "select * from [View_系統資源] where [資源編號] in (Select [下游編號] from [View_系統迴路] where [上游編號]=" + node.Value + ") order by [系統全名]";
                        ChildSQL = "Select * from [View_系統迴路] where [上游編號]=";
                        NodeAdd(node, SQL, ChildSQL, "系統迴路", "資源編號", "資源名稱", "資源功能");
                        break;
                    }
            }
        }
    }

    //選取節點而顯示該節點設定值------------------------------------------------------------------------------
    protected void TreeViewPower_SelectedNodeChanged(object sender, EventArgs e)
    {
        try
        {
            lbltbl.ToolTip = "接電迴路";
            TreeViewNet.SelectedNode.Selected = false;
            TreeViewAp.SelectedNode.Selected = false;
            TreeViewSys.SelectedNode.Selected = false;
        }
        catch { }
    }
    protected void TreeViewNet_SelectedNodeChanged(object sender, EventArgs e)
    {
        try
        {
            lbltbl.ToolTip = "接網迴路";
            TreeViewPower.SelectedNode.Selected = false;
            TreeViewAp.SelectedNode.Selected = false;
            TreeViewSys.SelectedNode.Selected = false;
        }
        catch { }
    }
    protected void TreeViewAp_SelectedNodeChanged(object sender, EventArgs e)
    {
        try
        {
            lbltbl.ToolTip = "作業主機";
            TreeViewPower.SelectedNode.Selected = false;
            TreeViewNet.SelectedNode.Selected = false;
            TreeViewSys.SelectedNode.Selected = false;            
        }
        catch { }
    }
    protected void TreeViewSys_SelectedNodeChanged(object sender, EventArgs e)
    {
        try
        {
            lbltbl.ToolTip = "系統迴路";
            TreeViewPower.SelectedNode.Selected = false;
            TreeViewNet.SelectedNode.Selected = false;
            TreeViewAp.SelectedNode.Selected = false;
        }
        catch { }
    }

    //讀取節點以顯示相關資訊------------------------------------------------------
    protected void ReadNode(string PkNo, TreeNode node,string Limit)
    {   
        string PowerNos="", NetNos="", DevNos = "", ApNos = "", SysNos = "";        
        if (node != null) node.Selected = true;

        switch (lbltbl.ToolTip)
        {
            case "系統迴路":
                lbltbl.Text = lbltbl.ToolTip;
                TextPkNo.Text = PkNo;
                lblName.Text = GetValue("select [資源名稱] from [系統資源] where [資源編號]=" + PkNo);
                lblFunc.Text = GetValue("select [資源功能] from [系統資源] where [資源編號]=" + PkNo);

                SysNos = PkNo;                
                SysNos = FormatChilds( GetSubSys(PkNo,"'系統','功能'") + "," + GetUpNos("View_系統迴路",PkNo));
                ViewState["SysNos"] = SysNos;

                ApNos = NosToNos(SysNos,"f2a",""); ViewState["ApNos"] = ApNos;  
                TreeViewAp.Nodes[0].ChildNodes.Clear();
                OpenTree(TreeViewAp.Nodes[0], "作業主機");

                NetNos = GetUpNos("接網", NosToNos(ApNos, "a2d", ""));
                NetNos = GetValues("select [設備編號] from [實體設備] where [設備編號] in (" + FormatChilds(NetNos) + ")");
                ViewState["NetNos"] = NetNos; ViewState["PowerNos"] = NetNos; ViewState["DevNos"] = NetNos;
                TreeViewNet.Nodes[0].ChildNodes.Clear();
                OpenTree(TreeViewNet.Nodes[0], "接網迴路");

                PowerNos = GetUpNos("接電", NetNos);
                PowerNos = GetValues("select [設備編號] from [實體設備] where [設備編號] in (" + FormatChilds(PowerNos) + ")");
                ViewState["PowerNos"] = PowerNos;
                
                DevNos = GetValues("select [設備編號] from [實體設備] where [設備編號] in (" + FormatChilds(PowerNos + "," + NetNos) + ")");
                ViewState["DevNos"] = DevNos;
                TreeViewPower.Nodes[0].ChildNodes.Clear();
                OpenTree(TreeViewPower.Nodes[0], "接電迴路");

                break;
            case "作業主機":

                if (node.Depth == 0) AddMsg("<script>alert('請選取作業主機下方節點當查詢條件！');</script>");    //根目錄
                else if (node.Depth > 2)    //作業主機
                {
                    lbltbl.Text = lbltbl.ToolTip;
                    TextPkNo.Text = PkNo;
                    lblName.Text = GetValue("select [主機名稱] from [作業主機] where [作業編號]=" + PkNo);
                    lblFunc.Text = GetValue("select [主要功能] from [作業主機] where [作業編號]=" + PkNo) + "　" + GetSysFunc(PkNo);

                    SysNos = GetValue("select [系統編號] from [作業主機] where [作業編號]=" + PkNo);
                    SysNos = FormatChilds( SysNos + "," + GetValue("select [資源編號] from [系統功能] where [作業編號]=" + PkNo));
                    SysNos = SysNos + "," + GetValue("select [所屬編號] from [系統資源] where [資源編號] in (" + SysNos + ")");                    
                    ViewState["SysNos"] = SysNos;

                    ApNos = PkNo; ViewState["ApNos"] = ApNos;

                    NetNos = GetUpNos("接網", NosToNos(ApNos, "a2d", ""));
                    NetNos = GetValues("select [設備編號] from [實體設備] where [設備編號] in (" + FormatChilds(NetNos) + ")");
                    ViewState["NetNos"] = NetNos; ViewState["PowerNos"] = NetNos; ViewState["DevNos"] = NetNos;
                    TreeViewNet.Nodes[0].ChildNodes.Clear();
                    OpenTree(TreeViewNet.Nodes[0], "接網迴路");

                    PowerNos = GetUpNos("接電", NetNos);
                    PowerNos = GetValues("select [設備編號] from [實體設備] where [設備編號] in (" + FormatChilds(PowerNos) + ")");
                    ViewState["PowerNos"] = PowerNos;
                    
                    DevNos = GetValues("select [設備編號] from [實體設備] where [設備編號] in (" + FormatChilds(PowerNos + "," + NetNos) + ")");
                    ViewState["DevNos"] = DevNos;
                    TreeViewPower.Nodes[0].ChildNodes.Clear();
                    OpenTree(TreeViewPower.Nodes[0], "接電迴路");                   
                }
                else //系統名稱
                {
                    lbltbl.Text = lbltbl.ToolTip;
                    TextPkNo.Text = PkNo;
                    lblName.Text = GetValue("select [資源名稱] from [系統資源] where [資源編號]=" + PkNo);
                    lblFunc.Text = GetValue("select [資源功能] from [系統資源] where [資源編號]=" + PkNo);

                    if (node.Depth == 1) SysNos = GetSubSys(PkNo, "'系統'");
                    else SysNos = FormatChilds(PkNo + "," + GetValue("select [所屬編號] from [系統資源] where [資源編號]=" + PkNo));
                    ViewState["SysNos"] = SysNos;

                    ApNos = GetValues("select [作業編號] from [作業主機] where [系統編號] in (" + GetSubSys(PkNo, "'系統'") + ")");
                    ViewState["ApNos"] = ApNos;

                    NetNos = GetUpNos("接網", NosToNos(ApNos, "a2d", ""));
                    NetNos = GetValues("select [設備編號] from [實體設備] where [設備編號] in (" + FormatChilds(NetNos) + ")");
                    ViewState["NetNos"] = NetNos; ViewState["PowerNos"] = NetNos; ViewState["DevNos"] = NetNos;
                    TreeViewNet.Nodes[0].ChildNodes.Clear();
                    OpenTree(TreeViewNet.Nodes[0], "接網迴路");

                    PowerNos = GetUpNos("接電", NetNos);
                    PowerNos = GetValues("select [設備編號] from [實體設備] where [設備編號] in (" + FormatChilds(PowerNos) + ")");
                    ViewState["PowerNos"] = PowerNos;
                    
                    DevNos = GetValues("select [設備編號] from [實體設備] where [設備編號] in (" + FormatChilds(PowerNos + "," + NetNos) + ")");
                    ViewState["DevNos"] = DevNos;
                    TreeViewPower.Nodes[0].ChildNodes.Clear();
                    OpenTree(TreeViewPower.Nodes[0], "接電迴路");
                }
                break;
            case "接網迴路":
            case "接電迴路":
                lbltbl.Text = lbltbl.ToolTip;
                TextPkNo.Text = PkNo;                
                lblName.Text = GetValue("select [設備名稱] from [實體設備] where [設備編號]=" + PkNo);
                lblFunc.Text = GetValue("select [設備用途] from [實體設備] where [設備編號]=" + PkNo);

                ViewState["SysNos"] = SysNos; ViewState["ApNos"] = ApNos;                

                if (lbltbl.ToolTip == "接網迴路")
                {
                    NetNos = GetUpNos("接網", PkNo);
                    PowerNos = GetUpNos("接電", NetNos);
                    PowerNos = GetValues("select [設備編號] from [實體設備] where [設備編號] in (" + FormatChilds(PowerNos) + ")");

                    ViewState["NetNos"] = NetNos; ViewState["PowerNos"] = PowerNos;

                    DevNos = GetValues("select [設備編號] from [實體設備] where [設備編號] in (" + FormatChilds(PowerNos + "," + NetNos) + ")");
                    ViewState["DevNos"] = DevNos;
                    TreeViewPower.Nodes[0].ChildNodes.Clear();
                    OpenTree(TreeViewPower.Nodes[0], "接電迴路");
                }
                else   
                {
                    PowerNos = GetUpNos("接電", PkNo);
                    PowerNos = GetValues("select [設備編號] from [實體設備] where [設備編號] in (" + FormatChilds(PowerNos) + ")");

                    ViewState["NetNos"] = NetNos; ViewState["PowerNos"] = PowerNos;

                    DevNos = GetValues("select [設備編號] from [實體設備] where [設備編號] in (" + FormatChilds(PowerNos + "," + NetNos) + ")");
                    ViewState["DevNos"] = DevNos;
                }                
                                
                break;                       
        }

        AddMsg("<script>alert('查詢完成！您可選取迴路根節點後再按展開之連結，即可看到標為紅色的根因節點。');</script>");
    }    

    protected string GetSysFunc(string ApNo)   //列舉系統功能
    {
        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
        Conn.Open();
        SqlCommand cmd = new SqlCommand("select [資源名稱] from [系統資源] where [資源編號] in (select [資源編號] from [系統功能] where [作業編號]=" + ApNo + ")", Conn);
        SqlDataReader dr = null;
        dr = cmd.ExecuteReader();

        string cfg = "";
        while (dr.Read()) cfg = cfg + "[" + dr[0].ToString() + "]、";

        if (cfg == "") cfg = "[無系統功能]";
        else cfg = cfg.Substring(0,cfg.Length-1);

        cmd.Cancel(); cmd.Dispose(); dr.Close(); Conn.Close(); Conn.Dispose();

        return (cfg);
    }

    //取得下游受影響的集合---------------------------------------------------------------------------------------------------   
    protected string NosToNos(string Nos, string UpKind, string Limit)  //將所有異常：實體設備->作業編號，或作業主機->系統編號
    {
        string SQL="";

        switch (UpKind)
        {
            case "d2a": //實體設備轉作業主機
                {
                    SQL = "select distinct [作業編號] from [View_通用設備] where [作業編號] is not null"
                        + " and ([設備型態]='系統設備' or [作業編號] in (select [作業編號] from [系統功能])) and [設備編號] in (" + Nos + ")";
                    if (Limit == "Y") SQL = SQL + " and [作業編號] in (select distinct [作業編號] from [系統功能])";
                    if (Limit == "") SQL = SQL + " and ([作業狀態]='已上線' or [作業編號] in (select distinct [作業編號] from [系統功能]))";
                    break;
                }
            case "s2a": //系統名稱轉作業主機
                {
                    SQL = "select distinct [作業編號] from [作業主機] where [系統編號] in (" + Nos + ")";
                    if (Limit == "Y") SQL = SQL + " and [作業編號] in (select distinct [作業編號] from [系統功能])";
                    if (Limit == "") SQL = SQL + " and ([作業狀態]='已上線' or [作業編號] in (select distinct [作業編號] from [系統功能]))";
                    break;
                }
            case "a2s": SQL = "select distinct [系統編號] from [作業主機] where [作業編號] in (" + Nos + ")"; break; //作業主機轉系統名稱
            case "a2f": //作業主機轉系統功能
                {
                    if (Limit == "Y") SQL = "select distinct [資源編號] from [系統功能] where [資源編號] not in (select distinct [資源編號] from [系統功能] where [作業編號] not in (" + Nos + "))";
                    else SQL = "select distinct [資源編號] from [系統功能] where [作業編號] in (" + Nos + ")";
                    break;
                }
            case "f2a": SQL = "select distinct [作業編號] from [系統功能] where [資源編號] in (" + Nos + ")"; break; //系統功能轉作業主機
            case "a2d": SQL = "select distinct [設備編號] from [作業主機] where [作業編號] in (" + Nos + ")"; break; //作業主機轉實體設備
        }

        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
        Conn.Open();
        SqlCommand cmd = new SqlCommand(SQL, Conn);
        SqlDataReader dr = null;
        dr = cmd.ExecuteReader();

        string cfg = "";
        while (dr.Read()) cfg = cfg + dr[0].ToString() + ",";

        cmd.Cancel(); cmd.Dispose(); dr.Close(); Conn.Close(); Conn.Dispose();

        return (FormatChilds(cfg));
    }

    protected string GetSubSys(string SysNo,string Kinds)   //取得系統資源所有下層系統
    {
        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
        Conn.Open();
        SqlCommand cmd = new SqlCommand("select [資源編號] from [系統資源] where [資源種類] in (" + Kinds + ") and [所屬編號]=" + SysNo, Conn);
        SqlDataReader dr = null;
        dr = cmd.ExecuteReader();

        string cfg = SysNo;
        while (dr.Read()) cfg = cfg + "," + dr[0].ToString() + "," + GetSubSys(dr[0].ToString(),Kinds) + ",";

        cmd.Cancel(); cmd.Dispose(); dr.Close(); Conn.Close(); Conn.Dispose();

        return (FormatChilds(cfg));
    }

    protected string GetUpNos(string tbl,string Nos)   //取得所有上游編號
    {
        string SQL = "select [上游編號] from [" + tbl + "] where [下游編號] in (" + Nos + ")";
        
        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
        Conn.Open();
        SqlCommand cmd = new SqlCommand(SQL, Conn);
        SqlDataReader dr = null;
        dr = cmd.ExecuteReader();

        string cfg = Nos;
        while (dr.Read()) cfg = cfg + "," + dr[0].ToString() + "," + GetUpNos(tbl,dr[0].ToString()) + ",";

        cmd.Cancel(); cmd.Dispose(); dr.Close(); Conn.Close(); Conn.Dispose();

        return (FormatChilds(cfg));
    }

    protected string FormatChilds(string PkNos)    //格式化下游字串
    {
        string cfg = PkNos.Replace(",0,", ",").Replace(",,", ",");
        if (cfg == "" | cfg == ",") cfg = "0";
        if (cfg.Substring(cfg.Length - 1) == ",") cfg = cfg.Substring(0, cfg.Length - 1);
        if (cfg.Substring(0, 1) == ",") cfg = cfg.Substring(1);

        return (cfg);
    }    

    //公用函數-------------------------------------------------------------------------------------------------------
    protected DataSet RunQuery(SqlCommand sqlQuery)
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

    protected string GetValue(string SQL)   //取得單一資料
    {
        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
        Conn.Open();
        SqlCommand cmd = new SqlCommand(SQL, Conn);
        SqlDataReader dr = null;
        dr = cmd.ExecuteReader();
        string cfg = ""; if (dr.Read()) cfg = dr[0].ToString();

        cmd.Cancel(); cmd.Dispose(); dr.Close(); Conn.Close(); Conn.Dispose();

        return (cfg);
    }

    protected string GetValues(string SQL)   //取得多重資料
    {
        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
        Conn.Open();
        SqlCommand cmd = new SqlCommand(SQL, Conn);
        SqlDataReader dr = null;
        dr = cmd.ExecuteReader();
        string cfg = ""; while (dr.Read()) cfg = cfg + dr[0].ToString() + ",";

        cmd.Cancel(); cmd.Dispose(); dr.Close(); Conn.Close(); Conn.Dispose();

        return (FormatChilds(cfg));
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

    protected void AddMsg(string strMsg)
    {
        Literal Msg = new Literal();
        Msg.Text = strMsg;
        Page.Controls.Add(Msg);
    }

    //個別迴路操作(展開)-------------------------------------------------------------------------------------------------------
    protected void BtnPowerOpen_Click(object sender, EventArgs e)
    {
        if (TreeViewPower.SelectedValue == "") AddMsg("<script>alert('請先選取樹狀節點！');</script>");
        else TreeViewPower.SelectedNode.ExpandAll();
    }
    protected void BtnNetOpen_Click(object sender, EventArgs e)
    {
        if (TreeViewNet.SelectedValue == "") AddMsg("<script>alert('請先選取樹狀節點！');</script>");
        else TreeViewNet.SelectedNode.ExpandAll();
    }
    protected void BtnApOpen_Click(object sender, EventArgs e)
    {
        if (TreeViewAp.SelectedValue == "") AddMsg("<script>alert('請先選取樹狀節點！');</script>");
        else TreeViewAp.SelectedNode.ExpandAll();
    }
    protected void BtnSysOpen_Click(object sender, EventArgs e)
    {
        if (TreeViewSys.SelectedValue == "") AddMsg("<script>alert('請先選取樹狀節點！');</script>");
        else TreeViewSys.SelectedNode.ExpandAll();
    }
    //個別迴路操作(閉合)-------------------------------------------------------------------------------------------------------
    protected void BtnPowerClose_Click(object sender, EventArgs e)
    {
        if (TreeViewPower.SelectedValue == "") AddMsg("<script>alert('請先選取樹狀節點！');</script>");
        else
        {
            TreeViewPower.SelectedNode.CollapseAll();
            ClearDnNode(TreeViewPower.SelectedNode,lbltbl.ToolTip);
        }
    }
    protected void BtnNetClose_Click(object sender, EventArgs e)
    {
        if (TreeViewNet.SelectedValue == "") AddMsg("<script>alert('請先選取樹狀節點！');</script>");
        else
        {
            TreeViewNet.SelectedNode.CollapseAll();
            ClearDnNode(TreeViewNet.SelectedNode, lbltbl.ToolTip);
        }
    }
    protected void BtnApClose_Click(object sender, EventArgs e)
    {
        if (TreeViewAp.SelectedValue == "") AddMsg("<script>alert('請先選取樹狀節點！');</script>");
        else
        {
            TreeViewAp.SelectedNode.CollapseAll();
            ClearDnNode(TreeViewAp.SelectedNode, lbltbl.ToolTip);
        }
    }
    protected void BtnSysClose_Click(object sender, EventArgs e)
    {
        if (TreeViewSys.SelectedValue == "") AddMsg("<script>alert('請先選取樹狀節點！');</script>");
        else
        {
            TreeViewSys.SelectedNode.CollapseAll();
            ClearDnNode(TreeViewSys.SelectedNode, lbltbl.ToolTip);
        }
    }    
    //個別迴路操作-------------------------------------------------------------------------------------------------------
    protected void LinkEdit_Click(object sender, EventArgs e)
    {
        switch(lbltbl.ToolTip)
        {
            case "" : AddMsg("<script>alert('請先選取樹狀節點！');</script>");break;
            case "接電迴路": OpenExecWindow("window.open('../Device/DevEdit.aspx?DevNo=" + TreeViewPower.SelectedValue + "','_blank');"); break;
            case "接網迴路": OpenExecWindow("window.open('../Device/DevEdit.aspx?DevNo=" + TreeViewNet.SelectedValue + "','_blank');"); break;
            case "作業主機":
                {
                    if (TreeViewAp.SelectedNode.Depth > 2) OpenExecWindow("window.open('ApEdit.aspx?ApNo=" + TreeViewAp.SelectedValue + "','_blank');");
                    else OpenExecWindow("window.open('SysEdit.aspx?SysNo=" + TreeViewAp.SelectedValue + "','_self');");
                    break;
                }
            case "系統迴路": OpenExecWindow("window.open('SysEdit.aspx?SysNo=" + TreeViewSys.SelectedValue + "','_self');"); break;
        }
               
    }

    protected void LinkTree_Click(object sender, EventArgs e)
    {
        switch (lbltbl.ToolTip)
        {
            case "": AddMsg("<script>alert('請先選取樹狀節點！');</script>"); break;
            case "接電迴路": OpenExecWindow("window.open('../SOS/PowerTree.aspx?DevNo=" + TreeViewPower.SelectedValue + "','_self');"); break;
            case "接網迴路": OpenExecWindow("window.open('../SOS/NetTree.aspx?DevNo=" + TreeViewNet.SelectedValue + "','_self');"); break;
            case "作業主機":
                {
                    if (TreeViewAp.SelectedNode.Depth > 2) AddMsg("<script>alert('作業主機沒有樹狀迴路！');</script>");
                    else OpenExecWindow("window.open('SysTree.aspx?SysNo=" + TreeViewAp.SelectedValue + "','_self');");
                    break;
                }
            case "系統迴路": OpenExecWindow("window.open('SysTree.aspx?SysNo=" + TreeViewSys.SelectedValue + "','_self');"); break;
        }

    }

    protected void LinkConn_Click(object sender, EventArgs e)
    {
        switch (lbltbl.ToolTip)
        {
            case "": AddMsg("<script>alert('請先選取樹狀節點！');</script>"); break;
            case "接電迴路": OpenExecWindow("window.open('../Device/TreeEdit.aspx?DevNo=" + TreeViewPower.SelectedValue + "','_blank');"); break;
            case "接網迴路": OpenExecWindow("window.open('../Device/TreeEdit.aspx?DevNo=" + TreeViewNet.SelectedValue + "','_blank');"); break;
            case "作業主機":
                {
                    if (TreeViewAp.SelectedNode.Depth > 2) AddMsg("<script>alert('作業主機沒有迴路設定！');</script>");
                    else OpenExecWindow("window.open('SysEdit.aspx?SysNo=" + TreeViewAp.SelectedValue + "','_self');");
                    break;
                }
            case "系統迴路": OpenExecWindow("window.open('SysTree.aspx?SysNo=" + TreeViewSys.SelectedValue + "','_self');"); break;
        }

    }
    //查詢與報表---------------------------------------------------------------------------------------------------
    protected void BtnSearch_Click(object sender, EventArgs e)  //整合迴路查詢
    {
        //lblName.ForeColor = System.Drawing.Color.Green;
        //lblFunc.ForeColor = System.Drawing.Color.Green;
        
        try
        {
            switch (lbltbl.ToolTip)
            {
                case "接電迴路":
                    {
                        ReadNode(TreeViewPower.SelectedValue, TreeViewPower.SelectedNode, "N");
                        ClearDnNode(TreeViewNet.Nodes[0], "接網迴路");
                        ClearDnNode(TreeViewAp.Nodes[0], "作業主機");
                        ClearDnNode(TreeViewSys.Nodes[0], "系統迴路");

                        ClearRedNode(TreeViewPower.Nodes[0]);
                        SetRedNode(TreeViewPower,TreeViewPower.Nodes[0], ViewState["PowerNos"].ToString());
                        ClearDnNode(TreeViewPower.SelectedNode, lbltbl.ToolTip);                     
                        break;
                    }
                case "接網迴路":
                    {
                        ReadNode(TreeViewNet.SelectedValue, TreeViewNet.SelectedNode, "N");
                        ClearDnNode(TreeViewAp.Nodes[0], "作業主機");
                        ClearDnNode(TreeViewSys.Nodes[0], "系統迴路");

                        ClearRedNode(TreeViewNet.Nodes[0]);
                        SetRedNode(TreeViewNet,TreeViewNet.Nodes[0], ViewState["NetNos"].ToString());
                        ClearDnNode(TreeViewNet.SelectedNode, lbltbl.ToolTip); 
                        break;
                    }
                case "作業主機":
                    {
                        if (TreeViewAp.SelectedNode.Depth == 0) AddMsg("<script>alert('請選取作業主機下方節點當查詢條件！');</script>");
                        else
                        {
                            ReadNode(TreeViewAp.SelectedValue, TreeViewAp.SelectedNode, "N");
                            ClearDnNode(TreeViewSys.Nodes[0], "系統迴路");

                            ClearRedNode(TreeViewAp.Nodes[0]);
                            SetRedNode(TreeViewAp,TreeViewAp.Nodes[0], ViewState["ApNos"].ToString());
                            ClearDnNode(TreeViewAp.SelectedNode, lbltbl.ToolTip);
                        }
                        break;
                    }
                case "系統迴路":
                    {
                        if (TreeViewSys.SelectedNode.Depth == 0) AddMsg("<script>alert('請選取系統迴路下方節點當查詢條件！');</script>");
                        else
                        {
                            ReadNode(TreeViewSys.SelectedValue, TreeViewSys.SelectedNode, "N");

                            ClearRedNode(TreeViewSys.Nodes[0]);
                            SetRedNode(TreeViewSys,TreeViewSys.Nodes[0], ViewState["SysNos"].ToString());
                            ClearDnNode(TreeViewSys.SelectedNode, lbltbl.ToolTip);
                        }
                        break;
                    }
                default: AddMsg("<script>alert('請先選取樹狀節點當查詢查詢條件！');</script>"); break;
            }
        }
        catch (Exception ex) { }
    }
    protected void ClearDnNode(TreeNode node,string kind)  //查詢或閉合後要刪除以下節點
    {
        if (node.ChildNodes.Count > 0)  //不加此判斷條件，節點未展開會出錯(PopulateOnDemand 無法在已經有子系的節點上設定為 true)
        {
            node.ChildNodes.Clear();
            OpenTree(node, kind);
        }
    }
    protected void ClearRedNode(TreeNode node)  //查詢後要清除紅色異常標記(已刪節點可不理)
    {
        node.Text = node.Text.Replace("<font color=red>", "").Replace("</font>", "");
        foreach (TreeNode ChildNode in node.ChildNodes) ClearRedNode(ChildNode);
    }
    protected void SetRedNode(TreeView tree, TreeNode node,string Nos)  //查詢後要設定紅色異常標記
    {
        string str = "";

        if (tree.ID == "TreeViewAp" & node.Depth <= 2) str = ViewState["SysNos"].ToString();
        else str = Nos;

        if (("," + str + ",").IndexOf("," + node.Value + ",") >= 0) node.Text = "<font color=red>" + node.Text + "</font>"; 
        foreach (TreeNode ChildNode in node.ChildNodes) SetRedNode(tree, ChildNode, Nos);
    }

    protected void BtnReport_Click(object sender, EventArgs e)  //報表製作
    {
        if (lbltbl.Text == "") AddMsg("<script>alert('請先選取樹狀節點後執行查詢當查詢條件！');</script>");
        else
        {
            Session["Search"] = lbltbl.Text + " " + lblName.Text + "(" + TextPkNo.Text + ") " + lblFunc.Text + " 往上回溯";
            Session["DevNos"] = ViewState["DevNos"].ToString();
            Session["ApNos"] = ViewState["ApNos"].ToString();
            Session["SysNos"] = ViewState["SysNos"].ToString();
            Session["ChkSQL"] = "";
            OpenExecWindow("window.open('Report.aspx','_blank');");
        }
    }

    protected void BtnMap_Click(object sender, EventArgs e) //設備分佈圖
    {
        //使用Session["DevSQL"]傳值
        if (ViewState["DevNos"].ToString() != "" & ViewState["DevNos"].ToString() != "0")
        {
            Session["DevSQL"] = "select * FROM [View_設備管理] where [設備編號] in (" + ViewState["DevNos"].ToString() + ")";
            OpenExecWindow("window.open('../Lib/map.aspx','_blank','location=no;menubar=no;resizable=yes;scrollbars=no;status=no;toolbar=no;fullscreen=yes');");
        }
        else AddMsg("<script>alert('請先選取樹狀節點當查詢條件後按查詢！');</script>");
    }
}