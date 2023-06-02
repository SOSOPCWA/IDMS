using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.Configuration;
using System.Data;
using System.Data.SqlClient;

public partial class SOS_NetTree : System.Web.UI.Page
{   //-------------起始頁面---------------------------------------------------------------------
    protected void Page_PreRenderComplete(object sender, EventArgs e)   //放Page_Load無法顯示
    {
        if (!IsPostBack)   //第一次開網頁時
        {
            if (Request["DevNo"] != "" & Request["DevNo"] != null)
            {
                ReadDevice(Request["DevNo"].ToString(), TreeViewUp.Nodes[0], "Up");
                NodeAdd(TreeViewUp.Nodes[0], "select * from [View_設備管理] where [設備編號] in (select [上游編號] from [接網] where [下游編號]=" + TreeViewUp.Nodes[0].Value + ") order by [放置地點],[設備名稱]", "Up");
            }
        }
    }
    //------------產生樹狀結構------------------------------------------------------------------------------
    protected void TreeViewDn_TreeNodePopulate(object sender, TreeNodeEventArgs e)  //當使用者按一下開啟節點時
    {
        TreeNode node = e.Node;

        if (node.ChildNodes.Count == 0)   //子節點展開過了沒？沒有，則需要產生，否則不必重複產生
        {
            string SQL = "select * from [View_設備管理] where [設備編號] in (select [下游編號] from [接網] where [上游編號]=" + node.Value + ")";
            if (node.Depth == 0) SQL = SQL + " or [設備編號] in (select [上游編號] from [接網]) and [設備編號] not in (select [下游編號] from [接網])";

            NodeAdd(node, SQL + " order by [放置地點],[設備名稱]", "Dn");
        }
    }
    protected void TreeViewUp_TreeNodePopulate(object sender, TreeNodeEventArgs e)  //當使用者按一下開啟節點時
    {
        TreeNode node = e.Node;

        if (node.ChildNodes.Count == 0)   //子節點展開過了沒？沒有，則需要產生，否則不必重複產生
        {
            NodeAdd(node, "select * from [View_設備管理] where [設備編號] in (select [上游編號] from [接網] where [下游編號]=" + node.Value + ") order by [放置地點],[設備名稱]", "Up");
        }
    }

    protected Boolean HasChild(int DevNo, string UpDn)   //判斷是否還有子節點
    {
        Boolean ChildTF = false;
        string SQL = "Select * from [接網] where [上游編號]=" + DevNo;
        if (UpDn == "Up") SQL = "Select * from [接網] where [下游編號]=" + DevNo;

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

    protected void NodeAdd(TreeNode node, string SQL, string UpDn)    //產生有子節點之目錄節點
    {
        SqlCommand sqlQuery = new SqlCommand(SQL);
        DataSet ds = RunQuery(sqlQuery);
        string PosName="";

        if (ds.Tables.Count > 0)
        {
            foreach (DataRow row in ds.Tables[0].Rows)
            {
                PosName = row["定位名稱"].ToString();
                if (row["定位方式"].ToString() == "坐標") PosName = GetValue("select [定位名稱] from [定位設定] where [定位方式]='分區' and " 
                    + row["坐標X"].ToString() + " between [坐標X1] and [坐標X2] and " + row["坐標Y"].ToString() + " between [坐標Y1] and [坐標Y2]");
                TreeNode NewNode = new TreeNode(PosName + "：" + row["設備名稱"].ToString(), row[0].ToString());  //TreeNode(顯示，值)

                if (HasChild(int.Parse(NewNode.Value), UpDn))
                {
                    NewNode.PopulateOnDemand = true;
                    NewNode.SelectAction = TreeNodeSelectAction.SelectExpand;   //觸發TreeNodeExpanded
                    NewNode.Expanded = false;   //不展開節點                    
                }
                else
                {
                    NewNode.PopulateOnDemand = false; //還有子節點，使用者按一下會觸發TreeNodePopulate事件
                }

                NewNode.ToolTip = row["設備編號"].ToString() + ". " + row["設備用途"].ToString() + " " + row["放置地點"].ToString() + (" [" + FormatChilds(row["接網迴路"].ToString()) + "]").Replace("[0]", "[無上游]");
                node.ChildNodes.Add(NewNode);
            }
        }

        sqlQuery.Cancel(); ds.Dispose();
    }

    protected string FormatChilds(string PkNos)    //格式化下游字串
    {
        string cfg = PkNos.Replace(",0,", ",").Replace(",,", ",");
        if (cfg == "" | cfg == ",") cfg = "0";
        if (cfg.Substring(cfg.Length - 1) == ",") cfg = cfg.Substring(0, cfg.Length - 1);
        if (cfg.Substring(0, 1) == ",") cfg = cfg.Substring(1);

        return (cfg);
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

    protected void ReadDevice(string DevNo, TreeNode node, string UpDn)
    {
        if (node != null) node.Selected = true;

        TreeViewUp.Nodes[0].ChildNodes.Clear();
        TreeViewUp.Nodes[0].Text = "<font color='blue' size='large'><b>接網迴路(往上)</b></font>";
        TreeViewUp.Nodes[0].Value = "-3"; //設備編號若為負值，表為系統專用，(總電源)：-2，(總網源)：-3，-1為前二者接電接網虛擬上游，無資料對應；0為程式填值用，為免出錯不使用

        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
        Conn.Open();
        SqlCommand cmd = new SqlCommand("Select * from [View_通用設備] where [設備編號]=" + DevNo, Conn);
        SqlDataReader dr = null;
        dr = cmd.ExecuteReader();
        if (dr.Read())
        {
            LblDevNo.Text = dr["設備編號"].ToString();
            lblDevKind.Text = dr["設備種類"].ToString();
            lblDevName.Text = dr["設備名稱"].ToString();
            lblPurpose.Text = dr["設備用途"].ToString();            
            lblPlace.Text = dr["放置地點"].ToString() + "，高度：" + dr["高度"].ToString() + "，" + dr["空間大小"].ToString() + "(U)";
            lblBrand.Text = dr["廠牌"].ToString();
            lblStyle.Text = dr["型式"].ToString();
            lblIP.Text = dr["IP"].ToString();
            lblMaintainor.Text = dr["設備維護人員"].ToString();
            lblStatus.Text = dr["設備狀態"].ToString();
            lblMemo.Text = dr["設備備註說明"].ToString().Replace("\r\n", "<br />");

            TreeViewUp.Nodes[0].Text = "<font color='blue' size='large'><b>" + dr["設備名稱"].ToString() + "(往上)</b></font>";
            TreeViewUp.Nodes[0].Value = dr["設備編號"].ToString();
            TreeViewUp.Nodes[0].ToolTip = dr["設備編號"].ToString() + ". " + dr["設備用途"].ToString() + " " + dr["放置地點"].ToString() + (" [" + FormatChilds(dr["接網迴路"].ToString()) + "]").Replace("[0]", "[無上游]");
            if (HasChild(int.Parse(TreeViewUp.Nodes[0].Value), "Up"))
            {
                TreeViewUp.Nodes[0].PopulateOnDemand = true;
                TreeViewUp.Nodes[0].SelectAction = TreeNodeSelectAction.SelectExpand;   //觸發TreeNodeExpanded
                TreeViewUp.Nodes[0].Expanded = false;   //不展開節點                    
            }
            else
            {
                TreeViewUp.Nodes[0].PopulateOnDemand = false; //還有子節點，使用者按一下會觸發TreeNodePopulate事件
            }
        }

        cmd.Cancel(); cmd.Dispose(); dr.Close(); Conn.Close(); Conn.Dispose();
    }

    protected void TreeViewDn_SelectedNodeChanged(object sender, EventArgs e)
    {
        ReadDevice(TreeViewDn.SelectedNode.Value, TreeViewDn.SelectedNode, "Dn");
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

    protected void BtnEdit_Click(object sender, EventArgs e)
    {
        OpenExecWindow("window.open('../Device/DevEdit.aspx?DevNo=" + LblDevNo.Text + "','_blank');");
    }

    protected void BtnConn_Click(object sender, EventArgs e)
    {
        OpenExecWindow("window.open('../Device/TreeEdit.aspx?DevNo=" + LblDevNo.Text + "','_blank');");
    }

    protected void BtnPlace_Click(object sender, EventArgs e)
    {
        Session["DevSQL"] = "SELECT * FROM [View_設備管理] WHERE [設備編號]=" + LblDevNo.Text;
        OpenExecWindow("window.open('../Lib/map.aspx','_blank','location=no;menubar=no;resizable=yes;scrollbars=no;status=no;toolbar=no;fullscreen=yes');");
    }

    protected void BtnTree_Click(object sender, EventArgs e)
    {
        Session["DevSQL"] = "SELECT * FROM [View_設備管理] WHERE (1<>1" + GetTreeSQL(LblDevNo.Text, "接網") + ")";
        OpenExecWindow("window.open('../Device/Device.aspx?Search=Tree" + "','_self');");
        //Response.Write(Session["DevSQL"]);
    }

    protected string GetTreeSQL(string DevNo, string tbl)
    {
        string SQL = "";

        SqlCommand sqlQuery = new SqlCommand();
        sqlQuery.CommandText = "select [下游編號] from [" + tbl + "] where [上游編號]=" + DevNo;
        DataSet ds = RunQuery(sqlQuery);

        if (ds.Tables.Count > 0)
        {
            foreach (DataRow row in ds.Tables[0].Rows)
            {
                SQL = SQL + " or [設備編號]=" + row[0].ToString() + GetTreeSQL(row[0].ToString(), tbl);
            }
        }
        sqlQuery.Cancel(); ds.Dispose();

        return (SQL);
    }

    protected void TreeViewDn_TreeNodeCollapsed(object sender, TreeNodeEventArgs e)
    {
        ReadDevice(e.Node.Value, e.Node, "Dn");
    }

    protected void TreeViewDn_TreeNodeExpanded(object sender, TreeNodeEventArgs e)
    {
        ReadDevice(e.Node.Value, e.Node, "Dn");
    }

    protected string GetValue(string SQL) //讀取某系統設定值
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
}