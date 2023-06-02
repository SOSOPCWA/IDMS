using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.Configuration;
using System.Data;
using System.Data.SqlClient;

public partial class SOS_PowerPlace : System.Web.UI.Page
{   //-------------起始頁面---------------------------------------------------------------------
    protected void Page_Load(object sender, EventArgs e)
    {
        if (!IsPostBack)   //第一次開網頁時
        {
            
        }
    }
    //------------產生樹狀結構------------------------------------------------------------------------------
    protected void TreeView1_TreeNodePopulate(object sender, TreeNodeEventArgs e)  //當使用者按一下開啟節點時
    {
        TreeNode node = e.Node;

        if (node.ChildNodes.Count == 0)   //子節點展開過了沒？沒有，則需要產生，否則不必重複產生
        {
            switch (node.Depth)
            {
                case 0: NodePlaceAdd(node); break;
                case 1: NodePointerAdd(node); break;
                case 2: NodeCircuitAdd(node); break;
                case 3: NodeDeviceAdd(node); break;
            }
        }
    }

    protected Boolean HasChild(string SQL)   //判斷是否還有子節點
    {
        Boolean ChildTF=false;       

        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
        Conn.Open();
        SqlCommand cmd = new SqlCommand("", Conn);
        cmd.CommandText = SQL;
        SqlDataReader dr = null;
        dr = cmd.ExecuteReader();

        try { if (dr.Read()) ChildTF=true; }
        catch { ChildTF=false; }

        cmd.Cancel(); cmd.Dispose(); dr.Close(); Conn.Close(); Conn.Dispose();
        return (ChildTF);
    }

    protected void NodePlaceAdd(TreeNode node)    //產生區域名稱節點
    {
        SqlCommand sqlQuery = new SqlCommand("select * from Config where Kind='門禁管制區' order by Mark");
        DataSet ds = RunQuery(sqlQuery);

        if (ds.Tables.Count > 0)
        {
            foreach (DataRow row in ds.Tables[0].Rows)
            {
                NodeAdd(node, row[1].ToString(), "Pla" + row[3].ToString(), row[1].ToString(), true);
            }
        }

        sqlQuery.Cancel(); ds.Dispose();
    }

    protected string GetPosSQL(string PosNo)
    {
        string SQL = "select [定位編號] from [定位設定] where [定位方式]='坐標'";

        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
        Conn.Open();
        SqlCommand cmd = new SqlCommand("select * from [定位設定] where [定位編號]=" + PosNo, Conn);
        SqlDataReader dr = null;
        dr = cmd.ExecuteReader();
        if (dr.Read()) SQL = SQL + " and [坐標X]>=" + dr["坐標X1"].ToString() + " and [坐標X]<=" + dr["坐標X2"].ToString()
            + " and [坐標Y]>=" + dr["坐標Y1"].ToString() + " and [坐標Y]<=" + dr["坐標Y2"].ToString();

        cmd.Cancel(); cmd.Dispose(); dr.Close(); Conn.Close(); Conn.Dispose();
        return (SQL);
    }

    protected void NodePointerAdd(TreeNode node)    //產生定位名稱節點
    {
        SqlCommand sqlQuery = new SqlCommand("select [定位編號],[定位名稱] from [定位設定] where [區域名稱]='" + node.Text + "' and [定位方式]<>'坐標' order by [定位方式],[定位名稱]");
        DataSet ds = RunQuery(sqlQuery);

        if (ds.Tables.Count > 0)
        {
            foreach (DataRow row in ds.Tables[0].Rows)
            {
                if (HasChild("Select * from [接電] where [下游編號] in (select [設備編號] from [實體設備] where [定位編號]=" + row[0].ToString() + " or [定位編號] in (" + GetPosSQL(row[0].ToString()) + "))"))
                {
                    NodeAdd(node, row[1].ToString(), "Poi" + row[0].ToString(), row[0].ToString() + ". " + row[1].ToString(), true);
                }
            }
        }

        sqlQuery.Cancel(); ds.Dispose();
    }

    protected void NodeCircuitAdd(TreeNode node)    //產生接電迴路節點
    {
        SqlCommand sqlQuery = new SqlCommand("");        
        sqlQuery.CommandText = "select * from [View_設備管理] where [設備編號] in (select [上游編號] from [接電] where [下游編號] in (select [設備編號] from [實體設備] where [定位編號]=" + node.Value.Substring(3) + " or [定位編號] in (" + GetPosSQL(node.Value.Substring(3)) + "))) order by [設備名稱]";
        DataSet ds = RunQuery(sqlQuery);

        if (ds.Tables.Count > 0)
        {
            foreach (DataRow row in ds.Tables[0].Rows)
            {
                NodeAdd(node, row["設備名稱"].ToString(), "Cir" + row[0].ToString(), row["設備編號"].ToString() + ". " + row["設備用途"].ToString() + " " + row["放置地點"].ToString() + (" [" + FormatChilds(row["接電迴路"].ToString()) + "]").Replace("[0]", "[無上游]"), true);
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

    protected void NodeDeviceAdd(TreeNode node)    //產生設備節點
    {
        string SQL = "select * from [View_設備管理] where ([定位編號]=" + node.Parent.Value.Substring(3) + " or [定位編號] in (" + GetPosSQL(node.Parent.Value.Substring(3)) + ")) and [設備編號] in (select [下游編號] from [接電] where [上游編號]=" + node.Value.Substring(3) + ") order by [設備名稱]";
        SqlCommand sqlQuery = new SqlCommand(SQL);
        DataSet ds = RunQuery(sqlQuery);

        if (ds.Tables.Count > 0)
        {
            foreach (DataRow row in ds.Tables[0].Rows)
            {
                NodeAdd(node, row["設備名稱"].ToString(), "Dev" + row[0].ToString(), row["設備編號"].ToString() + ". " + row["設備用途"].ToString() + " " + row["放置地點"].ToString() + (" [" + FormatChilds(row["接電迴路"].ToString()) + "]").Replace("[0]", "[無上游]"), false);
            }
        }

        sqlQuery.Cancel(); ds.Dispose();
    }

    protected void NodeAdd(TreeNode node, string txt,string val,string tip,Boolean ChildYN)    //產生有子節點之目錄節點
    {
        TreeNode NewNode = new TreeNode(txt,val);  //TreeNode(顯示，值)
        NewNode.PopulateOnDemand = ChildYN;
        if (ChildYN)    //樹葉不設定
        {
            NewNode.SelectAction = TreeNodeSelectAction.Expand;   //觸發TreeNodeExpanded
            NewNode.Expanded = false;   //不展開節點                    
        }
        NewNode.ToolTip = tip;
        node.ChildNodes.Add(NewNode);
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

    protected void ReadDevice(TreeNode node)
    {
        node.Selected = true;
        
        if (node.Depth == 3 | node.Depth == 4)
        {
            SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
            Conn.Open();
            SqlCommand cmd = new SqlCommand("Select * from [實體設備],[定位設定] where [實體設備].[定位編號]=[定位設定].[定位編號] and [設備編號]=" + node.Value.Substring(3), Conn);
            SqlDataReader dr = null;
            dr = cmd.ExecuteReader();
            if (dr.Read())
            {
                LblDevNo.Text = dr["設備編號"].ToString();
                lblDevKind.Text = dr["設備種類"].ToString();
                lblDevName.Text = dr["設備名稱"].ToString();
                lblPurpose.Text = dr["設備用途"].ToString();
                lblPlace.Text = dr["區域名稱"].ToString() + dr["定位名稱"].ToString() + "，高度：" + dr["高度"].ToString() + "，" + dr["空間大小"].ToString() + "(U)";
                lblVoltage.Text = dr["用電電壓"].ToString();
                lblCurrent.Text = dr["額定電流"].ToString();
                lblMaintainor.Text = dr["維護人員"].ToString();
                lblStatus.Text = dr["設備狀態"].ToString();
                lblMemo.Text = dr["備註說明"].ToString().Replace("\r\n", "<br />");
            }  

            cmd.Cancel(); cmd.Dispose(); dr.Close(); Conn.Close(); Conn.Dispose();
        }
        else
        {
            LblDevNo.Text = "";
            lblDevKind.Text = "";
            lblDevName.Text = "";
            lblPurpose.Text = "";
            lblPlace.Text = "";
            lblVoltage.Text = "";
            lblCurrent.Text = "";
            lblMaintainor.Text = "";
            lblStatus.Text = "";
            lblMemo.Text = "";
        }
    }

    protected void TreeView1_SelectedNodeChanged(object sender, EventArgs e)
    {
        ReadDevice(TreeView1.SelectedNode);
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
        if (LblDevNo.Text !="") OpenExecWindow("window.open('../Device/DevEdit.aspx?DevNo=" + LblDevNo.Text + "','_blank');");
    }

    protected void BtnConn_Click(object sender, EventArgs e)
    {
        OpenExecWindow("window.open('../Device/TreeEdit.aspx?DevNo=" + LblDevNo.Text + "','_blank');");
    }

    protected void BtnPlace_Click(object sender, EventArgs e)
    {
        if (LblDevNo.Text != "")
        {
            Session["DevSQL"] = "SELECT * FROM [View_設備管理] WHERE [設備編號]=" + LblDevNo.Text;
            OpenExecWindow("window.open('../Lib/map.aspx','_blank','location=no;menubar=no;resizable=yes;scrollbars=no;status=no;toolbar=no;fullscreen=yes');");
        }
    }

    protected void BtnTree_Click(object sender, EventArgs e)
    {
        if (LblDevNo.Text != "")
        {
            Session["DevSQL"] = "SELECT * FROM [View_設備管理] WHERE (1<>1" + GetPowerTreeSQL(int.Parse(LblDevNo.Text)) + ")";
            OpenExecWindow("window.open('../Device/Device.aspx?Search=PowerTree" + "','_self');");
        }
    }

    protected string GetPowerTreeSQL(int DevNo)
    {
        string SQL = "";

        SqlCommand sqlQuery = new SqlCommand();
        sqlQuery.CommandText = "select [下游編號] from [接電] where [上游編號]=" + DevNo;
        DataSet ds = RunQuery(sqlQuery);

        if (ds.Tables.Count > 0)
        {
            foreach (DataRow row in ds.Tables[0].Rows)
            {
                SQL = SQL + " or [設備編號]=" + row[0].ToString() + GetPowerTreeSQL(int.Parse(row[0].ToString()));
            }
        }
        sqlQuery.Cancel(); ds.Dispose();

        return (SQL);
    }

    protected void TreeView1_TreeNodeCollapsed(object sender, TreeNodeEventArgs e)
    {
        ReadDevice(e.Node);
    }

    protected void TreeView1_TreeNodeExpanded(object sender, TreeNodeEventArgs e)
    {
        ReadDevice(e.Node);
    }   

    protected string GetConfig(string SQL) //讀取某系統設定值
    {
        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
        Conn.Open();
        SqlCommand cmd = new SqlCommand(SQL, Conn);
        SqlDataReader dr = null;
        dr = cmd.ExecuteReader();
        string cfg = ""; if (dr.Read()) cfg=dr[0].ToString();
        cmd.Cancel(); cmd.Dispose(); dr.Close(); Conn.Close(); Conn.Dispose();

        return (cfg);
    }
}