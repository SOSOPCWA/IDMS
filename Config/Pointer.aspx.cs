using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.Configuration;
using System.Data;
using System.Data.SqlClient;

public partial class Config_Pointer : System.Web.UI.Page
{   //-------------起始頁面---------------------------------------------------------------------
    protected void Page_Load(object sender, EventArgs e)
    {
        Session.Timeout = 720;
        if (!Page.IsPostBack)   //網頁非回傳(第一次)執行，用於初始化
        {
            for (int i = 0; i <= 50; i++)   //產生機櫃容量之選項
            {
                ListItem ItemX = new ListItem();
                ItemX.Text = i.ToString() + " (U)";
                ItemX.Value = i.ToString();
                SelSpace.Items.Add(ItemX);
            }

            SelPlace.DataSourceID = "SqlDataSource1";   //繫結區域名稱
            SelPointer.DataSourceID = "SqlDataSource2";   //繫結定位名稱
            //產生坐標之選項
            for (int i = -2; i <= 99; i++) SelX.Items.Add(i.ToString());
            for (int i = -2; i <= 31; i++) SelY.Items.Add(i.ToString());
            for (int i = -1; i <= 99; i++) SelX1.Items.Add(i.ToString());
            for (int i = -1; i <= 31; i++) SelY1.Items.Add(i.ToString());
            for (int i = -1; i <= 99; i++) SelX2.Items.Add(i.ToString());
            for (int i = -1; i <= 31; i++) SelY2.Items.Add(i.ToString());
        }

        ReadHelp(); //讀取欄位說明
    }

    //僅此事件在DataSource Binding之後，而Page_Load或其它事件，均在那之前
    protected void Page_PreRenderComplete(object sender, EventArgs e)   //放在此事件，改變選項才能自動設值，放Select_Change無效
    {
        TextPlace.Text = SelPlace.Text;
        TextPointer.Text = SelPointer.Text;
    }

    protected void ReadHelp() //讀取欄位說明
    {
        SqlCommand sqlQuery = new SqlCommand();
        sqlQuery.CommandText = "select Item,Memo from Config where Kind='定位設定' order by Mark";
        DataSet ds = RunQuery(sqlQuery);

        if (ds.Tables.Count > 0)
        {
            foreach (DataRow row in ds.Tables[0].Rows)
            {
                switch (row[0].ToString())
                {
                    case "定位編號": lblPointerNo.Text = row[1].ToString(); break;
                    case "定位方式": lblKind.Text = row[1].ToString(); break;
                    case "放置地點": lblPlace.Text = row[1].ToString(); break;
                    case "機櫃容量": lblSpace.Text = row[1].ToString(); break;
                    case "顯示坐標": lblXY.Text = row[1].ToString(); break;
                    case "範圍": lblRange.Text = row[1].ToString(); break;
                }
            }
        }

        sqlQuery.Cancel(); ds.Dispose();
    }
    //------------產生樹狀結構------------------------------------------------------------------------------
    protected void TreeView1_TreeNodePopulate(object sender, TreeNodeEventArgs e)   //當使用者按一下開啟節點時
    {
        if (e.Node.ChildNodes.Count == 0)   //子節點展開過了沒？沒有，則需要產生，否則不必重複產生
        {
            switch (e.Node.Depth)   //使用者所按節點深度
            {
                case 0:
                    RootNodeAdd(e.Node);    //若在根結點，則產生第一子節點(區域名稱)
                    break;
                case 1:
                    ChildNodeAdd(e.Node);   //若在第一子結點，則產生第二子節點(定位名稱)
                    break;
            }
        }
    }

    protected void RootNodeAdd(TreeNode node)    //在根結點產生第一子節點(區域名稱)
    {
        SqlCommand sqlQuery = new SqlCommand("select distinct [區域名稱] from [定位設定] order by [區域名稱]");
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

    protected void ChildNodeAdd(TreeNode node)  //在第一子結點產生第二子節點(定位名稱)
    {
        SqlCommand sqlQuery = new SqlCommand();
        sqlQuery.CommandText = "select [定位編號],[定位名稱] from [定位設定] where [區域名稱]=@PlaceX and [定位方式]<>'坐標' order by [定位名稱]";
        sqlQuery.Parameters.Add("@PlaceX", SqlDbType.NVarChar).Value = node.Value;
        DataSet ds = RunQuery(sqlQuery);

        if (ds.Tables.Count > 0)
        {
            foreach (DataRow row in ds.Tables[0].Rows)
            {
                if (!IsExistNode(node, row[0].ToString()))
                {
                    TreeNode NewNode = new TreeNode(row[1].ToString(), row[0].ToString());
                    NewNode.PopulateOnDemand = false;   //已無子節點，不再展開
                    node.ChildNodes.Add(NewNode);
                }
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
        dbAdapter.Fill(QueryDataSet);
        dbAdapter.Dispose(); Conn.Close(); Conn.Dispose();
        return (QueryDataSet);   
    }
    //------------選取節點而顯示該節點設定值------------------------------------------------------------------------------
    protected void TreeView1_SelectedNodeChanged(object sender, EventArgs e)
    {
        try
        {
            NodeSelect(TreeView1.SelectedValue.ToString()); //選取節點以顯示設定值
        }
        catch (Exception ex)
        {
            ClearData();
        }
    }

    protected void NodeSelect(string PointerNo) //選取節點以顯示設定值
    {
        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
        Conn.Open();
        SqlCommand cmd = new SqlCommand("select * from [定位設定] where [定位編號]=" + PointerNo, Conn);
        SqlDataReader dr = null;
        dr = cmd.ExecuteReader();
        if (dr.Read())
        {
            TextPointerNo.Text = dr["定位編號"].ToString();

            SelKind.ClearSelection(); for (int i = 0; i < SelKind.Items.Count; i++) if (SelKind.Items[i].Value == dr["定位方式"].ToString()) SelKind.Items[i].Selected = true;

            SelPlace.ClearSelection(); for (int i = 0; i < SelPlace.Items.Count; i++) if (SelPlace.Items[i].Value == dr["區域名稱"].ToString()) SelPlace.Items[i].Selected = true;
            TextPlace.Text = dr["區域名稱"].ToString();

            SelPointer.DataBind(); for (int i = 0; i < SelPointer.Items.Count; i++) if (SelPointer.Items[i].Value == dr["定位名稱"].ToString()) SelPointer.Items[i].Selected = true;
            TextPointer.Text = dr["定位名稱"].ToString();

            SelSpace.ClearSelection(); for (int i = 0; i < SelSpace.Items.Count; i++) if (SelSpace.Items[i].Value == dr["機櫃容量"].ToString()) SelSpace.Items[i].Selected = true;

            SelX.ClearSelection(); for (int i = 0; i < SelX.Items.Count; i++) if (SelX.Items[i].Value == dr["坐標X"].ToString()) SelX.Items[i].Selected = true;

            SelY.ClearSelection(); for (int i = 0; i < SelY.Items.Count; i++) if (SelY.Items[i].Value == dr["坐標Y"].ToString()) SelY.Items[i].Selected = true;

            SelX1.ClearSelection(); for (int i = 0; i < SelX1.Items.Count; i++) if (SelX1.Items[i].Value == dr["坐標X1"].ToString()) SelX1.Items[i].Selected = true;

            SelY1.ClearSelection(); for (int i = 0; i < SelY1.Items.Count; i++) if (SelY1.Items[i].Value == dr["坐標Y1"].ToString()) SelY1.Items[i].Selected = true;

            SelX2.ClearSelection(); for (int i = 0; i < SelX2.Items.Count; i++) if (SelX2.Items[i].Value == dr["坐標X2"].ToString()) SelX2.Items[i].Selected = true;

            SelY2.ClearSelection(); for (int i = 0; i < SelY2.Items.Count; i++) if (SelY2.Items[i].Value == dr["坐標Y2"].ToString()) SelY2.Items[i].Selected = true;
        }
        cmd.Cancel(); cmd.Dispose(); dr.Close(); Conn.Close(); Conn.Dispose();
    }
    //----------區域名稱之操作--------------------------------------------------------------------------------
    protected void ClearData()  //改選節點或區域名稱時，清除與位置有關的顯示值
    {
        SelSpace.ClearSelection();        
        SelX.ClearSelection();
        SelY.ClearSelection();
        SelX1.ClearSelection();
        SelY1.ClearSelection();
        SelX2.ClearSelection();
        SelY2.ClearSelection();
    }

    protected void SelPlace_SelectedIndexChanged(object sender, EventArgs e)    //改選區域名稱
    {
        ClearData();
        ChangePlace();
    }

    protected void ChangePlace()    //改選區域名稱後顯示定位名稱
    {
        //TextPlace.Text = SelPlace.Text;  //PreRenderComplete有，此處可省
        SelPointer.DataSourceID = "SqlDataSource2";
        SelPointer.DataBind();
    }

    protected void SelPointer_SelectedIndexChanged(object sender, EventArgs e)
    {
        //TextPointer.Text = SelPointer.SelectedValue;  //PreRenderComplete有，此處可省
    }
    //-------------異動執行-----------------------------------------------------------------------------
    protected void InsLifeLog(string SQL,string PkNo) //寫入生命履歷
    {
        string LifeNo = GetPKNo("履歷編號", "生命履歷").ToString(); //履歷編號
        string TblName = "定位設定";    //表格名稱
        string Mt = "SSM小組";    //維護人員
        string UN = Session["UserName"].ToString();   //登入的UserName：異動人員
        string LiftDT = DateTime.Now.ToString("yyyy/MM/dd HH:mm");  //異動日期
        string LifeIP = Request.ServerVariables["REMOTE_ADDR"].ToString();

        ExecDbSQL("Insert into [生命履歷] values(" + LifeNo + ",'" + TblName + "'," + PkNo + ",'" + SQL.Replace("'", "''") + "','" + Mt + "','" + Mt + "','" + Mt + "','" + UN + "','" + LiftDT + "','" + LifeIP + "')");
    }

    protected int GetPKNo(string PKfield, string PKtbl) //取得主鍵編號
    {
        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
        Conn.Open();
        SqlCommand cmd = new SqlCommand("select max(" + PKfield + ") from " + PKtbl, Conn);
        SqlDataReader dr = null;
        dr = cmd.ExecuteReader();
        int PkNo = 1; if (dr.Read()) PkNo=int.Parse(dr[0].ToString()) + 1;
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
        string UnitName = Session["UnitName"].ToString();   //登入的UnitName

        if (UserID != "operator" & UnitName == "電腦操作課") return (true);
        else return (false);
    }
    
    protected void BtnNew_Click(object sender, EventArgs e) //新增定位
    {
        Literal Msg = new Literal();

        if (!RightCheck())
        {
            Msg.Text = "<script>alert('您沒有權限異動資料，若有需求，請洽機房！');</script>";
            Page.Controls.Add(Msg);
        }
        else
        {
            int AutoPointerNo = GetSamePointerNo(TextPlace.Text, TextPointer.Text);
            if (AutoPointerNo > 0)  //是否已有同一定位名稱，若是則alert
            {
                Msg.Text = "<script>alert('該筆資料已存在，無法再新增！');</script>";
                Page.Controls.Add(Msg);
            }
            else //若否，則Ins DB & Tree
            {
                string PkNo = GetPKNo("[定位編號]", "[定位設定]").ToString();
                string SQL = "insert into [定位設定] values(" + PkNo + ",'"
                    + SelKind.SelectedValue + "','" + TextPlace.Text + "','"
                    + TextPointer.Text + "'," + SelSpace.SelectedValue + ","
                    + SelX.SelectedValue + "," + SelY.SelectedValue + ","
                    + SelX1.SelectedValue + "," + SelY1.SelectedValue + ","
                    + SelX2.SelectedValue + "," + SelY2.SelectedValue + ")";
                
                SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
                Conn.Open();
                SqlCommand cmd = new SqlCommand(SQL, Conn);

                cmd.ExecuteNonQuery();
                cmd.Cancel(); cmd.Dispose(); Conn.Close(); Conn.Dispose();

                AutoPointerNo = GetSamePointerNo(TextPlace.Text, TextPointer.Text);
                TextPointerNo.Text = AutoPointerNo.ToString();

                Node_Add(AutoPointerNo.ToString(), TextPlace.Text, TextPointer.Text);

                NodeSelect(AutoPointerNo.ToString()); //選取節點以顯示設定值，必須DB異動之後才能執行

                InsLifeLog(SQL,PkNo);
            }
        }
    }

    protected int GetSamePointerNo(string Place, string Pointer) //取得同一定位名稱之定位編號
    {
        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
        Conn.Open();
        SqlCommand cmd = new SqlCommand("select * from [定位設定] where [區域名稱]='" + Place + "' and [定位名稱]='" + Pointer + "'", Conn);
        SqlDataReader dr = null;
        dr = cmd.ExecuteReader();
        int Pno = 0; if (dr.Read()) Pno=int.Parse(dr[0].ToString());
        cmd.Cancel(); cmd.Dispose(); dr.Close(); Conn.Close(); Conn.Dispose();

        return (Pno);
    }    

    protected void Node_Add(string PointerNo, string Place, string Pointer)  //新增節點
    {
        TreeNode PlaceNode = new TreeNode();
        //---------------------------------------是否有區域名稱，若無要先新增
        Boolean flag = false;
        foreach (TreeNode node in TreeView1.Nodes[0].ChildNodes)  //判斷是否有區域名稱
        {
            if (node.Value == Place)
            {
                flag = true;
                PlaceNode = node;
                break;
            }
        }

        if (!flag)  //若無區域名稱要先新增
        {
            PlaceNode.Text = Place;
            PlaceNode.Value = Place;
            PlaceNode.PopulateOnDemand = true;    //還有子節點，使用者按一下會觸發TreeNodePopulate事件
            PlaceNode.SelectAction = TreeNodeSelectAction.Expand;   //觸發TreeNodeExpanded
            TreeView1.Nodes[0].ChildNodes.Add(PlaceNode);  //新增位置節點

            SelPlace.Items.Add(Place);  //新增區域名稱的下拉式選單
        }
        //---------------------------------------新增子節點(定位名稱)
        TreeNode PointerNode = new TreeNode(Pointer, PointerNo);
        PointerNode.PopulateOnDemand = false;   //已無子節點，不再展開
        PointerNode.SelectAction = TreeNodeSelectAction.Select; //觸發SelectedNodeChanged
        PlaceNode.Expand(); 

        SelPointer.Items.Add(Pointer);  //新增定位名稱的下拉式選單

        if (!IsExistNode(PlaceNode, PointerNo)) PlaceNode.ChildNodes.Add(PointerNode);  //新增子節點 (第一次不排除的話，會多add一個node)

        foreach (TreeNode NewNode in PlaceNode.ChildNodes)//選取子節點
        {
            if (NewNode.Value == PointerNo)
            {
                NewNode.Select();                
                break;
            }
        }        
    }

    protected Boolean IsExistNode(TreeNode PNode,string CNodeVal)  //是否存在某節點
    {
        Boolean flag = false;
        
        for (int i = 0; i < PNode.ChildNodes.Count; i++)
        {
            if (PNode.ChildNodes[i].Value == CNodeVal) flag=true;
        }
        return (flag);
    }

    protected void Node_Del()  //刪除節點
    {
        if (TreeView1.SelectedNode.Parent.ChildNodes.Count > 1) //若尚有其它子節點，則移除子節點
        {
            for (int i = 0; i < SelPointer.Items.Count; i++)  //移除區域名稱的下拉式選單
            {
                if (SelPointer.Items[i].Value == TreeView1.SelectedNode.Text) { SelPointer.Items.RemoveAt(i); }
            }

            TreeView1.SelectedNode.Parent.Collapse();   //摺疊節點
            TreeView1.SelectedNode.Parent.ChildNodes.Remove(TreeView1.SelectedNode);            
        }
        else //若父節點已其它無子節點，則移除父節點
        {
            for (int i = 0; i < SelPlace.Items.Count; i++)  //移除定位名稱的下拉式選單
            {
                if (SelPlace.Items[i].Value == TreeView1.SelectedNode.Parent.Text) { SelPlace.Items.RemoveAt(i); }
            }
            
            TreeView1.SelectedNode.Parent.Parent.ChildNodes.Remove(TreeView1.SelectedNode.Parent);           
        }
    }

    protected void BtnEdit_Click(object sender, EventArgs e)    //修改定位
    {
        Literal Msg = new Literal();

        if (!RightCheck())
        {
            Msg.Text = "<script>alert('您沒有權限異動資料，若有需求，請洽機房！');</script>";
            Page.Controls.Add(Msg);
        }
        else
        {
            if (TextPointerNo.Text != "")
            {
                string SQL = "update [定位設定] set [定位方式]='" + SelKind.SelectedValue + "',[區域名稱]='" + TextPlace.Text
                    + "',[定位名稱]='" + TextPointer.Text + "',[機櫃容量]=" + SelSpace.SelectedValue
                    + ",[坐標X]=" + SelX.SelectedValue + ",[坐標Y]=" + SelY.SelectedValue
                    + ",[坐標X1]=" + SelX1.SelectedValue + ",[坐標Y1]=" + SelY1.SelectedValue
                    + ",[坐標X2]=" + SelX2.SelectedValue + ",[坐標Y2]=" + SelY2.SelectedValue
                    + " where [定位編號]=" + TextPointerNo.Text;
                
                SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
                Conn.Open();
                SqlCommand cmd = new SqlCommand(SQL, Conn);
                cmd.ExecuteNonQuery();
                cmd.Cancel(); cmd.Dispose(); Conn.Close(); Conn.Dispose();

                Node_Del();  //先刪節點
                Node_Add(TextPointerNo.Text, TextPlace.Text, TextPointer.Text);  //再新增節點
                NodeSelect(TextPointerNo.Text); //選取節點以顯示設定值，必須DB異動之後才能執行

                InsLifeLog(SQL,TextPointerNo.Text);
            }
            else
            {
                Msg.Text = "<script>alert('您尚未點選欲修改之資料！');</script>";
                Page.Controls.Add(Msg);
            }
        }
    }

    protected void BtnDel_Click(object sender, EventArgs e) //刪除定位
    {
        Literal Msg = new Literal();

        if (!RightCheck())
        {
            Msg.Text = "<script>alert('您沒有權限異動資料，若有需求，請洽機房！');</script>";
            Page.Controls.Add(Msg);
        }
        else
        {
            if (TextPointerNo.Text != "")
            {
                string SQL = "delete [定位設定] where [定位編號]=" + TextPointerNo.Text;
                
                SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
                Conn.Open();
                SqlCommand cmd = new SqlCommand(SQL, Conn);
                cmd.ExecuteNonQuery();
                cmd.Cancel(); cmd.Dispose(); Conn.Close(); Conn.Dispose();

                InsLifeLog(SQL + ",[區域名稱]='" + TextPlace.Text + "',[定位名稱]='" + TextPointer.Text, TextPointerNo.Text);

                ClearData();
                TextPointerNo.Text = "";
                Node_Del();  //刪除節點
            }
            else
            {
                Msg.Text = "<script>alert('您尚未點選欲刪除之資料！');</script>";
                Page.Controls.Add(Msg);
            }
        }
    }
    protected void BtnPlace_Click(object sender, EventArgs e)
    {
        Literal Msg = new Literal();

        if (TextPno.Text == "") Msg.Text = "請先輸入所要查詢的定位編號！";
        else
        {
            SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
            Conn.Open();
            SqlCommand cmd = new SqlCommand("select * from [定位設定] where [定位編號]=" + TextPno.Text, Conn);
            SqlDataReader dr = null;

            dr = cmd.ExecuteReader();
            if (dr.Read()) Msg.Text = dr["區域名稱"].ToString() + dr["定位名稱"].ToString();
            else Msg.Text = "查無定位編號為" + TextPno.Text + "的資料！";

            cmd.Cancel(); cmd.Dispose(); dr.Close(); Conn.Close(); Conn.Dispose();
        }
        Msg.Text = "<script>alert('" + Msg.Text + "');</script>";
        Page.Controls.Add(Msg);        
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

    protected void BtnMap_Click(object sender, EventArgs e) //設備分佈圖
    {
        //使用Session["DevSQL"]傳值
        string PointerNo = TextPointerNo.Text;
        Session["DevSQL"] = "SELECT * FROM [View_設備管理] WHERE 1=1";
        if (PointerNo != "") Session["DevSQL"] = Session["DevSQL"] + " and [定位編號]=" + PointerNo + " or [定位編號] in (" + GetPosSQL(PointerNo) + ")";
        OpenExecWindow("window.open('../Lib/map.aspx','_blank','location=no;menubar=no;resizable=yes;scrollbars=no;status=no;toolbar=no;fullscreen=yes');");
    }
    protected string GetPosSQL(string PosNo)
    {
        string SQL = "select [定位編號] from [定位設定]";

        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
        Conn.Open();
        SqlCommand cmd = new SqlCommand("select * from [定位設定] where [定位編號]=" + PosNo, Conn);
        SqlDataReader dr = null;
        dr = cmd.ExecuteReader();
        if (dr.Read()) SQL = SQL + " where [坐標X]>=" + dr["坐標X1"].ToString() + " and [坐標X]<=" + dr["坐標X2"].ToString()
            + " and [坐標Y]>=" + dr["坐標Y1"].ToString() + " and [坐標Y]<=" + dr["坐標Y2"].ToString();

        cmd.Cancel(); cmd.Dispose(); dr.Close(); Conn.Close(); Conn.Dispose();
        return (SQL);
    }
}