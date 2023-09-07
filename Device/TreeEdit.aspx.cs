using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.Configuration;
using System.Data;
using System.Data.SqlClient;

public partial class Device_TreeEdit : System.Web.UI.Page
{   
    protected void Page_PreRenderComplete(object sender, EventArgs e)   //放Page_Load無法顯示
    {   
        if (!IsPostBack)
        {
            if (Request["DevNo"] != null) TextDevNo.Text = HttpUtility.HtmlEncode(Request["DevNo"]);

            if (TextDevNo.Text == "" | TextDevNo.Text == "0") AddMsg("<script>alert('您尚未指定設備！ 請先完成新增設備後再設定設備迴路資訊！');window.close();</script>");
            else
            {
                ReadDevice(TextDevNo.Text);   //讀取實體設備資訊        
                ReadConn(TextDevNo.Text, "接電", "Up", ListUpPower);  //讀取上游接電迴路選單
                ReadConn(TextDevNo.Text, "接電", "Dn", ListDnPower); //讀取下游接電迴路選單
                ReadConn(TextDevNo.Text, "接網", "Up", ListUpNet); //讀取上游接網迴路選單 
                ReadConn(TextDevNo.Text, "接網", "Dn", ListDnNet); //讀取下游接網迴路選單 
                ReadConn(TextDevNo.Text, "空調", "Up", ListUpCold);
                ReadConn(TextDevNo.Text, "空調", "Dn", ListDnCold);
            }

            ReadHelp();
            SelPlace.Items.Add("空調群組");
            Place_Selected(ListSource, SelPlace, SelPointer, SqlDataSourcePointer, SqlDataSourceSource);
        }
    }

    protected void ReadDevice(string DevNo)    //讀取實體設備
    {
        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
        Conn.Open();
        SqlCommand cmd;
        SqlDataReader dr = null;
        if(DevNo.Length<5){ //一般設備
            cmd = new SqlCommand("SELECT * FROM [View_設備管理] where [設備編號]=@DevNo", Conn);
            cmd.Parameters.Add("DevNo", DevNo);            
            dr = cmd.ExecuteReader();
            if (dr.Read())
            {
                lblDevName.Text = HttpUtility.HtmlEncode(dr["設備名稱"].ToString());
                lblDevice.Text =  HttpUtility.HtmlEncode(dr["設備用途"].ToString() + " " + dr["放置地點"].ToString() + " " + dr["高度"].ToString() + "(U)");
                lblMt.Text = HttpUtility.HtmlEncode(dr["維護人員"].ToString());
                lblTree.Text = HttpUtility.HtmlEncode(GetPath(dr["設備編號"].ToString(), "接電", "") + "＊ AAD" + GetPath(dr["設備編號"].ToString(), "接網", "") + "＊ AAD" + GetPath(dr["設備編號"].ToString(), "空調", "") + "＊").Replace("AAD", "</br>");
            }        
        }
        else{ //空調設備
            cmd = new SqlCommand("SELECT * FROM [實體群組] where [群組編號]=@vno", Conn);
            cmd.Parameters.Add("vno", DevNo);
            dr = cmd.ExecuteReader();
            if (dr.Read())
            {
                lblDevName.Text = HttpUtility.HtmlEncode(dr["群組名稱"].ToString());
                lblDevice.Text =  HttpUtility.HtmlEncode(dr["群組說明"].ToString());
                lblMt.Text = HttpUtility.HtmlEncode(dr["擺放地點"].ToString());
                
            }  
        
        }
        

        cmd.Cancel(); cmd.Dispose(); dr.Close(); Conn.Close(); Conn.Dispose();
    }

    protected void ReadConn(string DevNo, string tbl, string UpDn, ListBox ListName)    //讀取上下游設備迴路選單
    {
        SqlCommand sqlQuery = new SqlCommand();
        string UpDn1 = "上", UpDn2 = "下";
        if (UpDn == "Dn")
        {
            UpDn1 = "下"; UpDn2 = "上";
        }
        if(tbl == "空調"){
            sqlQuery.CommandText = String.Format(@"SELECT [設備編號],[設備名稱] 
                FROM ( SELECT d.設備編號,d.設備名稱
                        FROM [dbo].[View_設備管理] d
                        UNION
                        Select gp.群組編號,gp.群組名稱
                        FROM [dbo].[實體群組] gp
                        ) gt,[空調]
                WHERE gt.[設備編號]=[{0}游編號]
                and [{1}游編號] = {2} ORDER BY [設備名稱]", UpDn1, UpDn2, DevNo);

        }else{            
            sqlQuery.CommandText = "SELECT [設備編號],[設備名稱] FROM [實體設備],[" + tbl + "]"
            + " WHERE [實體設備].[設備編號]=[" + tbl + "].[" + UpDn1 + "游編號]"
            + " and [" + tbl + "].[" + UpDn2 + "游編號] = " + DevNo + " ORDER BY [設備名稱]";
        }
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

    protected void ReadHelp() //讀取欄位說明
    {
        SqlCommand sqlQuery = new SqlCommand();
        sqlQuery.CommandText = "select [Item],[Memo],[Config] from [Config] where [Kind]='設備迴路' order by [Mark]";
        DataSet ds = RunQuery(sqlQuery);

        if (ds.Tables.Count > 0)
        {
            foreach (DataRow row in ds.Tables[0].Rows)
            {
                Label helpObj = (Label)form1.FindControl("help" + row[2].ToString()); //說明文字
                helpObj.Text = HttpUtility.HtmlEncode(row[1].ToString()).Replace("\r\n", "<br />");
            }
        }

        sqlQuery.Cancel(); ds.Dispose();
    }

    //新增迴路----------------------------------------------------------------
    protected Boolean BeenConned(string No, ListBox ListName) //是否有重複的設備迴路
    {
        Boolean TF = false;
        for (int i = 0; i < ListName.Items.Count; i++)
        {
            if (ListName.Items[i].Value == No) TF = true;
        }

        return (TF);
    }

    protected void LinkUpPowerAdd_Click(object sender, EventArgs e)
    {
        LinkAdd_Click(ListSource, ListUpPower);
    }
    protected void LinkDnPowerAdd_Click(object sender, EventArgs e)
    {
        LinkAdd_Click(ListSource, ListDnPower);
    }
    protected void LinkUpNetAdd_Click(object sender, EventArgs e)
    {
        LinkAdd_Click(ListSource, ListUpNet);
    }
    protected void LinkDnNetAdd_Click(object sender, EventArgs e)
    {
        LinkAdd_Click(ListSource, ListDnNet);
    }
    protected void LinkUpColdAdd_Click(object sender, EventArgs e)
    {
        LinkAdd_Click(ListSource, ListUpCold);
    }
    protected void LinkDnColdAdd_Click(object sender, EventArgs e)
    {
        LinkAdd_Click(ListSource, ListDnCold);
    }

    protected void LinkAdd_Click(ListBox ListSource, ListBox ListTarget)
    {
        string tbl = "", UpDn = "", UpDnC = "", UpNo = "", DnNo = "",TreeNos="";
        Boolean TF = false;
        
        if (ListSource.SelectedValue == "") AddMsg("<script>alert('請先選取您要新增的迴路！');</script>");
        else if (ListSource.SelectedValue == TextDevNo.Text) AddMsg("<script>alert('您不能新增自己為設備迴路！');</script>");
        else if (BeenConned(ListSource.SelectedValue, ListTarget)) AddMsg("<script>alert('該迴路已存在，您不必再新增！');</script>");
        else
        {
            if (RightCheck())
            {
                TF = true;

                if (ListTarget.ID == "ListUpPower" | ListTarget.ID == "ListDnPower")
                {
                    tbl = "接電";
                    if (ListTarget.ID == "ListUpPower")
                    {
                        UpDn = "Up"; UpDnC = "上游";
                        UpNo = ListSource.SelectedValue;
                        DnNo = TextDevNo.Text;                        
                    }
                    else
                    {
                        UpDn = "Dn"; UpDnC = "下游";
                        UpNo = TextDevNo.Text;
                        DnNo = ListSource.SelectedValue;
                    }

                    if (GetValue("select count(*) from [接電] where [下游編號]=@uno" , "@uno", UpNo) == "0")  //作為上游之迴路無電力上游！
                    {
                    //    TF = false;
                        AddMsg("<script>alert('作為上游之迴路(" + HttpUtility.HtmlEncode(ListSource.SelectedItem.Text) + ")無電力上游！若有需要，請新增至(總電源)之下。');</script>");
                    }                                        
                }
                else if(ListTarget.ID == "ListUpNet" | ListTarget.ID == "ListDnNet")
                {
                    tbl = "接網";
                    if (ListTarget.ID == "ListUpNet")
                    {
                        UpDn = "Up"; UpDnC = "上游";
                        UpNo = ListSource.SelectedValue;
                        DnNo = TextDevNo.Text;

                        if (!NetCheck(UpNo)) TF = false;
                    }
                    else
                    {
                        UpDn = "Dn"; UpDnC = "下游";
                        UpNo = TextDevNo.Text;
                        DnNo = ListSource.SelectedValue;

                        if (!NetCheck(DnNo)) TF = false;
                    }

                    if (GetValue("select count(*) from [接網] where [下游編號]=@uno", "@uno" , UpNo) == "0")
                    {
                    //    TF = false;
                        AddMsg("<script>alert('作為上游之迴路(" + HttpUtility.HtmlEncode(ListSource.SelectedItem.Text) + ")無網路上游！若有需要，請新增至(總網源)之下。');</script>");
                    }
                }
                else{
                    tbl = "空調";
                    if (ListTarget.ID == "ListUpCold")
                    {
                        UpDn = "Up"; UpDnC = "上游";
                        UpNo = ListSource.SelectedValue;
                        DnNo = TextDevNo.Text;

                       
                    }
                    else
                    {
                        UpDn = "Dn"; UpDnC = "下游";
                        UpNo = TextDevNo.Text;
                        DnNo = ListSource.SelectedValue;

                    }

                    if (GetValue("select count(*) from [空調] where [下游編號]= @uno", "@uno" , UpNo) == "0")
                    {
                    //    TF = false;
                        AddMsg("<script>alert('作為上游之迴路(" + HttpUtility.HtmlEncode(ListSource.SelectedItem.Text) + ")無空調上游！');</script>");
                    }
                }              
            }            

            if (TF) //不能把節點之上下游樹系中任一節點再新增為上下游(網路放寬)
            {
                if (tbl == "接電") TreeNos = FormatChilds(GetTreeNos(tbl, TextDevNo.Text, "Up") + "," + GetTreeNos(tbl, TextDevNo.Text, "Dn"));
                else if (UpDn == "Up") TreeNos = GetTreeNos(tbl, TextDevNo.Text, "Dn");
                else TreeNos = GetTreeNos(tbl, TextDevNo.Text, "Up");

                if (("," + TreeNos + ",").IndexOf("," + ListSource.SelectedValue + ",") >= 0)
                {
                    TF = false;
                    if (tbl == "接電") AddMsg("<script>alert('不能把節點之上下游樹系中任一節點再新增為上下游！');</script>");
                    else if (UpDn == "Up") AddMsg("<script>alert('不能把節點之下游樹系中任一節點再新增為上游！');</script>");
                    else AddMsg("<script>alert('不能把節點之上游樹系中任一節點再新增為下游！');</script>");
                }
                else
                {
                    string SQL = "Insert into [" + tbl + "] values(" + UpNo + "," + DnNo + ")";
                    Trace.Write(SQL);
                    ExecDbSQL(SQL);
                    InsLifeLog("新增 [" + lblDevName.Text + "] 的" + UpDnC + tbl + "迴路 [" + ListSource.SelectedItem.Text + "]：" + SQL);
                    AddMsg("<script>alert('已完成新增" + UpDnC + tbl + "迴路：" + ListSource.SelectedItem.Text + "');</script>");
                    ReadConn(TextDevNo.Text, tbl, UpDn, ListTarget);  //讀取設備迴路選單                
                }
            }
            else
            {
                if (tbl == "接電") AddMsg("<script>alert('接電迴路您沒有設定此動作的權限！若涉及環境小組，請向其洽詢！');</script>");
                if (tbl == "接網") AddMsg("<script>alert('接網迴路您沒有設定此動作的權限！若涉及網管小組，請向其洽詢！');</script>");
            }
        }        
    }

    protected string GetTreeNos(string tbl, string PkNo,string UpDn)   //依某節點取得所有樹系上下游編號
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

    protected string FormatChilds(string PkNos)    //格式化下游字串
    {
        string cfg = PkNos.Replace(",0,", ",").Replace(",,", ",");
        if (cfg == "" | cfg == ",") cfg = "0";
        if (cfg.Substring(cfg.Length - 1) == ",") cfg = cfg.Substring(0, cfg.Length - 1);
        if (cfg.Substring(0, 1) == ",") cfg = cfg.Substring(1);

        return (cfg);
    } 

    //刪除迴路----------------------------------------------------------------
    protected void LinkUpPowerDel_Click(object sender, EventArgs e)
    {
        LinkDel_Click(ListUpPower);
    }
    protected void LinkDnPowerDel_Click(object sender, EventArgs e)
    {
        LinkDel_Click(ListDnPower);
    }
    protected void LinkUpNetDel_Click(object sender, EventArgs e)
    {
        LinkDel_Click(ListUpNet);
    }
    protected void LinkDnNetDel_Click(object sender, EventArgs e)
    {
        LinkDel_Click(ListDnNet);
    }
    protected void LinkUpColdDel_Click(object sender, EventArgs e)
    {
        LinkDel_Click(ListUpCold);
    }
    protected void LinkDnColdDel_Click(object sender, EventArgs e)
    {
        LinkDel_Click(ListDnCold);
    }

    protected void LinkDel_Click(ListBox ListTarget)
    {
        string tbl = "", UpDn = "", UpDnC = "", UpNo = "", DnNo = "";
        Boolean TF = false;
        
        if (ListTarget.SelectedValue == "") AddMsg("<script>alert('請先選取您要刪除的迴路！');</script>");
        else if (RightCheck())
            {
                Trace.Write("Del");
                TF = true;
                switch (ListTarget.ID)
                {
                    case "ListUpPower":
                        {
                            tbl = "接電";
                            UpDn = "Up"; UpDnC = "上游";
                            UpNo = ListTarget.SelectedValue;
                            DnNo = TextDevNo.Text;

                            if (ListTarget.Items.Count <= 1 & ListDnPower.Items.Count > 0)
                            {
                                TF = false;
                                AddMsg("<script>alert('上游為單迴路且尚有下游時禁止刪除上游！');</script>");
                            }
                            break;
                        }
                    case "ListDnPower":
                        {
                            tbl = "接電";
                            UpDn = "Dn"; UpDnC = "下游";
                            UpNo = TextDevNo.Text;
                            DnNo = ListTarget.SelectedValue;

                            string Count = GetValue("select count(*) from [" + tbl + "] where [下游編號]=@dno", "@dno" , DnNo);
                            if ((Count == "0" | Count == "1") & GetValue("select [下游編號] from [" + tbl + "] where [上游編號]=@dno", "@dno" , DnNo) != "")
                            {
                                TF = false;
                                AddMsg("<script>alert('下游為單迴路且下游尚有下游時禁止刪除下游！');</script>");
                            }
                            break;
                        }
                    case "ListUpNet":
                        {
                            tbl = "接網";
                            UpDn = "Up"; UpDnC = "上游";
                            UpNo = ListTarget.SelectedValue;
                            DnNo = TextDevNo.Text;
                            if (ListTarget.Items.Count <= 1 & ListDnNet.Items.Count > 0)
                            {
                                TF = false;
                                AddMsg("<script>alert('上游為單迴路且尚有下游時禁止刪除上游！');</script>");
                            }

                            if (!NetDelCheck(UpNo)) TF = false;
                            break;
                        }
                    case "ListDnNet":
                        {
                            tbl = "接網";
                            UpDn = "Dn"; UpDnC = "下游";
                            UpNo = TextDevNo.Text;
                            DnNo = ListTarget.SelectedValue;
                            string Count = GetValue("select count(*) from [" + tbl + "] where [下游編號]=@dno", "@dno" , DnNo);
                            if ((Count == "0" | Count == "1") & GetValue("select [下游編號] from [" + tbl + "] where [上游編號]=@uno", "@uno",  DnNo) != "")
                            {
                                TF = false;
                                AddMsg("<script>alert('下游為單迴路且下游尚有下游時禁止刪除下游！');</script>");
                            }

                            if (!NetDelCheck(DnNo)) TF = false;
                            break;
                        }
                    case "ListUpCold":
                        {
                            tbl = "空調";
                            UpDn = "Up"; UpDnC = "上游";
                            DnNo = TextDevNo.Text;
                            UpNo = ListTarget.SelectedValue;
                            

                            if (!NetDelCheck(DnNo)) TF = false;
                            break;
                        }
                    case "ListDnCold":
                        {
                            tbl = "空調";
                            UpDn = "Dn"; UpDnC = "下游";
                            UpNo = TextDevNo.Text;
                            DnNo = ListTarget.SelectedValue;
                            
                            
                            if (!NetDelCheck(DnNo)) TF = false;
                            Trace.Write(TF.ToString());
                            break;
                        }
                }                
            }
        else AddMsg("<script>alert('您沒有刪除迴路的權限，若涉及網管小組接網設備，請向其洽詢！');</script>");

        if (TF)
        {
            string SQL = "delete from [" + tbl + "] where [下游編號]=@dno and [上游編號]=@uno";
            Trace.Write(SQL);
            List<SqlParameter> pars = new List<SqlParameter>();
            pars.Add(new SqlParameter("@dno", DnNo));
            pars.Add(new SqlParameter("@uno", UpNo));
            
            ExecDbSQL(SQL, pars);
            InsLifeLog("刪除 [" + HttpUtility.HtmlEncode(lblDevName.Text) + "] 的" + UpDnC + tbl + "迴路 [" + HttpUtility.HtmlEncode(ListTarget.SelectedItem.Text) + "] ： " + SQL);
            AddMsg("<script>alert('已完成刪除" + UpDnC + tbl + "迴路：" + HttpUtility.HtmlEncode(ListTarget.SelectedItem.Text) + "');</script>");
            ReadConn(TextDevNo.Text, tbl, UpDn, ListTarget);  //讀取設備迴路選單                    
        }
		else
		{
			if (tbl == "接電") AddMsg("<script>alert('接電迴路您沒有設定此動作的權限！若涉及環境小組，請向其洽詢！');</script>");
			if (tbl == "接網") AddMsg("<script>alert('接網迴路您沒有設定此動作的權限！若涉及網管小組，請向其洽詢！');</script>");
            if (tbl == "空調") AddMsg("<script>alert('空調迴路您沒有設定此動作的權限！若涉及環境小組，請向其洽詢！');</script>");
		}		
    }

    //編輯或設定迴路----------------------------------------------------------------
    protected void LinkUpPowerEdit_Click(object sender, EventArgs e)
    {
        LinkEdit_Click(ListUpPower, "DevEdit.aspx");
    }
    protected void LinkDnPowerEdit_Click(object sender, EventArgs e)
    {
        LinkEdit_Click(ListDnPower, "DevEdit.aspx");
    }
    protected void LinkUpNetEdit_Click(object sender, EventArgs e)
    {
        LinkEdit_Click(ListUpNet, "DevEdit.aspx");
    }
    protected void LinkDnNetEdit_Click(object sender, EventArgs e)
    {
        LinkEdit_Click(ListDnNet, "DevEdit.aspx");
    }
    protected void LinkUpColdEdit_Click(object sender, EventArgs e)
    {
        LinkEdit_Click(ListUpCold, "DevEdit.aspx");
    }
    protected void LinkDnColdEdit_Click(object sender, EventArgs e)
    {
        LinkEdit_Click(ListDnCold, "DevEdit.aspx");
    }
    protected void LinkSourceEdit_Click(object sender, EventArgs e)
    {
        LinkEdit_Click(ListSource, "DevEdit.aspx");
    }

    protected void LinkUpPowerConn_Click(object sender, EventArgs e)
    {
        LinkEdit_Click(ListUpPower, "TreeEdit.aspx");
    }
    protected void LinkDnPowerConn_Click(object sender, EventArgs e)
    {
        LinkEdit_Click(ListDnPower, "TreeEdit.aspx");
    }
    protected void LinkUpNetConn_Click(object sender, EventArgs e)
    {
        LinkEdit_Click(ListUpNet, "TreeEdit.aspx");
    }
    protected void LinkDnNetConn_Click(object sender, EventArgs e)
    {
        LinkEdit_Click(ListDnNet, "TreeEdit.aspx");
    }
    protected void LinkUpColdConn_Click(object sender, EventArgs e)
    {
        LinkEdit_Click(ListUpCold, "TreeEdit.aspx");
    }
    protected void LinkDnColdConn_Click(object sender, EventArgs e)
    {
        LinkEdit_Click(ListDnCold, "TreeEdit.aspx");
    }
    protected void LinkSourceConn_Click(object sender, EventArgs e)
    {
        LinkEdit_Click(ListSource, "TreeEdit.aspx");
    }
    

    protected void LinkEdit_Click(ListBox ListName,string code)
    {
        if (ListName.SelectedValue == "") AddMsg("<script>alert('請先選取您要編輯或設定的迴路！');</script>");
        else OpenExecWindow("window.open('" + code + "?DevNo=" + HttpUtility.HtmlEncode(ListName.SelectedValue) + "','_self');");
    }

    //展開合併設備迴路----------------------------------------------------------------
    protected void LinkSourceUp_Click(object sender, EventArgs e)  //依所選取迴路列出其上游
    {
        if (ListSource.SelectedValue == "") AddMsg("<script>alert('請先選取待選清單項目！');</script>");
        else
        {
            SqlDataSourceSource.SelectCommand = "SELECT [設備編號],[設備名稱] FROM [View_設備管理] "
                + "WHERE [設備編號] in (select [上游編號] from [View_設備迴路] where [下游編號]=" + ListSource.SelectedValue + ") ORDER BY [設備名稱]";
            ListSource.DataBind();
        }
    } 
    protected void LinkSourceDn_Click(object sender, EventArgs e)  //依所選取迴路列出其下游
    {
        if (ListSource.SelectedValue == "") AddMsg("<script>alert('請先選取待選清單項目！');</script>");
        else
        {
            SqlDataSourceSource.SelectCommand = "SELECT [設備編號],[設備名稱] FROM [View_設備管理] "
                + "WHERE [設備編號] in (select [下游編號] from [View_設備迴路] where [上游編號]=" + ListSource.SelectedValue + ") ORDER BY [設備名稱]";
            ListSource.DataBind();            
        }
    }     

    //選取放置地點------------------------------------------------------------------------ 
    protected void SelPlace_SelectedIndexChanged(object sender, EventArgs e)
    {
        if(SelPlace.SelectedValue=="空調群組"){
            ReadGroup();
        }else{
            Place_Selected(ListSource, SelPlace, SelPointer, SqlDataSourcePointer, SqlDataSourceSource);
        }
        
    }
    protected void Place_Selected(ListBox ListName, DropDownList SelPlace, DropDownList SelPointer, SqlDataSource SqlDS, SqlDataSource SqlDSList)   
    {
        ListName.Items.Clear();
        SqlDS.SelectCommand = "SELECT [定位編號],[定位名稱] FROM [定位設定] WHERE [定位方式]<>'坐標' AND [區域名稱]='" + SelPlace.SelectedValue + "' order by [定位名稱]";
        SelPointer.DataBind();
        Pointer_Selected(ListName, SelPointer, SqlDSList);
    }
    //選取定位名稱------------------------------------------------------------------------ 
    protected void SelPointer_SelectedIndexChanged(object sender, EventArgs e)
    {
        Pointer_Selected(ListSource, SelPointer, SqlDataSourceSource);
    }
    protected void Pointer_Selected(ListBox ListName, DropDownList SelPointer, SqlDataSource SqlDS)
    {
        string SQL = "[定位編號]=" + SelPointer.SelectedValue;
        if (SelPointer.SelectedValue == "") SQL = "[定位編號]=" + SelPointer.Items[0].Value;

        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
        Conn.Open();
        //[定位編號],[定位方式],[區域名稱],[定位名稱],[機櫃容量],[坐標X],[坐標Y],[坐標X1],[坐標Y1],[坐標X2],[坐標Y2]
        SqlCommand cmd = new SqlCommand("select * from [定位設定] where " + SQL, Conn);
        SqlDataReader dr = null;
        dr = cmd.ExecuteReader();
        if (dr.Read())
        {
            //定位方式若為分區，則顯示出以坐標定位之設備，以免選不到設備
            if (dr["定位方式"].ToString() == "分區") SQL = "[定位方式]<>'機櫃' and [坐標X] between " + dr["坐標X1"].ToString() + " and " + dr["坐標X2"].ToString() + " and [坐標Y] between " + dr["坐標Y1"].ToString() + " and " + dr["坐標Y2"].ToString();
            SqlDS.SelectCommand = "SELECT [設備編號],[設備名稱] FROM [View_設備管理] WHERE " + SQL + " ORDER BY [設備名稱]";
            ListName.DataBind();
        }
        cmd.Cancel(); cmd.Dispose(); dr.Close(); Conn.Close(); Conn.Dispose();
    }

    //顯示迴路路徑及放置地點------------------------------------------------------------------------ 
    protected void ListSource_SelectedIndexChanged(object sender, EventArgs e)
    {
        Conn_Selected(ListSource, lblSource, "設備");
    }

    protected void ListUpPower_SelectedIndexChanged(object sender, EventArgs e) 
    {
        Conn_Selected(ListUpPower, lblUpPower,"接電");
    }
    protected void ListDnPower_SelectedIndexChanged(object sender, EventArgs e) 
    {
        Conn_Selected(ListDnPower, lblDnPower, "接電");
    }

    protected void ListUpNet_SelectedIndexChanged(object sender, EventArgs e) 
    {
        Conn_Selected(ListUpNet, lblUpNet, "接網");
    }
    protected void ListDnNet_SelectedIndexChanged(object sender, EventArgs e)
    {
        Conn_Selected(ListDnNet, lblDnNet, "接網");
    }
    protected void ListUpCold_SelectedIndexChanged(object sender, EventArgs e) 
    {
        Conn_Selected(ListUpCold, lblUpCold, "空調");
    }
    protected void ListDnCold_SelectedIndexChanged(object sender, EventArgs e)
    {
        Conn_Selected(ListDnCold, lblDnCold, "空調");
    }

    protected void Conn_Selected(ListBox ListName,Label lbl,string tbl) //點選選項，列出放置地點與迴路路徑
    {
        lbl.Text = HttpUtility.HtmlEncode(GetValue("select [放置地點] from [View_設備管理] where [設備編號]=@LN", "@LN", ListName.SelectedValue));
        if (tbl == "接電" | tbl == "設備") lbl.Text = HttpUtility.HtmlEncode(lbl.Text + "AAD" + GetPath(ListName.SelectedValue, "接電", "") + "＊").Replace("AAD", "<br/>");
        if (tbl == "接網" | tbl == "設備") lbl.Text = HttpUtility.HtmlEncode(lbl.Text + "AAD" + GetPath(ListName.SelectedValue, "接網", "") + "＊").Replace("AAD", "<br/>");        
        if (tbl == "接冷" | tbl == "設備") lbl.Text = HttpUtility.HtmlEncode(lbl.Text + "AAD" + GetPath(ListName.SelectedValue, "空調", "") + "＊").Replace("AAD", "<br/>");        
    }
    protected string GetPath(string DevNo,string tbl, string tail)   //取得迴路路徑
    {
        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
        Conn.Open();
        SqlCommand cmd = new SqlCommand("select [上游編號],[設備名稱] from [" + tbl + "],[實體設備] where [" + tbl + "].[上游編號]=[實體設備].[設備編號] and [" + tbl + "].[下游編號]=@DN" , Conn);
        cmd.Parameters.AddWithValue("@DN", DevNo);
        SqlDataReader dr = null;
        dr = cmd.ExecuteReader();
        string path = "", UpPath = "", DnPath = "";

        int i = 0;
        while (dr.Read())
        {
            i++;
            DnPath = dr["設備名稱"].ToString() + " → " + tail;
            UpPath = GetPath(dr["上游編號"].ToString(), tbl, DnPath);
            if (i == 1) path = UpPath + DnPath;
            else path = path + "＊ AAD" + UpPath + DnPath;

            path = path.Replace(DnPath + DnPath, DnPath);
        }

        cmd.Cancel(); cmd.Dispose(); dr.Close(); Conn.Close(); Conn.Dispose();        
        return (path);
    }

    //權限管理------------------------------------------------------------------------ 
    protected Boolean RightCheck() //是否有權修改資料
    {
        string UserID = Session["UserID"].ToString().ToLower();
        string UserName = Session["UserName"].ToString();   //登入的UserName
        string UnitName = Session["UnitName"].ToString();   //登入的UnitName
        string Hw = lblMt.Text; //設備維護人員
        string Dno = TextDevNo.Text; if (Dno == "") Dno = "0";
        string Older = GetValue("select [維護人員] from [實體設備] where [設備編號]=@dno", "@dno" , Dno);

        if (UserID != "operator" & (InGroup(UserName, Hw) | InGroup(UserName, Older) | UnitName == "系統控制課" | UnitName == "電腦操作課" | InGroup(UserName, "網管小組"))) return (true);
        else return (false);
    }
    protected Boolean NetCheck(string Dno) //是否有權修改接網迴路
    {
        string Mt = "網管小組";     //網管小組權責設備接網上下游僅網管小組可設定
        string UserName = Session["UserName"].ToString();   //登入的UserName
        string Hw = lblMt.Text;     //設備維護人員
        string Dw = GetValue("select [維護人員] from [實體設備] where [設備編號]=@dno", "@dno" , Dno);   //上下游維護人員

        if ((Hw != Mt & Dw != Mt) | InGroup(UserName, Mt) | InGroup(UserName, "WII小組")) return (true);   //設備或上下游權責均非網管小組，或為網管成員
        else return (false);
    }
	protected Boolean NetDelCheck(string Dno) //是否有權 刪除 接網迴路
    {
        string Mt = "網管小組";     //網管小組權責設備接網上下游僅網管小組可設定
        string UserName = Session["UserName"].ToString();   //登入的UserName
        string Hw = lblMt.Text;     //設備維護人員
        string Dw = GetValue("select [維護人員] from [實體設備] where [設備編號]=@Dno", "Dno", Dno);   //上下游維護人員
        
        if ((Hw != Mt & Dw != Mt) | InGroup(UserName, Mt) | InGroup(UserName, "WII小組") | (Hw==UserName)| InGroup(UserName,Hw)) return (true);   //設備或上下游權責均非網管小組，或為網管成員 ，或為設備負責人(所在群組)
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
    protected void InsLifeLog(string SQL) //寫入生命履歷
    {
        string LifeNo = GetPKNo("履歷編號", "生命履歷").ToString(); //履歷編號
        string TblName = "設備迴路";    //表格名稱
        string PKno = TextDevNo.Text;   //主鍵編號
        string Mt = lblMt.Text;    //維護人員        
        string OldMt = GetValue("select [維護人員] from [實體設備] where [設備編號]=@TDN", "TDN", TextDevNo.Text);    //原負責人  
        string OldKeeper = GetValue("select [保管人員] from [View_設備管理] where [設備編號]=@TDN", "@TDN", TextDevNo.Text);    //原保管人
        string UN = Session["UserName"].ToString();   //登入的UserName：異動人員
        string LiftDT = DateTime.Now.ToString("yyyy/MM/dd HH:mm");  //異動日期
        string LifeIP = Request.ServerVariables["REMOTE_ADDR"].ToString();
        if(PKno.Length>=5){
            Mt = OldMt = OldKeeper = "環境小組";
        }
        Trace.Write(OldMt);Trace.Write(OldKeeper);

        List<SqlParameter> pars = new List<SqlParameter>();

        pars.Add(new SqlParameter("@LifeNo", LifeNo));
		pars.Add(new SqlParameter("@TblName", TblName));
		pars.Add(new SqlParameter("@PKno", PKno));
		pars.Add(new SqlParameter("@SQL", SQL.Replace("'", "''")));
		pars.Add(new SqlParameter("@Mt", Mt));	
        pars.Add(new SqlParameter("@OldMt", OldMt));
        pars.Add(new SqlParameter("@OldKeeper", OldKeeper));
		pars.Add(new SqlParameter("@UN", UN));
		pars.Add(new SqlParameter("@LiftDT", LiftDT));		
		pars.Add(new SqlParameter("@LifeIP", LifeIP));

        //ExecDbSQL("Insert into [生命履歷] values(" + LifeNo + ",'" + TblName + "'," + PKno + ",'" + SQL.Replace("'", "''") + "','" + Mt + "','" + OldMt + "','" + OldKeeper + "','" + UN + "','" + LiftDT + "','" + LifeIP + "')");
        ExecDbSQL("Insert INTO [生命履歷] VALUES( @LifeNo , @TblName, @PKno, @SQL, @Mt, @OldMt, @OldKeeper, @UN, @LiftDT, @LifeIP)", pars);
    
        //ExecDbSQL("Insert into [生命履歷] values(" + LifeNo + ",'" + TblName + "'," + PKno + ",'" + SQL.Replace("'", "''") + "','" + Mt + "','" + OldMt + "','" + OldKeeper + "','" + UN + "','" + LiftDT + "','" + LifeIP + "')");
    }
    protected int GetPKNo(string PKfield, string PKtbl) //取得主鍵編號
    {
        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
        Conn.Open();
        SqlCommand cmd = new SqlCommand("select max(" + PKfield + ") from " + PKtbl, Conn);
        SqlDataReader dr = null;
        dr = cmd.ExecuteReader();
        int PkNo = 1;
        if (dr.Read()) PkNo = int.Parse(dr[0].ToString()) + 1;
        cmd.Cancel(); cmd.Dispose(); dr.Close(); Conn.Close(); Conn.Dispose();

        return (PkNo);
    }
    protected void BtnEdit_Click(object sender, EventArgs e) //回編輯頁
    {
        OpenExecWindow("window.open('DevEdit.aspx?DevNo=" + TextDevNo.Text + "','_self');");
    }
    protected void BtnLife_Click(object sender, EventArgs e) //查詢生命履歷
    {
        Session["LifeSQL"] = "select * from [生命履歷] where [表格名稱]='設備迴路' and [主鍵編號]=" + TextDevNo.Text + " order by [履歷編號] desc";
        OpenExecWindow("window.open('../Life/LifeLog.aspx?Search=Conn&Tbl=設備迴路&PK=" + HttpUtility.HtmlEncode(TextDevNo.Text) + "','_self');");
    } 
    //公用函數------------------------------------------------------------------------ 
    protected DataSet RunQuery(SqlCommand sqlQuery) //讀取DB資訊
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

    protected void ExecDbSQL(string SQL) //執行資料庫異動
    {
        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
        Conn.Open();
        SqlCommand cmd = new SqlCommand(SQL, Conn);
        //Response.Write(cmd.CommandText);
        //Response.End();
        cmd.ExecuteNonQuery();
        cmd.Cancel(); cmd.Dispose(); Conn.Close(); Conn.Dispose();
    }
    protected void ExecDbSQL(string SQL, List<SqlParameter> pars) //執行資料庫異動
    {
        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
        Conn.Open();
        SqlCommand cmd = new SqlCommand(SQL, Conn);
        cmd.Parameters.AddRange(pars.ToArray());
        //Response.Write(cmd.CommandText);
        //Response.End();
        cmd.ExecuteNonQuery();
        cmd.Cancel(); cmd.Dispose(); Conn.Close(); Conn.Dispose();
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

    protected string GetValue(string SQL, string key, string value)   //取得單一資料
    {
        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
        Conn.Open();
        SqlCommand cmd = new SqlCommand(SQL, Conn);
        cmd.Parameters.AddWithValue(key, value);
        SqlDataReader dr = null;
        dr = cmd.ExecuteReader();
        string cfg = ""; if (dr.Read()) cfg = dr[0].ToString();

        cmd.Cancel(); cmd.Dispose(); dr.Close(); Conn.Close(); Conn.Dispose();

        return (cfg);
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

    protected void ReadGroup(){
        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
        Conn.Open();
        string sql = @"select * FROM [dbo].[實體群組]";
        SqlCommand cmd = new SqlCommand(sql, Conn);
        DataSet ds = RunQuery(cmd);
        foreach(DataRow row in ds.Tables[0].Rows){
            ListSource.Items.Add(new ListItem(String.Format("<{0}> {1}",row[0].ToString(), row[1].ToString()), row[0].ToString()));
        }
        
    }
}