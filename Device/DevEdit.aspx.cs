using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.Configuration;
using System.Data;
using System.Data.SqlClient;

public partial class Device_DevEdit : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        if (!IsPostBack)
        {
			InitSession();//假如Session沒有設定數值，初始化設置 //Revise
			
            ReadAssets(); //讀取資產編號下拉式選單
            
            for (int i = 0; i <= 42; i++)   //產生機櫃高度之選項
            {
                ListItem ItemX = new ListItem(); ItemX.Text = i.ToString(); ItemX.Value = i.ToString(); SelRU.Items.Add(ItemX);
            }
            for (int i = 0; i <= 30; i++)   //產生空間大小之選項
            {
                ListItem ItemX = new ListItem(); ItemX.Text = i.ToString(); ItemX.Value = i.ToString(); SelSpace.Items.Add(ItemX);
            }

            //填入預設值
            TextPointerNo.Text = "0";
            TextVoltage.Text = "0";
            TextSpecCurrent.Text = "0";
            SelRU.SelectedIndex = 1;
            SelSpace.SelectedIndex = 1;            
        }
    }
	
	protected void InitSession(){ //初始化設置Session //Revise
		if (Session["UserID"] == null){			//取得系統資訊，除了UserID = Request.LogonUserIdentity.Name，全部照抄修改
			string Port = Request.Url.Port.ToString();
			string Sv = Server.MachineName.ToUpper(); 
			string Ws = Request.UserHostName.ToUpper();
			string IP = Request.ServerVariables["REMOTE_ADDR"].ToString();
			string IO = "安外"; if (IP.IndexOf("172.")==0 | IP.IndexOf("61.60") == 0) IO = "安內";
			
			string UserID = Request.LogonUserIdentity.Name;//我覺得全系統含master那邊要全改成這句判斷，不要用cookie。//Revise
			string UserName = "", UnitID = "", UnitName = "";
			
			int pos = UserID.IndexOf("\\");
			if( pos >= 0 )
				UserID = UserID.Substring(pos + 1);
			if (Sv != "SOS-VM18" & Sv != "10.6.1.11") UserID = "SsmMic"; //非sos-vm16一律當SsmMic
			if (UserID == "W2KADM" | UserID == "OPERATOR" | UserID == "OP1" | UserID == "A389") UserID = "operator";
			Session["UserID"] = UserID;
			if (UserID != ""){
				using(SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString)){
					SqlCommand cmd = new SqlCommand("Select Kind,Item from Config where Mark= @MarkName ", Conn);
					cmd.Parameters.Add("@MarkName", SqlDbType.NVarChar).Value = UserID;
					Conn.Open();
					SqlDataReader dr = cmd.ExecuteReader();
					if (dr.Read())
					{
						UserName = dr[1].ToString(); 
                        UnitName = dr[0].ToString(); 
                        dr.Close();
						cmd.CommandText = "Select Config from Config where Kind='數值資訊組' and Item= @ItemName";
						cmd.Parameters.Add("@ItemName", SqlDbType.VarChar).Value = UnitName;
						dr = cmd.ExecuteReader();
						if (dr.Read())
						{
							UnitID = dr[0].ToString();
							Session["UserID"] = UserID;
							Session["UserName"] = UserName;
							Session["UnitID"] = UnitID;
							Session["UnitName"] = UnitName;
						}
					}
				}
			}
		}
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

    protected void Page_PreRenderComplete(object sender, EventArgs e)   //放Page_Load無法顯示
    {
        if (Request["DevNo"] != null) TextDevNo.Text = Request["DevNo"];
        else
        {
            if (Request["PropA"] != null & !IsPostBack)   //由秘總帶入財產資料
            {
                TextPropNoA.Text = Request["PropA"].ToString();
                TextPropNoB.Text = Request["PropB"].ToString();
                LinkPropIn_Click(null, null);
            }
        }

        if (TextDevNo.Text == "")
        {
            BtnEdit.Enabled = false;
            BtnDel.Enabled = false;
        }
        else
        {
            BtnEdit.Enabled = true;
            BtnDel.Enabled = true;
        }

        if (!IsPostBack)   //不知為何需要DataBind()才能顯示，且須在SelDevType.Selected之後
        {
            if (Request["DevType"] != null)    //設定設備型態預設值，不能放在Page_Load，因選項在其之後才產生，且要放在ReadHelp之前
            {
                for (int i = 0; i < SelDevType.Items.Count; i++) if (SelDevType.Items[i].Value == Request["DevType"].ToString()) SelDevType.SelectedIndex = i;
                SelDevKind.DataBind();  //連帶選項需要觸發
                if (Request["DevKind"] != null) for (int i = 0; i < SelDevKind.Items.Count; i++) if (SelDevKind.Items[i].Value == Request["DevKind"].ToString()) SelDevKind.SelectedIndex = i;
            }
            for (int i = 0; i < SelDevStatus.Items.Count; i++) if (SelDevStatus.Items[i].Text == "建置中") SelDevStatus.Items[i].Selected = true;

            SelSource.DataBind();

            if (TextDevNo.Text != "") ReadAll(int.Parse(TextDevNo.Text));  //讀取所有值

            ShowHideHelp();  //顯示或隱藏欄位說明
        }

        ReadDevHelp();  //隱藏欄位
        ReadAPHelp();  //隱藏欄位         

        SelRepair.DataBind();
        for (int i = 0; i < SelRepair.Items.Count; i++) if (SelRepair.Items[i].Value == TextRepair.Text) SelRepair.SelectedIndex = i;
        GetRepair();    //顯示維護廠商名

        GetMaintainor();    //顯示維護人員清單    
        SelAssets_SelectedIndexChanged(null, null);   //顯示資產描述
        SelPowerOff_SelectedIndexChanged(null, null);   //關機順序之說明
    }
    protected void SelPowerOff_SelectedIndexChanged(object sender, EventArgs e)   //關機順序之說明
    {
        lblPowerOff.Text = GetValueWithItemName("IDMS", "select [Memo] from [Config] where [Kind]='關機順序' and [Item]= @ItemName", SelPowerOff.SelectedValue);//Revise
    }

    protected void ReadAll(int DevNo)  //讀取所有值
    {
        ReadDevice(DevNo);  //讀取實體設備
        ReadProp();    //lbl讀取秘總財產
        ReadAP(DevNo);  //讀取作業主機
        //讀取設備迴路
        ReadDevTree(DevNo.ToString(), "接電", "Up", lblUpPower);
        ReadDevTree(DevNo.ToString(), "接電", "Dn", lblDnPower);
        ReadDevTree(DevNo.ToString(), "接網", "Up", lblUpNet);
        ReadDevTree(DevNo.ToString(), "接網", "Dn", lblDnNet);
        ReadDevTree(DevNo.ToString(), "空調", "Up", lblUpCold);
        ReadDevTree(DevNo.ToString(), "空調", "Dn", lblDnCold);

        

        string CheckYear = GetValue("IDMS", "select max(清查年度) from [機器清查]");
        List<SqlParameter> pars = new List<SqlParameter>();
        pars.Add(new SqlParameter("@CheckYear", CheckYear));
        pars.Add(new SqlParameter("@DevNo", DevNo));
        //lblDevCheck.Text = GetValue("IDMS", "select [清查年度] + ' ' + [清查人員] + '(' + [清查狀態] + ') ' + [清查結果] + ' ' + [備註說明] from [機器清查] where [清查年度]='" + CheckYear + "' and [設備編號]=" + DevNo);   //取得機器清查
        lblDevCheck.Text = GetValue("IDMS", "select [清查年度] + ' ' + [清查人員] + '(' + [清查狀態] + ') ' + [清查結果] + ' ' + [備註說明] from [機器清查] where [清查年度]= @CheckYear and [設備編號]=@DevNo", pars);   //取得機器清查
        if (lblDevCheck.Text.IndexOf("(已結案)") < 0) lblDevCheck.ForeColor = System.Drawing.Color.Red;
        if (lblDevCheck.Text == "")
        {
            lblDevCheck.Text = "尚未列入清查項目!";
            BtnDevCheck.Visible = false;
        }

        lblChecks.Text = GetChecks(DevNo);   //取得資料查核
        lblRepeat.Text = HasSameHost(DevNo).Replace("\\n", "<br/>");   //取得重複查核        
    }

    protected void ReadDevice(int DevNo)    //讀取實體設備
    {
        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
        Conn.Open();
        SqlCommand cmd = new SqlCommand("SELECT * FROM [實體設備],[定位設定] WHERE [實體設備].[定位編號]=[定位設定].[定位編號] and [設備編號]= @DeviceNumber", Conn);
		cmd.Parameters.Add("@DeviceNumber", SqlDbType.Int).Value = DevNo;//Revise
        SqlDataReader dr = null;
        dr = cmd.ExecuteReader();

        if (dr.Read())
        {
            for (int i = 0; i < SelAssets.Items.Count; i++) if (SelAssets.Items[i].Value == dr["資產編號"].ToString()) SelAssets.SelectedIndex = i;
            for (int i = 0; i < SelDevType.Items.Count; i++) if (SelDevType.Items[i].Value == dr["設備型態"].ToString()) SelDevType.SelectedIndex = i;

            SelDevKind.DataBind();  //連帶選項需要觸發
            for (int i = 0; i < SelDevKind.Items.Count; i++) if (SelDevKind.Items[i].Value == dr["設備種類"].ToString()) SelDevKind.SelectedIndex = i;
            TextDevName.Text = dr["設備名稱"].ToString();
            TextPurpose.Text = dr["設備用途"].ToString();

            TextPropNoA.Text = dr["財產編號A"].ToString();
            TextPropNoB.Text = dr["財產編號B"].ToString();

            TextBrand.Text = dr["廠牌"].ToString();
            TextStyle.Text = dr["型式"].ToString();

            TextKeeper.Text = dr["保管人員"].ToString();

            TextPointerNo.Text = dr["定位編號"].ToString();
            for (int i = 0; i < SelPlace.Items.Count; i++) if (SelPlace.Items[i].Value == dr["區域名稱"].ToString()) SelPlace.SelectedIndex = i;
            SelPointer.DataBind();  //連帶選項需要觸發
            if (dr["定位方式"].ToString() != "坐標")
            {
                for (int i = 0; i < SelPointer.Items.Count; i++) if (SelPointer.Items[i].Value == dr["定位編號"].ToString()) SelPointer.SelectedIndex = i;
            }
            else
            {
                ListItem ItemX = new ListItem();
                ItemX.Text = dr["定位名稱"].ToString();
                ItemX.Value = dr["定位編號"].ToString();
                SelPointer.Items.Add(ItemX);
                SelPointer.SelectedIndex = SelPointer.Items.Count - 1;
            }
            for (int i = 0; i < SelRU.Items.Count; i++) if (SelRU.Items[i].Value == dr["高度"].ToString()) SelRU.SelectedIndex = i;
            for (int i = 0; i < SelSpace.Items.Count; i++) if (SelSpace.Items[i].Value == dr["空間大小"].ToString()) SelSpace.SelectedIndex = i;

            TextSource.Text = dr["取得來源"].ToString();
            TextSpec.Text = dr["規格"].ToString();
            TextSerial.Text = dr["硬體序號"].ToString();
            TextiLoIP.Text = dr["iLoIP"].ToString();

            TextVoltage.Text = dr["用電電壓"].ToString();
            TextSpecCurrent.Text = dr["額定電流"].ToString();

            TextRepair.Text = dr["維護廠商"].ToString();
            string HwUnit = GetUnit(dr["維護人員"].ToString());
            for (int i = 0; i < SelHwUnit.Items.Count; i++) if (SelHwUnit.Items[i].Value == HwUnit) SelHwUnit.SelectedIndex = i;
            SelMaintainor.DataBind();  //連帶選項需要觸發
            for (int i = 0; i < SelMaintainor.Items.Count; i++) if (SelMaintainor.Items[i].Value == dr["維護人員"].ToString()) SelMaintainor.SelectedIndex = i;

            for (int i = 0; i < SelDevStatus.Items.Count; i++) if (SelDevStatus.Items[i].Value == dr["設備狀態"].ToString()) SelDevStatus.SelectedIndex = i;
            for (int i = 0; i < SelPowerOff.Items.Count; i++) if (SelPowerOff.Items[i].Value == dr["關機順序"].ToString()) SelPowerOff.SelectedIndex = i;


            DateTime hd;
            if(DateTime.TryParse(dr["建立日期"].ToString(), out hd)){
                LblCreateDT.Text = DateTime.Parse(dr["建立日期"].ToString()).ToString("yyyy/MM/dd HH:mm");
            }
            else{LblCreateDT.Text = null;}
            
            LblCreater.Text = dr["建立人員"].ToString();
            if(DateTime.TryParse(dr["修改日期"].ToString(), out hd)){
                LblUpdateDT.Text = DateTime.Parse(dr["修改日期"].ToString()).ToString("yyyy/MM/dd HH:mm");
            }
            else{LblUpdateDT.Text = null;}
            
            LblUpdater.Text = dr["修改人員"].ToString();
            TextMemo.Text = dr["備註說明"].ToString();
        }

        cmd.Cancel(); cmd.Dispose(); dr.Close(); Conn.Close(); Conn.Dispose();
    }

    protected void ReadAP(int DevNo)    //讀取作業主機
    {
        string ApCount = GetValueWithDevNo("IDMS", "select Count([作業編號]) from [作業主機] where [設備編號]= @DevNo", DevNo.ToString());//Revise
        
        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
        Conn.Open();
        SqlCommand cmd = new SqlCommand("SELECT * FROM [作業主機] WHERE  [設備編號]= @DeviceNumber", Conn);
		cmd.Parameters.Add("@DeviceNumber", SqlDbType.Int).Value = DevNo;//Revise
        SqlDataReader dr = null;
        dr = cmd.ExecuteReader();
        if (dr.Read())
        {
            TextHostName.Text = dr["主機名稱"].ToString();
            lblApNo.Text = "作業編號:" + dr["作業編號"].ToString();

            if (ApCount == "0") lblAP.Text = "(無)";
            else if (ApCount == "1") lblAP.Text = "(" + dr["作業編號"].ToString() + ") " + dr["主機名稱"].ToString() + " " + GetSysAlias(dr["系統編號"].ToString(), "SysName");
            else lblAP.Text = "(多重主機繫結)";

            TextIP.Text = dr["IP"].ToString();
            TextMosIP.Text = dr["監控IP"].ToString();
            for (int i = 0; i < SelSaveIO.Items.Count; i++) if (SelSaveIO.Items[i].Value == dr["安內外"].ToString()) SelSaveIO.SelectedIndex = i;
            for (int i = 0; i < SelOnCall.Items.Count; i++) if (SelOnCall.Items[i].Value == dr["緊急程度"].ToString()) SelOnCall.SelectedIndex = i;
        }

        cmd.Cancel(); cmd.Dispose(); dr.Close(); Conn.Close(); Conn.Dispose();
    }

    protected string GetSysAlias(string SysNo, string Kind)   //取得系統相關資訊
    {
        if (Kind == "SysName") 
        return (GetValueWithSysNo("IDMS", "select [系統全名] from [View_系統資源] where [資源編號]= @SysNo", SysNo));//Revise
        else
        return (GetValueWithSysNo("IDMS", @"select [資源功能]" +
         " -- " + "[備註說明] from [系統資源] where [資源編號]= @SysNo", SysNo));//Revise
    }

    protected void ReadDevTree(string DevNo,string tbl,string UpDn,Label lbl)    //讀取設備迴路選單
    {
		SqlCommand sqlQuery = new SqlCommand();
        if(UpDn=="Up"){//ReviseX
			sqlQuery.CommandText = "SELECT [設備編號],[設備名稱] FROM [實體設備],[" + tbl + "] WHERE [實體設備].[設備編號]=[" + tbl + "].[上游編號] and [" + tbl + "].[下游編號] = @DeviceNumber ORDER BY [設備名稱]";
			sqlQuery.Parameters.Add("@DeviceNumber", SqlDbType.Int).Value = DevNo;
		}
        else{//ReviseX
			sqlQuery.CommandText = "SELECT [設備編號],[設備名稱] FROM [實體設備],[" + tbl + "] WHERE [實體設備].[設備編號]=[" + tbl + "].[下游編號] and [" + tbl + "].[上游編號] = @DeviceNumber ORDER BY [設備名稱]";
			sqlQuery.Parameters.Add("@DeviceNumber", SqlDbType.Int).Value = DevNo;
		}
		
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

    protected string GetUnit(string Pname) //讀取人員單位或分組名稱
    {
        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
        Conn.Open();
        SqlCommand cmd = new SqlCommand("SELECT [課別] FROM [View_組織架構] WHERE [性質]='員工' and [成員]=@Pname ", Conn);
        cmd.Parameters.AddWithValue("@Pname", Pname);
        SqlDataReader dr = null;
        dr = cmd.ExecuteReader();
        string cfg = "";
        if (dr.Read()) cfg = dr[0].ToString(); //先讀數值資訊組各課成員(各課優先)
        else  //再讀其它中心或維護群組(分組次之)
        {
            cmd.Cancel(); cmd.Dispose(); dr.Close();
            //cmd = new SqlCommand("SELECT Kind FROM Config WHERE Kind not in (select Item from Config where Kind='數值資訊組') and Item='" + Pname + "'", Conn);
            cmd = new SqlCommand("SELECT Kind FROM Config WHERE Kind not in (select Item from Config where Kind='數值資訊組') and Item= @Pname ", Conn);
            cmd.Parameters.AddWithValue("@Pname", Pname);
            
            dr = cmd.ExecuteReader();
            if (dr.Read()) cfg = dr[0].ToString();
        }
        cmd.Cancel(); cmd.Dispose(); dr.Close(); Conn.Close(); Conn.Dispose();

        return (cfg);
    }

    protected string GetChecks(int PkNo)   //取得資料查核
    {
        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
        Conn.Open();
        SqlCommand cmd = new SqlCommand("select [查核結果] FROM [View_資料查核] WHERE [表格名稱] in ('實體設備','設備迴路','接網迴路') and [主鍵編號] = @PrimaryKeyNumber", Conn);
		cmd.Parameters.Add("@PrimaryKeyNumber", SqlDbType.BigInt).Value = PkNo;//
        SqlDataReader dr = null;
        dr = cmd.ExecuteReader();
        string cfg = "";
        while (dr.Read()) cfg = (cfg + dr[0].ToString() + ' ');

        cmd.Cancel(); cmd.Dispose(); dr.Close(); Conn.Close(); Conn.Dispose();

        return (cfg);
    }

    protected void ReadDevHelp() //讀取欄位說明，並根據不同設備型態隱藏不同欄位
    {
        SqlCommand sqlQuery = new SqlCommand();
        sqlQuery.CommandText = "select Item,Memo,Config from Config where Kind='實體設備' order by Mark";
        DataSet ds = RunQuery(sqlQuery);

        string obj = "";
        int DevTypeIdx = 0;

        switch (SelDevType.SelectedValue)  //設備型態欄位顯隱索引
        {
            case "系統設備": DevTypeIdx = 0; break;
            case "網路設備": DevTypeIdx = 1; break;
            case "環境設備": DevTypeIdx = 2; break;
            case "零附件": DevTypeIdx = 3; break;
            case "週邊設備": DevTypeIdx = 4; break;
            case "辦公設備": DevTypeIdx = 5; break;
        }

        if (ds.Tables.Count > 0)
        {
            foreach (DataRow row in ds.Tables[0].Rows)
            {
                obj = row[2].ToString().Substring(7);   //物件名稱

                TableRow Sel = (TableRow)form1.FindControl("row" + obj);    //欄位隱藏
                if (row[2].ToString().Substring(DevTypeIdx, 1) == "0") Sel.Visible = false;
                else Sel.Visible = true;

                Label helpObj = (Label)form1.FindControl("help" + obj); //說明文字
                helpObj.Text = row[1].ToString().Replace("\r\n", "<br />");
            }
        }

        sqlQuery.Cancel(); ds.Dispose();
    }

    protected void ReadAPHelp() //讀取欄位說明 //根據不同設備型態隱藏不同欄位
    {
        SqlCommand sqlQuery = new SqlCommand();
        sqlQuery.CommandText = "select Item,Memo from Config where Kind='作業主機' order by Mark";
        DataSet ds = RunQuery(sqlQuery);

        if (ds.Tables.Count > 0)
        {
            foreach (DataRow row in ds.Tables[0].Rows)
            {
                switch (row[0].ToString())
                {
                    case "主機名稱": helpHostName.Text = row[1].ToString().Replace("\r\n", "<br />"); break;
                    case "IP": helpIP.Text = row[1].ToString().Replace("\r\n", "<br />"); break;
                    case "安內外": helpSaveIO.Text = row[1].ToString().Replace("\r\n", "<br />"); break;
                    case "緊急程度": helpOnCall.Text = row[1].ToString().Replace("\r\n", "<br />"); break;
                }
            }
        }

        if (SelDevType.SelectedValue != "網路設備")    //僅網路設備需要顯示
        {
            rowHostName.Visible = false; rowIP.Visible = false; rowSaveIO.Visible = false; rowOnCall.Visible = false;
        }
        else
        {
            rowHostName.Visible = true; rowIP.Visible = true; rowSaveIO.Visible = true; rowOnCall.Visible = true;
        }

        sqlQuery.Cancel(); ds.Dispose();
    }
	
	//Revise 需要改寫5個
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

    protected void BtnAdd_Click(object sender, EventArgs e) //按下新增按鈕
    {
        Literal Msg = new Literal();
        string IsSame = HasSameHost(0).Replace("<br/>", "");        
		
		InitSession();//假如Session沒有設定數值，初始化設置 //Revise
		string outans = "";
        if (RightCheck(ref outans))
        {
            string DevNo = TextDevNo.Text; if (DevNo == "") DevNo = "0";

            if (TextDevName.Text == "")
            {
                Msg.Text = "<script>alert('您未填設備名稱!');</script>";
            }
            else if (SelMaintainor.SelectedValue == "")
            {
                Msg.Text = "<script>alert('您未填維護人員!');</script>";
            }
            else if (SelDevStatus.SelectedValue == "已報廢" & GetValueWithDevNo("IDMS", "select count(*) from [作業主機] where [設備編號]= @DevNo", DevNo) != "0")//Revise
            {
                Msg.Text = "<script>alert('請先刪除安裝的系統作業，再將狀態設為已報廢!');</script>";
            }
            else
            {
                LblCreateDT.Text = DateTime.Now.ToString("yyyy/MM/dd HH:mm");
                LblCreater.Text = Session["UserName"].ToString();
                if (TextVoltage.Text == "") TextVoltage.Text = "0";
                if (TextSpecCurrent.Text == "") TextSpecCurrent.Text = "0";
                LblUpdateDT.Text = LblCreateDT.Text;
                LblUpdater.Text = LblCreater.Text;

                BlankProp();    //若設備財產等於秘總財產，清空之
                InsData();

                Msg.Text = ("<script>alert('新增資料[" + (TextDevName.Text) + "] 完成！\\n\\n" + (GetChecks(int.Parse(TextDevNo.Text))) + "\\n" + IsSame + "');window.open('DevEdit.aspx?DevNo=" + (TextDevNo.Text) + "','_self');</script>");  //新增後關視窗是因無法改PK
            }
        }
        else Msg.Text = "<script>alert('" + outans +"\\n您不是維護人員，沒有新增 [" + TextDevName.Text + "] 的權限！\\n您可將維護人員設為自己或自己所屬的小組，但不能改為別人！若有需要，可請當事者將這筆資料的維護人員設為自己。');</script>";

        Page.Controls.Add(Msg);
    }

    protected void BlankProp() //若設備財產等於秘總財產，清空之
    {
        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
        Conn.Open();
        SqlCommand cmd = new SqlCommand("select [廠牌],[型式],[保管人員] from [財產主檔] where [財產編號A]= @PropertyA and [財產編號B]= @PropertyB ", Conn);//Revise
		//cmd.Parameters.AddWithValue("@PropertyA", TextPropNoA.Text);//Revise
		//cmd.Parameters.AddWithValue("@PropertyB", TextPropNoB.Text);//Revise
		cmd.Parameters.Add("@PropertyA", SqlDbType.NVarChar).Value = TextPropNoA.Text;//Revise
		cmd.Parameters.Add("@PropertyB", SqlDbType.NVarChar).Value = TextPropNoB.Text;//Revise
		
        SqlDataReader dr = null;
        dr = cmd.ExecuteReader();
        if (dr.Read())
        {
            if (dr[0].ToString() == TextBrand.Text) TextBrand.Text = "";
            if (dr[1].ToString() == TextStyle.Text) TextStyle.Text = "";
            if (dr[2].ToString() == TextKeeper.Text) LinkKeeperClear_Click(null, null);
        }
        cmd.Cancel(); cmd.Dispose(); dr.Close(); Conn.Close(); Conn.Dispose();
    }

    protected int GetPKNo(string PKfield, string PKtbl) //取得主鍵編號
    {
        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
        Conn.Open();
        SqlCommand cmd = new SqlCommand("select max(" + PKfield + ") from " + PKtbl, Conn);
        SqlDataReader dr = null;
        dr = cmd.ExecuteReader();
        int PkNo = 1;
        if (dr.Read()) PkNo = (int.Parse(dr[0].ToString()) + 1);
        cmd.Cancel(); cmd.Dispose(); dr.Close(); Conn.Close(); Conn.Dispose();

        return (PkNo);
    }

    protected string HasSameHost(int DevNo) //判斷是否有重複之設備或主機名稱     雙斜線避免alert讀不到
    {
        string Msg = "";    //回傳重複訊息

        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
        Conn.Open();
        SqlCommand cmd = new SqlCommand("select * from [View_設備管理] where [設備名稱]<>'' and [設備名稱]= @DeviceName "
            + " and [設備種類] not in (select [Item] from [Config] where [Kind] in (select [Item] from [Config] where [Kind]='設備型態') and [Config]='重複')"
            + " and [設備編號]<> @DeviceNumber", Conn);//Revise
		//cmd.Parameters.AddWithValue("@DeviceName", TextDevName.Text);//Revise
		//cmd.Parameters.AddWithValue("@DeviceNumber", DevNo);//Revise
		cmd.Parameters.Add("@DeviceName", SqlDbType.NVarChar).Value = TextDevName.Text;//Revise
		cmd.Parameters.Add("@DeviceNumber", SqlDbType.Int).Value = DevNo;//Revise
		
        SqlDataReader dr = null;
        dr = cmd.ExecuteReader();
        while (dr.Read())
        {
            Msg = Msg + "設備名稱重複:" + dr["設備編號"].ToString() + " " + dr["設備種類"].ToString() + " " + dr["設備名稱"].ToString() + " " + dr["財產編號A"].ToString() + " " + dr["財產編號B"].ToString() + "\\n";
        }
        dr.Close();
		
        cmd.CommandText = "select * from [View_設備管理] where [財產編號A]<>'' and [財產編號B]<>'' and [財產編號A]<>'00000000000'"
            + " and [財產編號A]= @PropertyA and [財產編號B]= @PropertyB and [設備編號]<> @DeviceNumber";//Revise
		cmd.Parameters.Clear();//Revise
		//cmd.Parameters.AddWithValue("@PropertyA", TextPropNoA.Text);//Revise
		//cmd.Parameters.AddWithValue("@PropertyB", TextPropNoB.Text);//Revise
		//cmd.Parameters.AddWithValue("@DeviceNumber", DevNo);//Revise
		cmd.Parameters.Add("@PropertyA", SqlDbType.NVarChar).Value = TextPropNoA.Text;//Revise
		cmd.Parameters.Add("@PropertyB", SqlDbType.NVarChar).Value = TextPropNoB.Text;//Revise
		cmd.Parameters.Add("@DeviceNumber", SqlDbType.Int).Value = DevNo;//Revise
		
		
        dr = cmd.ExecuteReader();
        while (dr.Read())
        {
            Msg = Msg + "財產編號重複：" + dr["設備編號"].ToString() + " " + dr["設備種類"].ToString() + " " + dr["設備名稱"].ToString() + " " + dr["財產編號A"].ToString() + " " + dr["財產編號B"].ToString() + "\\n";
        }
        dr.Close();

        if (SelDevType.SelectedValue == "網路設備")
        {
            cmd.CommandText = "select * from [作業主機] where [主機名稱]<>'' and [主機名稱]= @HostName "
                + " and [設備編號] not in (select [設備編號] from [實體設備] where [設備種類]='局屬網路')"
                + " and [設備編號]<> @DeviceNumber";//Revise
			cmd.Parameters.Clear();//Revise
			//cmd.Parameters.AddWithValue("@HostName", TextHostName.Text);//Revise
			//cmd.Parameters.AddWithValue("@DeviceNumber", DevNo);//Revise
			cmd.Parameters.Add("@HostName", SqlDbType.NVarChar).Value = TextHostName.Text;//Revise
			cmd.Parameters.Add("@DeviceNumber", SqlDbType.Int).Value = DevNo;//Revise
			
            dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                Msg = Msg + "主機名稱重複:" + dr["作業編號"].ToString() + " " + dr["主機名稱"].ToString() + " " + dr["IP"].ToString() + " " + GetSysAlias(dr["系統編號"].ToString(), "SysName") + "\\n";
            }
            dr.Close();

            cmd.CommandText = "select * from [作業主機] where [IP]<>'' and [IP]= @IP and [設備編號]<> @DeviceNumber";//Revise
			cmd.Parameters.Clear();//Revise
			//cmd.Parameters.AddWithValue("@IP", TextIP.Text);//Revise
			//cmd.Parameters.AddWithValue("@DeviceNumber", DevNo);//Revise
			cmd.Parameters.Add("@IP", SqlDbType.NVarChar).Value = TextIP.Text;//Revise
			cmd.Parameters.Add("@DeviceNumber", SqlDbType.Int).Value = DevNo;//Revise
			
            dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                Msg = Msg + "IP位址重複；" + dr["作業編號"].ToString() + " " + dr["主機名稱"].ToString() + " " + dr["IP"].ToString() + " " + GetSysAlias(dr["系統編號"].ToString(), "SysName") + "\\n";
            }
            dr.Close();
        }

        cmd.Cancel(); cmd.Dispose(); Conn.Close(); Conn.Dispose();

        
        return (Msg);
    }

    protected Boolean ExistsCheck(string SQL, string TextDevNo) //判斷設備編號or資產編號是否合法 //Revise
    {
        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
        Conn.Open();
        SqlCommand cmd = new SqlCommand(SQL, Conn);
		//cmd.Parameters.AddWithValue("@DeviceNumber", TextDevNo);//Revise
		cmd.Parameters.Add("@DeviceNumber", SqlDbType.Int).Value = TextDevNo;//Revise
        SqlDataReader dr = null;
        dr = cmd.ExecuteReader();
        Boolean TF = false; if (dr.Read()) TF = true;
        cmd.Cancel(); cmd.Dispose(); dr.Close(); Conn.Close(); Conn.Dispose();

        return (TF);
    }

    protected void BtnDel_Click(object sender, EventArgs e)    //真刪除，請注意生命履歷要詳細記錄所有所刪資料之詳細所有欄位資料
    {
        Literal Msg = new Literal();
		InitSession();//假如Session沒有設定數值，初始化設置 //Revise
        string outans = "";
        if (RightCheck())
        {
            string DevNo = TextDevNo.Text; if (DevNo == "") DevNo = "0";
            
            if (SelDevType.SelectedValue != "網路設備" & GetValueWithDevNo("IDMS", "select count(*) from [作業主機] where [設備編號]= @DevNo", DevNo) != "0")//Revise
                Msg.Text = "<script>alert('請先將所有安裝於本設備之系統作業轉移到其它設備或移除(清除設備編號)，再刪除本設備！');</script>";
            else if (ExistsCheck("select [下游編號] from [接電] where [上游編號]= @DeviceNumber" , TextDevNo.Text))//Revise
                Msg.Text = "<script>alert('請先移除或轉移下游接電迴路再刪除本設備！');</script>";
            else if (ExistsCheck("select [下游編號] from [接網] where [上游編號]= @DeviceNumber" , TextDevNo.Text))//Revise
                Msg.Text = "<script>alert('請先移除或轉移下游接網迴路再刪除本設備！');</script>";
            else if (int.Parse(TextDevNo.Text) <= 0) 
                Msg.Text = "<script>alert('系統專用，禁止刪除！');</script>";
            else
            {
                string SQL = " [實體設備]‧[設備編號]=" + TextDevNo.Text;

                string LifeSQL = "刪除 [" + SelDevType.SelectedValue + "]．[" + SelDevKind.SelectedValue + "]‧[" + TextDevName.Text + "] ： " + SQL
                        + "，原始資料：" + GetInsDevSQL(TextDevNo.Text);
                if (SelDevType.SelectedValue == "網路設備") LifeSQL = LifeSQL + "；" + GetInsApSQL("?", TextDevNo.Text);
                InsLifeLog(LifeSQL);

                DeleteDevNo("DELETE FROM [接電] WHERE [上游編號]= @DevNo or [下游編號]= @DevNo", TextDevNo.Text);//Revise
                DeleteDevNo("DELETE FROM [接網] WHERE [上游編號]= @DevNo or [下游編號]= @DevNo", TextDevNo.Text);//Revise
                DeleteDevNo("DELETE FROM [作業主機] WHERE [設備編號]= @DevNo", TextDevNo.Text);//Revise
                DeleteDevNo("DELETE FROM [實體設備] WHERE [設備編號]= @DevNo", TextDevNo.Text);//Revise

                Msg.Text = "<script>alert('完成刪除[" + TextDevName.Text + "]設備、接電、接網迴路及系統作業！');window.close();</script>";
            }
        }
        else Msg.Text = "<script>alert('您沒有刪除 [" + TextDevName.Text + "] 的權限！');</script>";

        Page.Controls.Add(Msg);
    }

    protected void BtnEdit_Click(object sender, EventArgs e) //按下修改按鈕
    {
		InitSession();//假如Session沒有設定數值，初始化設置 //Revise
        Literal Msg = new Literal();
        string IsSame = HasSameHost(int.Parse(TextDevNo.Text)).Replace("<br/>", "");
        string UserName = Session["UserName"]==null?"":Session["UserName"].ToString();
        string LifeSQL = "",DevSQL="";
        string outans = "";
        if (RightCheck(ref outans))
        {
            string DevNo = TextDevNo.Text; if (DevNo == "") DevNo = "0";

            if (TextDevName.Text == "") Msg.Text = "<script>alert('您未填設備名稱！');</script>";
            else if (SelMaintainor.SelectedValue == "") Msg.Text = "<script>alert('您未填維護人員！');</script>";
            else if (SelDevType.SelectedValue != "網路設備" & SelDevStatus.SelectedValue == "已報廢" & GetValueWithDevNo("IDMS", "select count(*) from [作業主機] where [設備編號]=@DevNo" , DevNo) != "0")//Revise
                Msg.Text = "<script>alert('請先刪除或轉移安裝的系統作業，再將狀態設為已報廢！');</script>";
            else if (SelDevStatus.SelectedValue=="已報廢" & ExistsCheck("select [下游編號] from [接電] where [上游編號]= @DeviceNumber" , TextDevNo.Text))//Revise
                Msg.Text = "<script>alert('請先移除或轉移下游接電迴路，再將狀態設為已報廢！');</script>";
            else if (SelDevStatus.SelectedValue == "已報廢" & ExistsCheck("select [下游編號] from [接網] where [上游編號]= @DeviceNumber" , TextDevNo.Text))//Revise
                Msg.Text = "<script>alert('請先移除或轉移下游接網迴路，再將狀態設為已報廢！');</script>";
            else if (SelDevStatus.SelectedValue == "已報廢" & int.Parse(TextDevNo.Text) <= 0)
                Msg.Text = "<script>alert('系統專用，禁止將狀態設為已報廢！');</script>";
            else
            {
                LblUpdateDT.Text = DateTime.Now.ToString("yyyy/MM/dd HH:mm");
                LblUpdater.Text = UserName;
                if (TextVoltage.Text == "") TextVoltage.Text = "0";
                if (TextSpecCurrent.Text == "") TextSpecCurrent.Text = "0";
                BlankProp(); //若設備財產等於秘總財產，清空之

                List<SqlParameter> pars = new List<SqlParameter>();
                pars.Add(new SqlParameter("@TextDevNo", TextDevNo.Text));
                DevSQL = GetUpdateDev(int.Parse(TextDevNo.Text), "SQL", pars);
                LifeSQL = "修改 [" + SelDevType.SelectedValue + "].[" + SelDevKind.SelectedValue + "].[" + TextDevName.Text + "] 的實體設備 ：  " + GetUpdateDev(int.Parse(TextDevNo.Text), "Life");
                
                if (DevSQL != "") ExecDbSQL("Update [實體設備] set " + DevSQL + " where [設備編號]=@TextDevNo", pars);        

                if (SelDevType.SelectedValue != "網路設備")
                {
                    if (DevSQL != "") InsLifeLog(LifeSQL);
                }
                else//網路設備
                {
                    string UpdateApLife = GetUpdateAp(int.Parse(TextDevNo.Text), "Life");
                    if (UpdateApLife != "")//有動到主機 作業主機上也要更新
                    {
                        List<SqlParameter> pars2 = new List<SqlParameter>();
                        pars2.Add(new SqlParameter("@TextDevNo", TextDevNo.Text));

                        LifeSQL = LifeSQL + "  ；  修改 [" + SelDevType.SelectedValue + "].[" + SelDevKind.SelectedValue + "].[" + TextDevName.Text + "] 的作業主機 ： " + UpdateApLife;
                        ExecDbSQL("Update [作業主機] set " + GetUpdateAp(int.Parse(TextDevNo.Text), "SQL", pars2) + " where [設備編號]=@TextDevNo", pars2);
                    }
                    else//無更新主機的話
                    {
                        if (GetValueWithDevNo("IDMS", "SELECT [設備編號] FROM [作業主機] WHERE [設備編號] = @DevNo" , TextDevNo.Text) == "")//Revise
                        {
                            string ApNo = GetPKNo("作業編號", "作業主機").ToString();
                            List<SqlParameter> pars3 = new List<SqlParameter>();
                            string SQL = GetInsApSQL(ApNo, TextDevNo.Text, pars3);
                            string SQL_log = GetInsApSQL(ApNo, TextDevNo.Text);
                            ExecDbSQL(SQL, pars3);
                            LifeSQL = LifeSQL + "　；　" + SQL_log;
                        }
                    }
                    InsLifeLog(LifeSQL);
                }

                if (SelDevStatus.SelectedValue == "已報廢")
                {
					DeleteDevNo("delete from [接電] where [上游編號]= @DevNo or [下游編號]= @DevNo", TextDevNo.Text);//Revise
					DeleteDevNo("delete from [接網] where [上游編號]= @DevNo or [下游編號]= @DevNo", TextDevNo.Text);//Revise
					DeleteDevNo("delete from [作業主機] where [設備編號]= @DevNo", TextDevNo.Text);//Revise
                }
                //Test.Text = "sscript>alert('更新資料[" + TextDevName.Text + "] 完成！\\n\\n" + GetChecks(int.Parse(TextDevNo.Text)) + "\\n" + IsSame + "');/script>"; 
                Msg.Text = "<script>alert('更新資料[" + TextDevName.Text + "] 完成！\\n\\n" + GetChecks(int.Parse(TextDevNo.Text)) + "\\n" + IsSame + "');</script>";
                ReadAll(int.Parse(TextDevNo.Text));
            }
        }
        else {
            //Test.Text = outans;
            Msg.Text = "<script>alert('"+outans+ "\\n您不是原維護人員，沒有修改 [" + TextDevName.Text + "] 的權限！\\n您可將維護人員設為自己或自己所屬的小組，但不能改為別人！若有需要，可請當事者將這筆資料設為自己。');</script>";
        }

        Page.Controls.Add(Msg);
    }

    protected string GetConfig(string SQL) //讀取某系統設定值
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

    protected string GetConfig(string SQL, string key, string value) //讀取某系統設定值
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
                if (ColName != "定位編號")
                {
                    if (Source == "") Source = "(空白)";
                    if (Target == "") Target = "(空白)";
                    SQL = SQL + ",[" + ColName + "]：" + Source + " -> " + Target;
                }
                else
                {
                    List<SqlParameter> pars = new List<SqlParameter>();
                    
                    SQL = SQL + ",[" + ColName + "]：" + Source + "(" + GetConfig("select [區域名稱]+[定位名稱] as [放置地點] from [定位設定] where [定位編號]=@Source", "Source", Source)
                        + ") -> " + Target + "(" + GetConfig("select [區域名稱]+[定位名稱] as [放置地點] from [定位設定] where [定位編號]=@Target", "Target", Target) + ")";
                }
            }
        }
        return (SQL);
    }

    protected string GetUpdateCol(string ColName, string Source, string Target, string Kind, string SQLorLife, List<SqlParameter> pars) //取得單一欄位修改資料的語法
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
                    case "string":case "date":case "datetime": 
                        SQL = SQL + ",[" + ColName + "]=@" + Convert.ToString(pars.Count);
                        pars.Add(new SqlParameter("@"+Convert.ToString(pars.Count), Target));
                        break;
                    case "integer":case "money":                        
                        SQL = SQL + ",[" + ColName + "]=@" + Convert.ToString(pars.Count); 
                        pars.Add(new SqlParameter("@"+Convert.ToString(pars.Count), SqlDbType.Int));
                        pars.Last().Value = int.Parse(Target);
                        break;
                    case "null": 
                        SQL = SQL + ",[" + ColName + "]=" + "null"; 
						break;
                    default:
                        SQL = SQL + ",[" + ColName + "]=" + "@" + Convert.ToString(pars.Count); 
						pars.Add(new SqlParameter("@" + Convert.ToString(pars.Count), Target));				
						break;
                }
            }
            /*
            else if (SQLorLife == "Life")
            {
                if (ColName != "定位編號")
                {
                    if (Source == "") Source = "(空白)";
                    if (Target == "") Target = "(空白)";
                    SQL = SQL + ",[" + ColName + "]：" + Source + " -> " + Target;
                }
                else
                {
                    SQL = SQL + ",[" + ColName + "]：" + Source + "(" + GetConfig("select [區域名稱]+[定位名稱] as [放置地點] from [定位設定] where [定位編號]=" + Source)
                        + ") -> " + Target + "(" + GetConfig("select [區域名稱]+[定位名稱] as [放置地點] from [定位設定] where [定位編號]=" + Target) + ")";
                }
            }
            */
        }
        return (SQL);
    }

    protected string GetUpdateDev(int DevNo, string SQLorLife) //取得修改資料的語法
    {
        string SQL = "";
        
        string PointerNo = SelPointer.SelectedValue;
        if (TextPointerNo.Text != "") PointerNo = TextPointerNo.Text;

        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
        Conn.Open();
        SqlCommand cmd = new SqlCommand("select * from [實體設備] where [設備編號]= @DeviceNumber", Conn);//Revise
		//cmd.Parameters.AddWithValue("@DeviceNumber", DevNo);//Revise
		cmd.Parameters.Add("@DeviceNumber", SqlDbType.Int).Value = DevNo;//Revise
        SqlDataReader dr = null;
        dr = cmd.ExecuteReader();
        if (dr.Read())
        {
            string last_edit;DateTime hd;
            if(DateTime.TryParse(dr["修改日期"].ToString(), out hd)){
                last_edit = DateTime.Parse(dr["修改日期"].ToString()).ToString("yyyy/MM/dd HH:mm");
            }
            else{last_edit = null;}

            SQL = GetUpdateCol("資產編號", dr["資產編號"].ToString(), SelAssets.SelectedValue, "string", SQLorLife)
                + GetUpdateCol("設備型態", dr["設備型態"].ToString(), SelDevType.SelectedValue, "string", SQLorLife)
                + GetUpdateCol("設備種類", dr["設備種類"].ToString(), SelDevKind.SelectedValue, "string", SQLorLife)
                + GetUpdateCol("設備名稱", dr["設備名稱"].ToString(), TextDevName.Text, "string", SQLorLife)
                + GetUpdateCol("設備用途", dr["設備用途"].ToString(), TextPurpose.Text, "string", SQLorLife)
                + GetUpdateCol("財產編號A", dr["財產編號A"].ToString(), TextPropNoA.Text, "string", SQLorLife)
                + GetUpdateCol("財產編號B", dr["財產編號B"].ToString(), TextPropNoB.Text, "string", SQLorLife)
                + GetUpdateCol("廠牌", dr["廠牌"].ToString(), TextBrand.Text, "string", SQLorLife)
                + GetUpdateCol("型式", dr["型式"].ToString(), TextStyle.Text, "string", SQLorLife)
                + GetUpdateCol("保管人員", dr["保管人員"].ToString(), TextKeeper.Text, "string", SQLorLife)
                + GetUpdateCol("定位編號", dr["定位編號"].ToString(), PointerNo, "integer", SQLorLife)
                + GetUpdateCol("高度", dr["高度"].ToString(), SelRU.SelectedValue, "integer", SQLorLife)
                + GetUpdateCol("空間大小", dr["空間大小"].ToString(), SelSpace.SelectedValue, "integer", SQLorLife)
                + GetUpdateCol("取得來源", dr["取得來源"].ToString(), TextSource.Text, "string", SQLorLife)
                + GetUpdateCol("規格", dr["規格"].ToString(), TextSpec.Text, "string", SQLorLife)
                + GetUpdateCol("硬體序號", dr["硬體序號"].ToString(), TextSerial.Text, "string", SQLorLife)
                + GetUpdateCol("iLoIP", dr["iLoIP"].ToString(), TextiLoIP.Text, "string", SQLorLife)
                + GetUpdateCol("用電電壓", dr["用電電壓"].ToString(), TextVoltage.Text, "integer", SQLorLife)
                + GetUpdateCol("額定電流", dr["額定電流"].ToString(), TextSpecCurrent.Text, "float", SQLorLife)
                + GetUpdateCol("維護廠商", dr["維護廠商"].ToString(), TextRepair.Text, "string", SQLorLife)
                + GetUpdateCol("維護人員", dr["維護人員"].ToString(), SelMaintainor.SelectedValue, "string", SQLorLife)
                + GetUpdateCol("設備狀態", dr["設備狀態"].ToString(), SelDevStatus.SelectedValue, "string", SQLorLife)
                + GetUpdateCol("關機順序", dr["關機順序"].ToString(), SelPowerOff.SelectedValue, "string", SQLorLife)
                + GetUpdateCol("修改日期", last_edit, DateTime.Now.ToString("yyyy/MM/dd HH:mm"), "datetime", SQLorLife)
                + GetUpdateCol("修改人員", dr["修改人員"].ToString(), Session["UserName"].ToString(), "string", SQLorLife)
                + GetUpdateCol("備註說明", dr["備註說明"].ToString(), TextMemo.Text, "string", SQLorLife);

            if (SQL != "") SQL = SQL.Substring(1);
        }
        cmd.Cancel(); cmd.Dispose(); dr.Close(); Conn.Close(); Conn.Dispose();

        return (SQL);
    }

    protected string GetUpdateDev(int DevNo, string SQLorLife, List<SqlParameter> pars) //取得修改資料的語法
    {
        string SQL = "";
        
        string PointerNo = SelPointer.SelectedValue;
        if (TextPointerNo.Text != "") PointerNo = TextPointerNo.Text;

        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
        Conn.Open();
        SqlCommand cmd = new SqlCommand("select * from [實體設備] where [設備編號]= @DeviceNumber", Conn);//Revise
		//cmd.Parameters.AddWithValue("@DeviceNumber", DevNo);//Revise
		cmd.Parameters.Add("@DeviceNumber", SqlDbType.Int).Value = DevNo;//Revise
        SqlDataReader dr = null;
        dr = cmd.ExecuteReader();
        if (dr.Read())
        {
            string last_edit;DateTime hd;
            if(DateTime.TryParse(dr["修改日期"].ToString(), out hd)){
                last_edit = DateTime.Parse(dr["修改日期"].ToString()).ToString("yyyy/MM/dd HH:mm");
            }
            else{last_edit = null;}

            SQL = GetUpdateCol("資產編號", dr["資產編號"].ToString(), SelAssets.SelectedValue, "string", SQLorLife,pars)
                + GetUpdateCol("設備型態", dr["設備型態"].ToString(), SelDevType.SelectedValue, "string", SQLorLife,pars)
                + GetUpdateCol("設備種類", dr["設備種類"].ToString(), SelDevKind.SelectedValue, "string", SQLorLife,pars)
                + GetUpdateCol("設備名稱", dr["設備名稱"].ToString(), TextDevName.Text, "string", SQLorLife, pars)
                + GetUpdateCol("設備用途", dr["設備用途"].ToString(), TextPurpose.Text, "string", SQLorLife, pars)
                + GetUpdateCol("財產編號A", dr["財產編號A"].ToString(), TextPropNoA.Text, "string", SQLorLife, pars)
                + GetUpdateCol("財產編號B", dr["財產編號B"].ToString(), TextPropNoB.Text, "string", SQLorLife, pars)
                + GetUpdateCol("廠牌", dr["廠牌"].ToString(), TextBrand.Text, "string", SQLorLife, pars)
                + GetUpdateCol("型式", dr["型式"].ToString(), TextStyle.Text, "string", SQLorLife, pars)
                + GetUpdateCol("保管人員", dr["保管人員"].ToString(), TextKeeper.Text, "string", SQLorLife, pars)
                + GetUpdateCol("定位編號", dr["定位編號"].ToString(), PointerNo, "integer", SQLorLife, pars)
                + GetUpdateCol("高度", dr["高度"].ToString(), SelRU.SelectedValue, "integer", SQLorLife, pars)
                + GetUpdateCol("空間大小", dr["空間大小"].ToString(), SelSpace.SelectedValue, "integer", SQLorLife, pars)
                + GetUpdateCol("取得來源", dr["取得來源"].ToString(), TextSource.Text, "string", SQLorLife, pars)
                + GetUpdateCol("規格", dr["規格"].ToString(), TextSpec.Text, "string", SQLorLife, pars)
                + GetUpdateCol("硬體序號", dr["硬體序號"].ToString(), TextSerial.Text, "string", SQLorLife, pars)
                + GetUpdateCol("iLoIP", dr["iLoIP"].ToString(), TextiLoIP.Text, "string", SQLorLife, pars)
                + GetUpdateCol("用電電壓", dr["用電電壓"].ToString(), TextVoltage.Text, "integer", SQLorLife, pars)
                + GetUpdateCol("額定電流", dr["額定電流"].ToString(), TextSpecCurrent.Text, "float", SQLorLife, pars)
                + GetUpdateCol("維護廠商", dr["維護廠商"].ToString(), TextRepair.Text, "string", SQLorLife, pars)
                + GetUpdateCol("維護人員", dr["維護人員"].ToString(), SelMaintainor.SelectedValue, "string", SQLorLife, pars)
                + GetUpdateCol("設備狀態", dr["設備狀態"].ToString(), SelDevStatus.SelectedValue, "string", SQLorLife, pars)
                + GetUpdateCol("關機順序", dr["關機順序"].ToString(), SelPowerOff.SelectedValue, "string", SQLorLife, pars)
                + GetUpdateCol("修改日期", last_edit, DateTime.Now.ToString("yyyy/MM/dd HH:mm"), "datetime", SQLorLife, pars)
                + GetUpdateCol("修改人員", dr["修改人員"].ToString(), Session["UserName"].ToString(), "string", SQLorLife, pars)
                + GetUpdateCol("備註說明", dr["備註說明"].ToString(), TextMemo.Text, "string", SQLorLife, pars);

            if (SQL != "") SQL = SQL.Substring(1);
        }
        cmd.Cancel(); cmd.Dispose(); dr.Close(); Conn.Close(); Conn.Dispose();

        return (SQL);
    }

    protected string GetUpdateAp(int DevNo, string SQLorLife) //取得修改資料的語法
    {
        string SQL = "";
        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
        Conn.Open();
        SqlCommand cmd = new SqlCommand("select * from [作業主機] where [設備編號]= @DeviceNumber", Conn);
		//cmd.Parameters.AddWithValue("@DeviceNumber", DevNo);//Revise
		cmd.Parameters.Add("@DeviceNumber", SqlDbType.Int).Value = DevNo;//Revise
		
        SqlDataReader dr = null;
        dr = cmd.ExecuteReader();
        if (dr.Read())
        {
            SQL = GetUpdateCol("主機名稱", dr["主機名稱"].ToString(), TextHostName.Text, "string", SQLorLife)
                + GetUpdateCol("IP", dr["IP"].ToString(), TextIP.Text, "string", SQLorLife)
                + GetUpdateCol("監控IP", dr["監控IP"].ToString(), TextMosIP.Text, "string", SQLorLife)
                + GetUpdateCol("系統編號", dr["系統編號"].ToString(),"9", "integer", SQLorLife)
                + GetUpdateCol("安內外", dr["安內外"].ToString(), SelSaveIO.SelectedValue, "string", SQLorLife)
                + GetUpdateCol("緊急程度", dr["緊急程度"].ToString(), SelOnCall.SelectedValue, "string", SQLorLife);

            if (SQL != "") SQL = SQL.Substring(1);
        }
        cmd.Cancel(); cmd.Dispose(); dr.Close(); Conn.Close(); Conn.Dispose();

        return (SQL);
    }

    protected string GetUpdateAp(int DevNo, string SQLorLife, List<SqlParameter> pars) //取得修改資料的語法
    {
        string SQL = "";
        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
        Conn.Open();
        SqlCommand cmd = new SqlCommand("select * from [作業主機] where [設備編號]= @DeviceNumber", Conn);
		//cmd.Parameters.AddWithValue("@DeviceNumber", DevNo);//Revise
		cmd.Parameters.Add("@DeviceNumber", SqlDbType.Int).Value = DevNo;//Revise
		
        SqlDataReader dr = null;
        dr = cmd.ExecuteReader();
        if (dr.Read())
        {
            SQL = GetUpdateCol("主機名稱", dr["主機名稱"].ToString(), TextHostName.Text, "string", SQLorLife, pars)
                + GetUpdateCol("IP", dr["IP"].ToString(), TextIP.Text, "string", SQLorLife, pars)
                + GetUpdateCol("監控IP", dr["監控IP"].ToString(), TextMosIP.Text, "string", SQLorLife, pars)
                + GetUpdateCol("系統編號", dr["系統編號"].ToString(),"9", "integer", SQLorLife, pars)
                + GetUpdateCol("安內外", dr["安內外"].ToString(), SelSaveIO.SelectedValue, "string", SQLorLife, pars)
                + GetUpdateCol("緊急程度", dr["緊急程度"].ToString(), SelOnCall.SelectedValue, "string", SQLorLife, pars);

            if (SQL != "") SQL = SQL.Substring(1);
        }
        cmd.Cancel(); cmd.Dispose(); dr.Close(); Conn.Close(); Conn.Dispose();

        return (SQL);
    }




    //因為引用該function數量過多，需要重新改寫，大概要替換掉15個
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

	//Revise ExecDbSQL版本A，替換掉7個
    protected void DeleteDevNo(string SQL, string DevNo) //執行資料庫異動
    {
        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
        Conn.Open();
        SqlCommand cmd = new SqlCommand(SQL, Conn);
		cmd.Parameters.Add("@DevNo", SqlDbType.Int).Value = DevNo;//替換掉@DevNo

        cmd.ExecuteNonQuery();
        cmd.Cancel(); cmd.Dispose(); Conn.Close(); Conn.Dispose();
    }
    protected void InsData() //新增資料
    {
        string DevNo = GetPKNo("設備編號", "實體設備").ToString();
        string ApNo = GetPKNo("作業編號", "作業主機").ToString();
        string LifeSQL = "";

        TextDevNo.Text = DevNo; //新增完成後，要賦予新取得之設備編號

        if (TextVoltage.Text == "") TextVoltage.Text = "0";

        List<SqlParameter> pars = new List<SqlParameter>();
        string SQL = GetInsDevSQL(DevNo, pars); LifeSQL = GetInsDevSQL(DevNo);//for life
        
        ExecDbSQL(SQL, pars);
        if (SelDevType.SelectedValue != "網路設備") InsLifeLog(LifeSQL);
        else
        {
            List<SqlParameter> pars2 = new List<SqlParameter>();
            SQL = GetInsApSQL(ApNo, DevNo);//作業主機也要新增一筆資料 for life
            string SQL2 = GetInsApSQL(ApNo, DevNo, pars2);//for db
            
            ExecDbSQL(SQL2,pars2);
            InsLifeLog(LifeSQL + "；\n" + SQL);
        }
    }

    protected string GetInsDevSQL(string DevNo) //取得新增資料的語法
    {
        string PointerNo = SelPointer.SelectedValue;
        if (TextPointerNo.Text != "") PointerNo = TextPointerNo.Text;

        return ("新增至 [實體設備] ("
            + DevNo + ",'"
            + SelAssets.SelectedValue + "','"
            + SelDevType.SelectedValue + "','"
            + SelDevKind.SelectedValue + "','"
            + TextDevName.Text + "','"
            + TextPurpose.Text + "','"
            + TextPropNoA.Text + "','"
            + TextPropNoB.Text + "','','','"
            + TextBrand.Text + "','"
            + TextStyle.Text + "','',0,'2000/01/01',0,'"
            + TextKeeper.Text + "',"
            + PointerNo + ","
            + SelRU.SelectedValue + ","
            + SelSpace.SelectedValue + ",'"
            + TextSource.Text + "','"
            + TextSpec.Text + "','"
            + TextSerial.Text + "','"
            + TextiLoIP.Text + "',"
            + TextVoltage.Text + ","
            + TextSpecCurrent.Text + ",'"
            + TextRepair.Text + "','"
            + SelMaintainor.SelectedValue + "','"
            + SelDevStatus.SelectedValue + "','"
            + SelPowerOff.SelectedValue + "','"
            + LblCreateDT.Text + "','"
            + LblCreater.Text + "','"
            + LblUpdateDT.Text + "','"
            + LblUpdater.Text + "','"
            + TextMemo.Text + "'"
            + ")");
    }

    protected string GetInsDevSQL(string DevNo, List<SqlParameter> pars) //取得新增資料的語法
    {
        string PointerNo = SelPointer.SelectedValue;
        if (TextPointerNo.Text != "") PointerNo = TextPointerNo.Text;

        pars.Add(new SqlParameter("@DevNo", DevNo));
        pars.Add(new SqlParameter("@SelAssets",SelAssets.SelectedValue));
        pars.Add(new SqlParameter("@SelDevType",SelDevType.SelectedValue));
        pars.Add(new SqlParameter("@SelDevKind",SelDevKind.SelectedValue));
        pars.Add(new SqlParameter("@TextDevName",TextDevName.Text));
        pars.Add(new SqlParameter("@TextPurpose",TextPurpose.Text));
        pars.Add(new SqlParameter("@TextPropNoA",TextPropNoA.Text));
        pars.Add(new SqlParameter("@TextPropNoB",TextPropNoB.Text));
        pars.Add(new SqlParameter("@TextBrand",TextBrand.Text));
        pars.Add(new SqlParameter("@TextStyle",TextStyle.Text));
        pars.Add(new SqlParameter("@TextKeeper",TextKeeper.Text));
        pars.Add(new SqlParameter("@PointerNo",PointerNo));
        pars.Add(new SqlParameter("@SelRU",SelRU.SelectedValue));
        pars.Add(new SqlParameter("@SelSpace",SelSpace.SelectedValue));
        pars.Add(new SqlParameter("@TextSource",TextSource.Text));
        pars.Add(new SqlParameter("@TextSpec",TextSpec.Text ));
        pars.Add(new SqlParameter("@TextSerial",TextSerial.Text));
        pars.Add(new SqlParameter("@TextiLoIP",TextiLoIP.Text));
        pars.Add(new SqlParameter("@TextVoltage",TextVoltage.Text));
        pars.Add(new SqlParameter("@TextSpecCurrent",TextSpecCurrent.Text));
        pars.Add(new SqlParameter("@TextRepair",TextRepair.Text));
        pars.Add(new SqlParameter("@SelMaintainor",SelMaintainor.SelectedValue));
        pars.Add(new SqlParameter("@SelDevStatus",SelDevStatus.SelectedValue ));
        pars.Add(new SqlParameter("@SelPowerOff",SelPowerOff.SelectedValue));
        pars.Add(new SqlParameter("@LblCreateDT",LblCreateDT.Text));
        pars.Add(new SqlParameter("@LblCreater",LblCreater.Text));
        pars.Add(new SqlParameter("@LblUpdateDT",LblUpdateDT.Text));
        pars.Add(new SqlParameter("@LblUpdater",LblUpdater.Text));
        pars.Add(new SqlParameter("@TextMemo",TextMemo.Text));



        return (@"INSERT INTO [實體設備] VALUES(
            @DevNo ,
            @SelAssets,
            @SelDevType,
            @SelDevKind,
            @TextDevName,
            @TextPurpose,
            @TextPropNoA,
            @TextPropNoB,'','',
            @TextBrand,
            @TextStyle,'',0,'2000/01/01',0,
            @TextKeeper,
            @PointerNo,
            @SelRU,
            @SelSpace,
            @TextSource,
            @TextSpec,
            @TextSerial,
            @TextiLoIP,
            @TextVoltage,
            @TextSpecCurrent,
            @TextRepair,
            @SelMaintainor,
            @SelDevStatus,
            @SelPowerOff,
            @LblCreateDT,
            @LblCreater,
            @LblUpdateDT,
            @LblUpdater,
            @TextMemo
            )");
    }

    protected string GetInsApSQL(string ApNo, string DevNo) //取
    {
        return ("新增至 [作業主機] ("
            + ApNo + ","
            + DevNo + ",'"
            + TextHostName.Text + "',"
            + "9,'"
            + "','"
            + "','"
            + "','"
            + TextIP.Text + "','"
            + TextMosIP.Text + "','"
            + SelSaveIO.SelectedValue + "','"
            + SelOnCall.SelectedValue + "','"
            + "','"
            + "','"
            + "','"
            + LblCreateDT.Text + "','"
            + LblCreater.Text + "','"
            + LblUpdateDT.Text + "','"
            + LblUpdater.Text + "','"
            + "')");
    }

    protected string GetInsApSQL(string ApNo, string DevNo, List<SqlParameter> pars) //取
    {
        pars.Add(new SqlParameter("@ApNo",ApNo));
        pars.Add(new SqlParameter("@DevNo",DevNo));
        pars.Add(new SqlParameter("@TextHostName",TextHostName.Text));
        pars.Add(new SqlParameter("@TextIP",TextIP.Text));
        pars.Add(new SqlParameter("@TextMosIP",TextMosIP.Text));
        pars.Add(new SqlParameter("@SelSaveIO",SelSaveIO.SelectedValue));
        pars.Add(new SqlParameter("@SelOnCall",SelOnCall.SelectedValue));    
        pars.Add(new SqlParameter("@LblCreateDT",LblCreateDT.Text));    ;
        pars.Add(new SqlParameter("@LblCreater",LblCreater.Text));
        pars.Add(new SqlParameter("@LblUpdateDT",LblUpdateDT.Text));
        pars.Add(new SqlParameter("@LblUpdater",LblUpdater.Text));
        

        return (@"INSERT INTO [作業主機] VALUES(
            @ApNo ,
            @DevNo ,
            @TextHostName,
            9,
            '',
            '',
            '',
            @TextIP,
            @TextMosIP,
            @SelSaveIO,
            @SelOnCall,
            '',
            '',
            '',
            @LblCreateDT,
            @LblCreater,
            @LblUpdateDT,
            @LblUpdater,
            ''
            )");
    }

    protected void SelPointer_SelectedIndexChanged(object sender, EventArgs e)  //定位編號
    {
        DropDownList Sel = (DropDownList)form1.FindControl("SelPointer");    //放入Table物件後，必須用FindControl才能取值
        TextPointerNo.Text="";
    }

    protected void SelRepair_SelectedIndexChanged(object sender, EventArgs e)  //維護廠商
    {
        DropDownList Sel = (DropDownList)form1.FindControl("SelRepair");    //放入Table物件後，必須用FindControl才能取值
        TextRepair.Text = Sel.SelectedValue;
    }

    protected void GetRepair()   //顯示維護廠商名
    {
        DropDownList Sel = (DropDownList)form1.FindControl("SelRepair");     //放入Table物件後，必須用FindControl才能取值
        Label Lbl = (Label)form1.FindControl("LblRepair");

        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
        Conn.Open();
        //SqlCommand cmd = new SqlCommand("select [Config] FROM [Config] WHERE [Kind]='維護廠商' and [Item]='" + Sel.SelectedValue + "'", Conn);
        SqlCommand cmd = new SqlCommand("select [Config] FROM [Config] WHERE [Kind]='維護廠商' and [Item]= @Sel", Conn);
        cmd.Parameters.AddWithValue("@Sel", Sel.SelectedValue);
        SqlDataReader dr = cmd.ExecuteReader();

        if (dr.Read()) Lbl.Text = dr[0].ToString();
        else Lbl.Text = "";

        cmd.Cancel(); cmd.Dispose(); dr.Close(); Conn.Close(); Conn.Dispose();
    }

    protected void GetMaintainor()   //顯示維護人員清單
    {
        DropDownList Sel = (DropDownList)form1.FindControl("SelMaintainor");     //放入Table物件後，必須用FindControl才能取值
        Label Lbl = (Label)form1.FindControl("LblMaintainor");

        SqlCommand sqlQuery = new SqlCommand();
        sqlQuery.CommandText = "select [Item] from [Config] where [Kind]= @Kind order by Mark";//Revise
		sqlQuery.Parameters.Add("@Kind", SqlDbType.VarChar).Value = Sel.SelectedValue;//Revise
        DataSet ds = RunQuery(sqlQuery);

        LblMaintainor.Text = "";
        if (ds.Tables.Count > 0)
        {
            foreach (DataRow row in ds.Tables[0].Rows)
            {
                Lbl.Text = Lbl.Text + " " + row[0].ToString();
            }
        }
        else
        {
            Lbl.Text = Sel.SelectedValue;
        }

        sqlQuery.Cancel(); ds.Dispose();
    }

    protected void SelSource_SelectedIndexChanged(object sender, EventArgs e)  //取得來源
    {
        DropDownList Sel = (DropDownList)form1.FindControl("SelSource");    //放入Table物件後，必須用FindControl才能取值
        TextSource.Text = Sel.SelectedValue;
    }

    protected Boolean RightCheck() //是否有權修改資料
    {
        string UserID = Session["UserID"].ToString().ToLower();
        string UserName = Session["UserName"]==null?"":Session["UserName"].ToString();   //登入的UserName
        string UnitName = Session["UnitName"]==null?"":Session["UnitName"].ToString();   //登入的UnitName
        string Hw = SelMaintainor.SelectedValue; //設備維護人員
        string Dno = TextDevNo.Text; if (Dno == "") Dno = "0";
        string Older = GetValueWithDevNo("IDMS", "select [維護人員] from [實體設備] where [設備編號]= @DevNo" , Dno);//Revise
        //Test.Text += "IN";
        //if (UserID != "operator" && (InGroup(UserName, Hw) || InGroup(UserName, Older) || UnitName == "系統管控科" && GetDuty(Hw) == UnitName || UnitName == "作業管理科" || InGroup(UserName, "軟體小組") && SelDevType.SelectedValue != "系統設備"))
        


        if (UserID != "operator" && (InGroup(UserName, Hw) || InGroup(UserName, Older) || UnitName == "系統管控科" || UnitName == "作業管理科" || InGroup(UserName, "軟體小組") && SelDevType.SelectedValue != "系統設備"))
        {
            if(SelDevKind.SelectedValue != "虛擬主機" || InGroup(UserName,"SSM小組")) return (true);   //課內只有SSM小組可修改虛擬主機，課外8888則可
            else return (false); 
        }
        
        
        return (false);
    }

    protected Boolean RightCheck(ref string answer) //是否有權修改資料//外串
    {
        string UserID = Session["UserID"].ToString().ToLower();
        string UserName = Session["UserName"]==null?"":Session["UserName"].ToString();   //登入的UserName
        string UnitName = Session["UnitName"]==null?"":Session["UnitName"].ToString();   //登入的UnitName
        string Hw = SelMaintainor.SelectedValue; //設備維護人員
        string Dno = TextDevNo.Text; if (Dno == "") Dno = "0";
        string Older = GetValueWithDevNo("IDMS", "select [維護人員] from [實體設備] where [設備編號]= @DevNo" , Dno);//Revise
        //Test.Text += "IN";
        //if (UserID != "operator" && (InGroup(UserName, Hw) || InGroup(UserName, Older) || UnitName == "系統管控科" && GetDuty(Hw) == UnitName || UnitName == "作業管理科" || InGroup(UserName, "軟體小組") && SelDevType.SelectedValue != "系統設備"))
        


        if ( UserID != "operator" && (InGroup(UserName, Hw) || InGroup(UserName, Older) || UnitName == "系統管控科" || UnitName == "作業管理科" || InGroup(UserName, "軟體小組") && SelDevType.SelectedValue != "系統設備"))
        {
            if(SelDevKind.SelectedValue == "個人電腦"){
                if(InGroup(UserName, "個人電腦小組") || UnitName == "系統管控科" || UnitName == "作業管理科")return true;
                else {
                    answer += "「非個人電腦小組成員，無權限編輯個人電腦，請聯繫個人電腦小組或系統科。」";
             
                    return false;
                }
            }

            if(SelDevKind.SelectedValue != "虛擬主機" || InGroup(UserName,"SSM小組")) return (true);            
            else return (false); 
        }
        
        
        return (false);
    }
    protected Boolean MaintainChange(){
        string UserName = Session["UserName"]==null?"":Session["UserName"].ToString();   //登入的UnitName
        string Hw = SelMaintainor.SelectedValue; //設備維護人員
        return InGroup(UserName, Hw);
    }

	//Revise總計19個地方引用該函數，一個個確認太麻煩了，需要全部改掉
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

    protected string GetValue(string DB, string SQL, List<SqlParameter> pars)   //取得單一資料
    {
        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings[DB + "ConnectionString"].ConnectionString);
        Conn.Open();
        SqlCommand cmd = new SqlCommand(SQL, Conn);
        cmd.Parameters.AddRange(pars.ToArray());
        SqlDataReader dr = null;
        dr = cmd.ExecuteReader();
        string cfg = ""; if (dr.Read()) cfg = dr[0].ToString();

        cmd.Cancel(); cmd.Dispose(); dr.Close(); Conn.Close(); Conn.Dispose();

        return (cfg);
    }
	
    protected string GetValueWithDevNo(string DB, string SQL, string DevNo)   //Revise ver.I
    {
        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings[DB + "ConnectionString"].ConnectionString);
        Conn.Open();
        SqlCommand cmd = new SqlCommand(SQL, Conn);
		cmd.Parameters.Add("@DevNo", SqlDbType.Int).Value = DevNo;//Revise
        SqlDataReader dr = null;
        dr = cmd.ExecuteReader();
        string cfg = ""; if (dr.Read()) cfg = dr[0].ToString();

        cmd.Cancel(); cmd.Dispose(); dr.Close(); Conn.Close(); Conn.Dispose();

        return (cfg);
    }
	
	protected string GetValueWithItemName(string DB, string SQL, string ItemName)   //Revise ver.II
    {
        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings[DB + "ConnectionString"].ConnectionString);
        Conn.Open();
        SqlCommand cmd = new SqlCommand(SQL, Conn);
		cmd.Parameters.Add("@ItemName", SqlDbType.VarChar).Value = ItemName;//Revise
        SqlDataReader dr = null;
        dr = cmd.ExecuteReader();
        string cfg = ""; if (dr.Read()) cfg = dr[0].ToString();

        cmd.Cancel(); cmd.Dispose(); dr.Close(); Conn.Close(); Conn.Dispose();

        return (cfg);
    }
	
    protected string GetValueWithSysNo(string DB, string SQL, string SysNo)   //Revise ver.III
    {
        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings[DB + "ConnectionString"].ConnectionString);
        Conn.Open();
        SqlCommand cmd = new SqlCommand(SQL, Conn);
		cmd.Parameters.Add("@SysNo", SqlDbType.Int).Value = SysNo;//Revise
        SqlDataReader dr = null;
        dr = cmd.ExecuteReader();
        string cfg = ""; if (dr.Read()) cfg = dr[0].ToString();

        cmd.Cancel(); cmd.Dispose(); dr.Close(); Conn.Close(); Conn.Dispose();

        return (cfg);
    }
	
    protected Boolean InGroup(string ChkName, string ChkUnit) //檢查ChkName是否為ChkUnit成員或本身
    {
        Boolean TF = false;
        
        if (ChkName == ChkUnit) TF=true;  //是否同名
        else if (ChkUnit == "") TF=false;	//檢查單位必填
        else //是否為成員UN (課別與小組同義)
        {
            SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
            Conn.Open();
            SqlCommand cmd = new SqlCommand("select * from Config where (Kind= @Kind) and Item= @Item", Conn);//Revise
			cmd.Parameters.Add("@Kind", SqlDbType.VarChar).Value = ChkUnit;//Revise
			cmd.Parameters.Add("@Item", SqlDbType.VarChar).Value = ChkName;//Revise
							
            SqlDataReader dr = null;
            dr = cmd.ExecuteReader();
            if (dr.Read()) TF=true;
            cmd.Cancel(); cmd.Dispose(); dr.Close(); Conn.Close(); Conn.Dispose();
        }

        return (TF);
    }

    protected string GetDuty(string Maintainor) //取得使用課別
    {
        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
        Conn.Open();
        string SQL = "(select Kind from Config where Item= @Item and Kind in (select Item from Config where Kind='數值資訊組'))"
            + " UNION (select Config from Config where Item= @Item and kind in (select Item from Config where Kind='維護群組') and Kind<>'數值資訊組')"
            + " UNION (select Item from Config where Item= @Item and Kind='數值資訊組')";//Revise
        SqlCommand cmd = new SqlCommand(SQL, Conn);
		cmd.Parameters.Add("@Item", SqlDbType.VarChar).Value = Maintainor;//Revise
        SqlDataReader dr = null;
        dr = cmd.ExecuteReader();
        string cfg = "";
        if (Maintainor == "") cfg = "";
        else if (dr.Read()) cfg = dr[0].ToString();
        else cfg = "";
        cmd.Cancel(); cmd.Dispose(); dr.Close(); Conn.Close(); Conn.Dispose();

        return (cfg);
    }

    protected void InsLifeLog(string SQL) //寫入生命履歷
    {
        string LifeNo = GetPKNo("履歷編號", "生命履歷").ToString(); //履歷編號
        string TblName = "實體設備";    //表格名稱
        string PKno = TextDevNo.Text;   //主鍵編號
        string Mt = SelMaintainor.SelectedValue;    //維護人員
        string OldMt = GetValueWithDevNo("IDMS", "select [維護人員] from [實體設備] where [設備編號] = @DevNo" , TextDevNo.Text);//Revise    //原負責人
        string OldKeeper = GetValueWithDevNo("IDMS", "select [保管人員] from [View_設備管理] where [設備編號] = @DevNo" , TextDevNo.Text);//Revise    //原保管人
        string UN = Session["UserName"].ToString();   //登入的UserName：異動人員
        string LiftDT = DateTime.Now.ToString("yyyy/MM/dd HH:mm");  //異動日期
        string LifeIP = Request.ServerVariables["REMOTE_ADDR"].ToString();

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

    protected void BtnLife_Click(object sender, EventArgs e) //查詢生命履歷
    {
        Session["LifeSQL"] = "select * from [生命履歷] where [表格名稱]='實體設備' and [主鍵編號]=" + TextDevNo.Text + " order by [履歷編號] desc";
        OpenExecWindow("window.open('../Life/LifeLog.aspx?Search=DevEdit&Tbl=實體設備&PK=" + HttpUtility.UrlEncode(TextDevNo.Text) + "','_self');");//Revise
    }

    protected void BtnIn_Click(object sender, EventArgs e) //列印移入單
    {
        OpenExecWindow("window.open('Enter.aspx?DevNo=" + HttpUtility.UrlEncode(TextDevNo.Text) + "','_self');");//Revise
    }

    protected void LinkPropIn_Click(object sender, EventArgs e)  //帶入
    {
        Literal Msg = new Literal();

        if (TextPropNoA.Text == "00000000000" | TextPropNoA.Text == "") Msg.Text = "<script>alert('您尚未輸入財產編號！');</script>";
        else
        {
            ReadProp();   //lbl讀取秘總財產
            if (lblProp.Text == "") Msg.Text = "<script>alert('秘總系統查無此財產編號(" + TextPropNoA.Text + " " + TextPropNoB.Text + ")！');</script>";
        }
        Page.Controls.Add(Msg);
    }

    protected void LinkKeeperClear_Click(object sender, EventArgs e)  //空白
    {
        SelKeeperUnit.SelectedIndex = 0;
        SelKeeper.Items.Clear();
        TextKeeper.Text = "";
    }

    protected void ReadProp()   //lbl讀取秘總財產
    {
        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
        Conn.Open();
        SqlCommand cmd = new SqlCommand("select * from [財產主檔] where [財產編號A]= @PropertyA and [財產編號B]= @PropertyB ", Conn);
		cmd.Parameters.Add("@PropertyA", SqlDbType.NVarChar).Value = TextPropNoA.Text;//Revise
		cmd.Parameters.Add("@PropertyB", SqlDbType.NVarChar).Value = TextPropNoB.Text;//Revise
        SqlDataReader dr = null;
        dr = cmd.ExecuteReader();
        if (dr.Read())
        {
            lblProp.Text = dr["財產別名"].ToString();
            lblBrand.Text = dr["廠牌"].ToString();
            lblStyle.Text = dr["型式"].ToString();
            lblAmounts.Text = dr["數量"].ToString() + dr["計量單位"].ToString();
            lblPrice.Text = String.Format("{0:F0}", dr["價值"]);
            lblBuyDate.Text = DateTime.Parse(dr["購買日期"].ToString()).ToString("yyyy/MM/dd");
            lblUseLife.Text = (int.Parse(dr["使用年限"].ToString()) / 12).ToString();
            lblKeeper.Text = dr["保管人員"].ToString();            
            lblSpec.Text = dr["財產附屬"].ToString();
            lblStatus.Text = ""; if (dr["無效註記"].ToString() == "N") lblStatus.Text = "已報廢";
        }
        else
        {
            lblProp.Text = "";
            lblBrand.Text = "";
            lblStyle.Text = "";
            lblAmounts.Text = "";
            lblPrice.Text = "";
            lblBuyDate.Text = "";
            lblUseLife.Text = "";
            lblKeeper.Text = "";
            lblUser.Text = "";
            lblSpec.Text = "";
            lblStatus.Text = ""; ;
        }
        cmd.Cancel(); cmd.Dispose(); dr.Close(); Conn.Close(); Conn.Dispose();

        if (TextDevNo.Text != "") lblUser.Text = GetValueWithDevNo("IDMS", "select [使用人員] from [View_通用設備] where [設備編號]= @DevNo", TextDevNo.Text);//Revise
    }

    protected void BtnPing_Click(object sender, EventArgs e) //偵測
    {
        int IsIn = Request.ServerVariables["REMOTE_ADDR"].IndexOf("172.");
        string UrlIn = "http://" + ReadConfig("Primary", "SSM(In)") + "/ping.asp?hostname=" + TextHostName.Text + "&IP=" + TextIP.Text;
        string UrlOut = "/Public/ping.aspx?hostname=" + TextHostName.Text + "&IP=" + TextIP.Text + "&PssmIn=sos-ws01";

        Literal Msg = new Literal();
        if (SelSaveIO.SelectedValue == "I" & IsIn < 0)
            Msg.Text = "<script>alert('[" + TextHostName.Text + "]為[" + SelSaveIO.SelectedValue + "]機器，請至該網段執行偵測！');</script>";
        else
            if (IsIn < 0)
                Msg.Text = "<script>window.open('" + UrlOut + "','_self');</script>";
            else
                Msg.Text = "<script>window.open('" + UrlIn + "','_self');</script>";
        Page.Controls.Add(Msg);
    }

    protected string ReadConfig(string Kind, string Item)
    {
        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["DiaryConnectionString"].ConnectionString);
        Conn.Open();
        SqlCommand cmd = new SqlCommand("select [Config] from [Config] where [Kind]= @Kind and [Item]= @Item", Conn);
		cmd.Parameters.Add("@Kind", SqlDbType.VarChar).Value = Kind;//Revise
		cmd.Parameters.Add("@Item", SqlDbType.VarChar).Value = Item;//Revise
        SqlDataReader dr = null;
        dr = cmd.ExecuteReader();
        string cfg = "";
        if (dr.Read()) cfg = dr[0].ToString();
        cmd.Cancel(); cmd.Dispose(); dr.Close(); Conn.Close(); Conn.Dispose();

        return (cfg);
    }

    protected void LinkAP_Click(object sender, EventArgs e)  //作業編輯
    {
        Literal Msg = new Literal();
        string DevNo = TextDevNo.Text; if (DevNo == "") DevNo = "0";
        string ApNo = GetValueWithDevNo("IDMS", "select [作業編號] from [作業主機] where [設備編號]= @DevNo" , DevNo);//Revise
        string ApCount = GetValueWithDevNo("IDMS", "select Count([作業編號]) from [作業主機] where [設備編號]= @DevNo" , DevNo);//Revise

        if (ApCount == "0") 
        {
            if (DevNo != "" & DevNo != "0") OpenExecWindow("window.open('../AP/ApEdit.aspx?DevNo=" + HttpUtility.UrlEncode(DevNo) + "','_self');");//Revise
            else Msg.Text = "<script>alert('實體設備尚未建檔，無法帶出作業編輯介面！');</script>";
        }
        else if (ApCount == "1") OpenExecWindow("window.open('../AP/ApEdit.aspx?ApNo=" + HttpUtility.UrlEncode(ApNo) + "','_self');");//Revise
        else
        {
            Session["ApSQL"] = "SELECT * FROM [View_作業主機] WHERE [設備編號]=" + DevNo;
            OpenExecWindow("window.open('../AP/Ap.aspx?Search=DevEdit','_self');");
        }

        if (Msg.Text != "") Page.Controls.Add(Msg);
    }

    protected void BtnDevTree_Click(object sender, EventArgs e)  //設備迴路
    {
        Literal Msg = new Literal();
        string DevNo = TextDevNo.Text; if (DevNo == "") DevNo = "0";
        
        if (DevNo != "" & DevNo != "0") OpenExecWindow("window.open('TreeEdit.aspx?DevNo=" + HttpUtility.UrlEncode(DevNo) + "','_self');");//Revise
        else Msg.Text = "<script>alert('實體設備尚未建檔，無法帶出設備迴路編輯介面！');</script>";        

        if (Msg.Text != "") Page.Controls.Add(Msg);
    }

    protected void SelAssets_SelectedIndexChanged(object sender, EventArgs e)   //顯示資產描述
    {
        lblAssets.Text = GetValueWithItemName("IDMS", "select [Memo] from [Config] where ([Kind]='資產清冊' or [Kind]='常用資產') and [Item] = @ItemName" , SelAssets.SelectedValue);//Revise
    }

    protected void SelKeeper_SelectedIndexChanged(object sender, EventArgs e)   //顯示保管人員
    {
        TextKeeper.Text = SelKeeper.SelectedValue;
    }

    protected void BtnDevCheck_Click(object sender, EventArgs e) //查詢機器清查設定介面
    {
        string CheckYear = GetValue("IDMS","select max(清查年度) from [機器清查]");
        OpenExecWindow("window.open('../SOS/ChkEdit.aspx?ChkYear=" + HttpUtility.UrlEncode(CheckYear) + "&DevNo=" + HttpUtility.UrlEncode(TextDevNo.Text) + "','_self');");//Revise
    }

    protected void LinkHideHelp_Click(object sender, EventArgs e)  //顯示隱藏欄位說明
    {
        if (LinkHideHelp.Text == "隱藏說明")
        {
            LinkHideHelp.Text = "顯示說明";
            for (int i = 0; i < Table1.Rows.Count; i++) Table1.Rows[i].Cells[2].Visible = false;
            WriteHideHelp("隱藏說明");
        }
        else
        {
            LinkHideHelp.Text = "隱藏說明";
            for (int i = 0; i < Table1.Rows.Count; i++) Table1.Rows[i].Cells[2].Visible = true;
            WriteHideHelp("");
        }
    }
    /*protected void WriteHideHelp(string txt)  //顯示隱藏欄位說明註記於DB
    {
		InitSession();//檢查session是否存在，無則初始化 //Revise
        string UserName = Session["UserName"].ToString();
        ExecDbSQL("delete from [Config] where [Kind]='隱藏說明' and [Item]='" + UserName + "'");
        if (txt != "") ExecDbSQL("insert into [Config] values('隱藏說明','" + UserName + "','','','')");
    }*/
	protected void WriteHideHelp(string txt){  //顯示隱藏欄位說明註記於DB //Revise重寫了啦!!
		InitSession();
        string userName = Session["UserName"].ToString();
		string sql1 = "delete from [Config] where [Kind]='隱藏說明' and [Item] = @ItemName ";
		string sql2 = "INSERT INTO [Config] values('隱藏說明',@ItemName,'','','')";
		using(SqlConnection conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString))
		using(SqlCommand cmd = new SqlCommand(sql1, conn)){
			cmd.Parameters.Add("@ItemName", SqlDbType.VarChar).Value = userName;
			conn.Open();
			cmd.ExecuteNonQuery();
		}
		if(txt != ""){
			using(SqlConnection conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString))
			using(SqlCommand cmd = new SqlCommand(sql2, conn)){
				cmd.Parameters.Add("@ItemName", SqlDbType.VarChar).Value = userName;
				conn.Open();
				cmd.ExecuteNonQuery();
			}
		}
    }

    protected void ShowHideHelp()  //顯示或隱藏欄位說明
    {
        string un = Session["UserName"]==null?"":Session["UserName"].ToString();
        string UserName = GetValueWithItemName("IDMS", "select [Item] from [Config] where [Kind]='隱藏說明' and [Item] = @ItemName", un);//Revise
        if (UserName != "")
        {
            for (int i = 0; i < Table1.Rows.Count; i++) Table1.Rows[i].Cells[2].Visible = false;
            LinkHideHelp.Text = "顯示說明";
        }
        else
        {
            for (int i = 0; i < Table1.Rows.Count; i++) Table1.Rows[i].Cells[2].Visible = true;
            LinkHideHelp.Text = "隱藏說明";
        }
    }
}