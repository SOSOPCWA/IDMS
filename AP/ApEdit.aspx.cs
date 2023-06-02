using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.Configuration;
using System.Data;
using System.Data.SqlClient;

public partial class Device_ApEdit : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {        
        if (!IsPostBack)
        {
            TextApNo.Text = Request["ApNo"];            
            ReadAPHelp(); //讀取欄位說明             

            //填入預設值
            if (Request["DevNo"] == "" | Request["DevNo"] == null) TextDevNo.Text = "0";
            else TextDevNo.Text = Request["DevNo"];

            //產生各種年月日下拉式選單
            AddSel(SelYYYY, 1999, DateTime.Now.Year); AddSel(SelMM, 1, 12); AddSel(SelDD, 1, 31);
            for (int i = 0; i < SelYYYY.Items.Count; i++) if (SelYYYY.Items[i].Value == DateTime.Now.ToString("yyyy")) SelYYYY.SelectedIndex = i;
            for (int i = 0; i < SelMM.Items.Count; i++) if (SelMM.Items[i].Value == DateTime.Now.ToString("MM")) SelMM.SelectedIndex = i;
            for (int i = 0; i < SelDD.Items.Count; i++) if (SelDD.Items[i].Value == DateTime.Now.ToString("dd")) SelDD.SelectedIndex = i;
        }
    }

    protected void AddSel(DropDownList Sel, int Idx1, int Idx2)   //產生各種年月日下拉式選單
    {
        Sel.Items.Add("");
        for (int i = Idx1; i <= Idx2; i++)
        {
            ListItem Item = new ListItem();

            if (i < 10)
            {
                Item.Text = "0" + i.ToString();
                Item.Value = "0" + i.ToString();
            }
            else
            {
                Item.Text = i.ToString();
                Item.Value = i.ToString();
            }
            Sel.Items.Add(Item);
        }
    }

    protected void Page_PreRenderComplete(object sender, EventArgs e)   //放Page_Load無法顯示
    {
        if (TextApNo.Text == "")    //新增狀態
        {
            BtnEdit.Enabled = false;
            BtnDel.Enabled = false;
        }
        else //編輯狀態
        {
            BtnEdit.Enabled = true;
            BtnDel.Enabled = true;
        }
        if (!IsPostBack & TextApNo.Text != "") ReadAp(int.Parse(TextApNo.Text));  //讀取資料

        ReadDevName(TextDevNo.Text);
        GetMaintainor();    //顯示維護人員清單 
        SelPowerOff_SelectedIndexChanged(null, null);   //關機順序之說明
        if (!IsPostBack)
        {
            SelOS_SelectedIndexChanged(null, null);   //產生平台核心
            ShowHideHelp();  //顯示或隱藏欄位說明
        }
    }

    protected void SelPowerOff_SelectedIndexChanged(object sender, EventArgs e)   //關機順序之說明
    {
        lblPowerOff.Text = GetValue("IDMS", "select [Memo] from [Config] where [Kind]='關機順序' and [Item]='" + SelPowerOff.SelectedValue + "'");
    }

    protected void ReadAp(int ApNo)
    {
        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
        Conn.Open();
        SqlCommand cmd = new SqlCommand("SELECT * FROM [View_作業主機] WHERE [作業編號]=" + TextApNo.Text, Conn);
        SqlDataReader dr = null;
        dr = cmd.ExecuteReader();

        if (dr.Read())
        {
            TextDevNo.Text = dr["設備編號"].ToString();

            TextHostName.Text = dr["主機名稱"].ToString();
            for (int i = 0; i < SelSysNo.Items.Count; i++) if (SelSysNo.Items[i].Value == dr["系統編號"].ToString()) SelSysNo.SelectedIndex = i;
            lblSysAlias.Text = GetSysAlias(dr["系統編號"].ToString(), "");
            lblAssets.Text = dr["系統資產"].ToString();
            TextFunctions.Text = dr["主要功能"].ToString();
            lblSysFun.Text = dr["系統功能"].ToString(); if (lblSysFun.Text == "") lblSysFun.Text = "(無)";

            for (int i = 0; i < SelOS.Items.Count; i++) if (SelOS.Items[i].Value.ToUpper() == dr["作業平台"].ToString().ToUpper()) SelOS.SelectedIndex = i;
            TextKernel.Text = dr["核心版本"].ToString();

            TextIP.Text = dr["IP"].ToString();
            TextMosIP.Text = dr["監控IP"].ToString();
            for (int i = 0; i < SelSaveIO.Items.Count; i++) if (SelSaveIO.Items[i].Value == dr["安內外"].ToString()) SelSaveIO.SelectedIndex = i;
            for (int i = 0; i < SelOnCall.Items.Count; i++) if (SelOnCall.Items[i].Value == dr["緊急程度"].ToString()) SelOnCall.SelectedIndex = i;
            for (int i = 0; i < SelApStatus.Items.Count; i++) if (SelApStatus.Items[i].Value == dr["作業狀態"].ToString()) SelApStatus.SelectedIndex = i;

            string SwUnit = GetSwUnit(dr["維護人員"].ToString());
            for (int i = 0; i < SelSwUnit.Items.Count; i++) if (SelSwUnit.Items[i].Value == SwUnit) SelSwUnit.SelectedIndex = i;
            SelMaintainor.DataBind();  //連帶選項需要觸發
            for (int i = 0; i < SelMaintainor.Items.Count; i++) if (SelMaintainor.Items[i].Value == dr["維護人員"].ToString()) SelMaintainor.SelectedIndex = i;

            for (int i = 0; i < SelYYYY.Items.Count; i++) if (SelYYYY.Items[i].Value == DateTime.Parse(dr["上線日期"].ToString()).ToString("yyyy")) SelYYYY.SelectedIndex = i;
            for (int i = 0; i < SelMM.Items.Count; i++) if (SelMM.Items[i].Value == DateTime.Parse(dr["上線日期"].ToString()).ToString("MM")) SelMM.SelectedIndex = i;
            for (int i = 0; i < SelDD.Items.Count; i++) if (SelDD.Items[i].Value == DateTime.Parse(dr["上線日期"].ToString()).ToString("dd")) SelDD.SelectedIndex = i;

            LblCreateDT.Text = DateTime.Parse(dr["建立日期"].ToString()).ToString("yyyy/MM/dd HH:mm");
            LblCreater.Text = dr["建立人員"].ToString();
            LblUpdateDT.Text = DateTime.Parse(dr["修改日期"].ToString()).ToString("yyyy/MM/dd HH:mm");
            LblUpdater.Text = dr["修改人員"].ToString();
            TextMemo.Text = dr["備註說明"].ToString();
            lblChecks.Text = GetChecks("作業主機", ApNo);   //取得資料查核
            lblRepeat.Text = HasSameHost(ApNo).Replace("\\n", "");   //取得重複查核
        }

        ReadDevName(TextDevNo.Text);

        cmd.Cancel(); cmd.Dispose(); dr.Close(); Conn.Close(); Conn.Dispose();

        string PowerOff = GetValue("IDMS", "select [關機順序] from [View_通用設備] where [作業編號]=" + TextApNo.Text); //關機順序
        for (int i = 0; i < SelPowerOff.Items.Count; i++) if (SelPowerOff.Items[i].Value == PowerOff) SelPowerOff.SelectedIndex = i;
    }

    protected string GetSysAlias(string SysNo,string Kind)   //取得系統相關資訊
    {
        if (Kind == "SysName") return (GetValue("IDMS", "select [系統全名] from [View_系統資源] where [資源編號]=" + SysNo));
        else return (GetValue("IDMS", "select [資源功能] + ' -- ' + [備註說明] as [系統說明] from [系統資源] where [資源編號]=" + SysNo));        
    }

    protected void ReadDevName(string DevNo)
    {
        if (DevNo == "") DevNo = "0";
        if (DevNo == "0") SelHostType.SelectedIndex = 0;
        
        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
        Conn.Open();
        SqlCommand cmd = new SqlCommand("select [設備編號],[設備種類],[設備名稱] from [實體設備] where [設備編號]=" + DevNo, Conn);
        SqlDataReader dr = null;
        dr = cmd.ExecuteReader();
        if (dr.Read())
        {
            LblDevName.Text = dr[2].ToString();

            if (dr[1].ToString() != "虛擬主機") SelHostType.SelectedIndex = 1;
            else for (int i = 0; i < SelHostType.Items.Count; i++) if (SelHostType.Items[i].Value == dr[0].ToString()) SelHostType.SelectedIndex = i;
        }
        else LblDevName.Text = "";

        cmd.Cancel(); cmd.Dispose(); dr.Close(); Conn.Close(); Conn.Dispose();
    }

    protected string GetSwUnit(string Sw) //讀取維護人員單位或分組名稱
    {
        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
        Conn.Open();
        SqlCommand cmd = new SqlCommand("SELECT Kind FROM Config WHERE Kind in (select Item from Config where Kind='資訊中心') and Item='" + Sw + "'", Conn);
        SqlDataReader dr = null;
        dr = cmd.ExecuteReader();
        string cfg = ""; 
        if (dr.Read()) cfg=dr[0].ToString();  //先讀資訊中心各課成員(各課優先)
        else //再讀其它中心或維護群組(分組次之)
        {
            cmd.Cancel(); cmd.Dispose(); dr.Close();
            cmd = new SqlCommand("SELECT Kind FROM Config WHERE Kind not in (select Item from Config where Kind='資訊中心') and Item='" + Sw + "'", Conn);
            dr = cmd.ExecuteReader();
            if (dr.Read()) cfg=dr[0].ToString();
        }
        cmd.Cancel(); cmd.Dispose(); dr.Close(); Conn.Close(); Conn.Dispose();

        return (cfg);
    }

    protected void GetMaintainor()   //顯示維護人員清單
    {
        DropDownList Sel = (DropDownList)form1.FindControl("SelMaintainor");     //放入Table物件後，必須用FindControl才能取值
        Label Lbl = (Label)form1.FindControl("LblMaintainor");

        SqlCommand sqlQuery = new SqlCommand();
        sqlQuery.CommandText = "select [Item] from [Config] where [Kind]='" + Sel.SelectedValue + "' order by Mark";
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

    protected string GetChecks(string tbl, int PkNo)   //取得資料查核
    {
        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
        Conn.Open();
        SqlCommand cmd = new SqlCommand("select [查核結果] FROM [View_資料查核] WHERE [表格名稱]='" + tbl + "' and [主鍵編號]=" + PkNo, Conn);
        SqlDataReader dr = null;
        dr = cmd.ExecuteReader();
        string cfg = ""; if (dr.Read()) cfg=dr[0].ToString();
        cmd.Cancel(); cmd.Dispose(); dr.Close(); Conn.Close(); Conn.Dispose();

        return (cfg);
    }

    protected void ReadAPHelp() //讀取欄位說明
    {
        SqlCommand sqlQuery = new SqlCommand();
        sqlQuery.CommandText = "select [Item],[Memo],[Config] from [Config] where [Kind]='作業主機' order by [Mark]";
        DataSet ds = RunQuery(sqlQuery);

        string obj = "";

        if (ds.Tables.Count > 0)
        {
            foreach (DataRow row in ds.Tables[0].Rows)
            {
                obj = row[2].ToString();   //物件名稱
                //TableRow Sel = (TableRow)form1.FindControl("row" + obj);    //欄位隱藏
                Label helpObj = (Label)form1.FindControl("help" + obj); //說明文字
                helpObj.Text = row[1].ToString().Replace("\r\n", "<br />");
            }
        }

        sqlQuery.Cancel(); ds.Dispose();

        helpPowerOff.Text = GetValue("IDMS", "select [Memo] from [Config] where [Kind]='實體設備' and [Item]='關機順序'").Replace("\r\n", "<br />");
    }

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
        if (TextDevNo.Text == "") TextDevNo.Text = "0";
        string IsSame = HasSameHost(0).Replace("<br/>", "");        

        if (RightCheck())
        {
            if (TextHostName.Text == "") AddMsg("<script>alert('您未填主機名稱！');</script>");
            else if (SelSysNo.SelectedValue == "0") AddMsg("<script>alert('您未選擇系統名稱！');</script>");
            else if (SelOS.SelectedValue == "") AddMsg("<script>alert('您未選擇作業平台！');</script>");
            else if (SelMaintainor.SelectedValue == "") AddMsg("<script>alert('您未選擇維護人員！');</script>");
            else if (TextDevNo.Text == "" | TextDevNo.Text == "0" | !ExistsCheck("select * from [實體設備] where [設備編號]=" + TextDevNo.Text))
            {
                AddMsg("<script>alert('無此設備編號，請設定之！');</script>");
            }
            else
            {
                LblCreateDT.Text = DateTime.Now.ToString("yyyy/MM/dd HH:mm");
                LblCreater.Text = Session["UserName"].ToString();
                LblUpdateDT.Text = LblCreateDT.Text;
                LblUpdater.Text = LblCreater.Text;

                InsData();

                AddMsg("<script>alert('新增資料[" + TextHostName.Text + "] 完成！\\n\\n" + GetChecks("作業主機",int.Parse(TextApNo.Text)) + "\\n\\n" + IsSame + "');window.open('ApEdit.aspx?ApNo=" + TextApNo.Text + "','_self');</script>");                
            }
        }
        else AddMsg("<script>alert('您不是維護人員，沒有修改 [" + TextHostName.Text + "] 的權限！您可將維護人員設為自己或自己所屬的小組，但不能改為別人！若有需要，可請當事者將這筆資料的維護人員設為自己。');</script>");
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

    protected string HasSameHost(int ApNo) //判斷是否有重複之主機名稱
    {
        string Msg = "";    //回傳重複訊息
        
        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
        Conn.Open();
        SqlCommand cmd = new SqlCommand("select * from [作業主機] where [主機名稱]<>'' and [主機名稱]='" + TextHostName.Text + "'"
                + " and [安內外]='" + SelSaveIO.SelectedValue + "' and [作業編號]<>" + ApNo, Conn);
        SqlDataReader dr = null;
        dr = cmd.ExecuteReader();
        while (dr.Read())
        {
            Msg = Msg + "主機名稱重複：" + dr["作業編號"].ToString() + " " + dr["主機名稱"].ToString() + " " + dr["IP"].ToString() + " " + GetSysAlias(dr["系統編號"].ToString(), "SysName") + "<br/>\\n";
        }
        dr.Close();

        cmd.CommandText = "select * from [作業主機] where [IP]<>'' and [IP]='" + TextIP.Text + "' and [作業編號]<>" + ApNo;
        dr = cmd.ExecuteReader();
        while (dr.Read())
        {
            Msg = Msg + "IP位址重複：" + dr["作業編號"].ToString() + " " + dr["主機名稱"].ToString() + " " + dr["IP"].ToString() + " " + GetSysAlias(dr["系統編號"].ToString(), "SysName") + "<br/>\\n";
        }
        dr.Close();

        cmd.CommandText = "select * from [作業主機] where [設備編號]>0 and [設備編號]=" + TextDevNo.Text + " and [設備編號] not in (select [設備編號] from [實體設備] where [設備種類]='虛擬主機') and [作業編號]<>" + ApNo;
        dr = cmd.ExecuteReader();
        while (dr.Read())
        {
            Msg = Msg + "多重安裝：" + dr["作業編號"].ToString() + " " + dr["主機名稱"].ToString() + "(" + dr["設備編號"].ToString() + ") " + dr["IP"].ToString() + " " + GetSysAlias(dr["系統編號"].ToString(), "SysName") + "<br/>\\n";
        }
        dr.Close();

        cmd.Cancel(); cmd.Dispose(); Conn.Close(); Conn.Dispose();

        lblRepeat.Text = Msg.Replace("\\n", "");
        return (Msg);
    }

    protected Boolean ExistsCheck(string SQL) //判斷設備編號是否合法
    {
        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
        Conn.Open();
        SqlCommand cmd = new SqlCommand(SQL, Conn);
        SqlDataReader dr = null;
        dr = cmd.ExecuteReader();
        Boolean TF = false; if (dr.Read()) TF=true;
        cmd.Cancel(); cmd.Dispose(); dr.Close(); Conn.Close(); Conn.Dispose();

        return (TF);
    }

    protected void BtnDel_Click(object sender, EventArgs e)    //真刪除，請注意生命履歷要詳細記錄所有所刪資料之詳細所有欄位資料
    {
        if (!RightCheck()) AddMsg("<script>alert('您沒有刪除 [" + TextHostName.Text + "] 的權限！');</script>");
        //*****************************待套裝軟體上線後啟用
        //else if (AskCheck()) 
        //{
        //    AddMsg("<script>alert('本作業尚有授權軟體！請先項軟體小組申請解除或轉移後再刪除本作業！');</script>");
        //}
        //*****************************
        else
        {
            string SQL = "delete from [作業主機] where [作業編號]=" + TextApNo.Text;
            InsLifeLog("刪除 [" + TextHostName.Text + "(" + TextIP.Text + ")] ： " + SQL + "，原本作業主機SQL ： " + GetInsApSQL(TextApNo.Text,TextDevNo.Text));
            ExecDbSQL("delete from [系統功能] where [作業編號]=" + TextApNo.Text);
            ExecDbSQL(SQL);

            AddMsg("<script>alert('已刪除刪除作業主機[" + TextHostName.Text + "]及其所屬系統功能！');window.close();</script>");
        }
    }

    protected void BtnEdit_Click(object sender, EventArgs e) //按下修改按鈕
    {
        if (TextDevNo.Text == "") TextDevNo.Text = "0";
        string IsSame = HasSameHost(int.Parse(TextApNo.Text)).Replace("<br/>", "");

        if (RightCheck())
        {
            if (TextApNo.Text == "") AddMsg("<script>alert('無作業編號，故無法修改，請按新增鈕！');</script>");
            else if (TextHostName.Text == "") AddMsg("<script>alert('您未填主機名稱！');</script>");
            else if (SelSysNo.SelectedValue == "0") AddMsg("<script>alert('您未選擇系統名稱！');</script>");
            else if (SelOS.SelectedValue == "") AddMsg("<script>alert('您未選擇作業平台！');</script>");
            else if (SelMaintainor.SelectedValue == "") AddMsg("<script>alert('您未選擇維護人員！');</script>");
            else if (TextDevNo.Text == "" | TextDevNo.Text == "0" | !ExistsCheck("select * from [實體設備] where [設備編號]=" + TextDevNo.Text))
            {
                AddMsg("<script>alert('無此設備編號，請設定之！');</script>");
            }
            else
            {
                LblUpdateDT.Text = DateTime.Now.ToString("yyyy/MM/dd HH:mm");
                LblUpdater.Text = Session["UserName"].ToString();

                InsLifeLog("修改 [" + TextHostName.Text + "(" + TextIP.Text + ")] 的作業主機 ：： " + GetUpdate(int.Parse(TextApNo.Text), "Life"));
                ExecDbSQL("Update [作業主機] set " + GetUpdate(int.Parse(TextApNo.Text),"SQL") + " where [作業編號]=" + TextApNo.Text);

                AddMsg("<script>alert('更新資料 [" + TextHostName.Text + "] 完成！\\n\\n" + GetChecks("作業主機", int.Parse(TextApNo.Text)) + "\\n" + IsSame + "');</script>");
                ReadAp(int.Parse(TextApNo.Text));
            }
        }
        else AddMsg("<script>alert('您不是維護人員，沒有修改 [" + TextHostName.Text + "] 的權限！您可將維護人員設為自己或自己所屬的小組，但不能改為別人！若有需要，可請當事者將這筆資料的維護人員設為自己。');</script>");
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
                    case "string": case "date": case "datetime": SQL = SQL + ",[" + ColName + "]='" + Target + "'"; break;
                    case "integer": case "money": SQL = SQL + ",[" + ColName + "]=" + Target; break;
                    case "null": SQL = SQL + ",[" + ColName + "]=" + null; break;
                    default: SQL = SQL + ",[" + ColName + "]='" + Target + "'"; break;
                }
            }
            else if (SQLorLife == "Life")
            {
                if (Source == "") Source = "(空白)";
                if (Target == "") Target = "(空白)";

                if (ColName != "系統編號")
                {
                    if (Source == "") Source = "(空白)";
                    if (Target == "") Target = "(空白)";
                    SQL = SQL + ",[" + ColName + "]：" + Source + " -> " + Target;
                }
                else
                {
                    string Sname = GetValue("IDMS", "select [資源名稱] from [系統資源] where [資源編號]=" + Source); if (Sname == "") Sname = "(空白)";
                    string Tname = GetValue("IDMS", "select [資源名稱] from [系統資源] where [資源編號]=" + Target); if (Tname == "") Tname = "(空白)";
                    SQL = SQL + ",[" + ColName + "]：" + Source + "(" + Sname + ") -> " + Target + "(" + Tname + ")";
                }
            }
        }
        return (SQL);
    }

    protected string GetUpdate(int ApNo, string SQLorLife) //取得修改資料的SQL語法
    {
        string SQL = "";
        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
        Conn.Open();
        SqlCommand cmd = new SqlCommand("select * from [作業主機] where [作業編號]=" + ApNo, Conn);
        SqlDataReader dr = null;
        dr = cmd.ExecuteReader();
        if (dr.Read())
        {
            SQL = GetUpdateCol("設備編號", dr["設備編號"].ToString(), TextDevNo.Text, "integer", SQLorLife)
                + GetUpdateCol("主機名稱", dr["主機名稱"].ToString(), TextHostName.Text, "string", SQLorLife)
                + GetUpdateCol("系統編號", dr["系統編號"].ToString(), SelSysNo.SelectedValue, "integer", SQLorLife)
                + GetUpdateCol("主要功能", dr["主要功能"].ToString(), TextFunctions.Text, "string", SQLorLife)
                + GetUpdateCol("作業平台", dr["作業平台"].ToString(), SelOS.SelectedValue, "string", SQLorLife)
                + GetUpdateCol("核心版本", dr["核心版本"].ToString(), TextKernel.Text, "string", SQLorLife)
                + GetUpdateCol("IP", dr["IP"].ToString(), TextIP.Text, "string", SQLorLife)
                + GetUpdateCol("監控IP", dr["監控IP"].ToString(), TextMosIP.Text, "string", SQLorLife)
                + GetUpdateCol("安內外", dr["安內外"].ToString(), SelSaveIO.SelectedValue, "string", SQLorLife)
                + GetUpdateCol("緊急程度", dr["緊急程度"].ToString(), SelOnCall.SelectedValue, "string", SQLorLife)
                + GetUpdateCol("作業狀態", dr["作業狀態"].ToString(), SelApStatus.SelectedValue, "string", SQLorLife)
                + GetUpdateCol("維護人員", dr["維護人員"].ToString(), SelMaintainor.SelectedValue, "string", SQLorLife)
                + GetUpdateCol("上線日期", DateTime.Parse(dr["上線日期"].ToString()).ToString("yyyy/MM/dd"), SelYYYY.SelectedValue + "/"  + SelMM.SelectedValue + "/"  + SelDD.SelectedValue, "date", SQLorLife)
                + GetUpdateCol("修改日期", DateTime.Parse(dr["修改日期"].ToString()).ToString("yyyy/MM/dd HH:mm"), DateTime.Now.ToString("yyyy/MM/dd HH:mm"), "datetime", SQLorLife)
                + GetUpdateCol("修改人員", dr["修改人員"].ToString(), Session["UserName"].ToString(), "string", SQLorLife)
                + GetUpdateCol("備註說明", dr["備註說明"].ToString(), TextMemo.Text, "string", SQLorLife);

            if (SQL != "") SQL = SQL.Substring(1);  
        }
        cmd.Cancel(); cmd.Dispose(); dr.Close(); Conn.Close(); Conn.Dispose();

        return (SQL);
    }

    protected void ExecDbSQL(string SQL) //執行資料庫異動
    {
        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
        Conn.Open();
        SqlCommand cmd = new SqlCommand(SQL, Conn);
        cmd.ExecuteNonQuery();
        cmd.Cancel(); cmd.Dispose(); Conn.Close(); Conn.Dispose();
    }

    protected void InsData() //新增資料
    {
        TextApNo.Text = GetPKNo("作業編號", "作業主機").ToString(); //新增完成後，要賦予新取得之作業編號
        string SQL=GetInsApSQL(TextApNo.Text,TextDevNo.Text);
        
        //LblApEdit.Text = SQL;
        ExecDbSQL(SQL);
        InsLifeLog(SQL);
    }

    protected string GetInsApSQL(string ApNo, string DevNo) //取得新增資料的語法
    {
        return ("insert into [作業主機] values("
            + TextApNo.Text + ","
            + TextDevNo.Text + ",'"
            + TextHostName.Text + "',"
            + SelSysNo.SelectedValue + ",'"
            + TextFunctions.Text + "','"
            + SelOS.SelectedValue + "','"
            + TextKernel.Text + "','"
            + TextIP.Text + "','"
            + TextMosIP.Text + "','"
            + SelSaveIO.SelectedValue + "','"
            + SelOnCall.SelectedValue + "','"
            + SelApStatus.SelectedValue + "','"
            + SelMaintainor.SelectedValue + "','"
            + SelYYYY.SelectedValue + "/" + SelMM.SelectedValue + "/" + SelDD.SelectedValue + "','"
            + DateTime.Now.ToString("yyyy/MM/dd HH:mm") + "','"
            + Session["UserName"].ToString() + "','"
            + DateTime.Now.ToString("yyyy/MM/dd HH:mm") + "','"
            + Session["UserName"].ToString() + "','"
            + TextMemo.Text + "')");
    }
    
    protected void BtnClearDevice_Click(object sender, EventArgs e)
    {
        TextDevNo.Text = "0";    //清除設備編號
        LblDevName.Text = "";
    }

    protected void LinkSysNo_Click(object sender, EventArgs e)
    {
        SelSysNo.DataBind();  //讀取系統名稱下拉式選單
    }

    protected void LinkOS_Click(object sender, EventArgs e)
    {
        SelOS.DataBind();  //讀取作業平台下拉式選單
    }
    
    protected string ReadConfig(string Kind,string Item)
    {
        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["DiaryConnectionString"].ConnectionString);
        Conn.Open();
        SqlCommand cmd = new SqlCommand("select [Config] from [Config] where [Kind]='" + Kind +"' and [Item]='" + Item + "'", Conn);
        SqlDataReader dr = null;
        dr = cmd.ExecuteReader();
        string cfg = ""; if (dr.Read()) cfg=dr[0].ToString();
        cmd.Cancel(); cmd.Dispose(); dr.Close(); Conn.Close(); Conn.Dispose();

        return (cfg);
    }

    protected void BtnPing_Click(object sender, EventArgs e) //偵測
    {
        int IsIn = Request.ServerVariables["REMOTE_ADDR"].IndexOf("172.");
        string UrlIn = "http://" + ReadConfig("Primary","SSM(In)") + "/ping.asp?hostname=" + TextHostName.Text + "&IP=" + TextIP.Text;
        string UrlOut = "/Public/ping.aspx?hostname=" + TextHostName.Text + "&IP=" + TextIP.Text + "&PssmIn=sos-ws01";

        if (SelSaveIO.SelectedValue == "I" & IsIn < 0)
            AddMsg("<script>alert('[" + TextHostName.Text + "]為[" + SelSaveIO.SelectedValue + "]機器，請至該網段執行偵測！');</script>");
        else
            if (IsIn < 0) AddMsg("<script>window.open('" + UrlOut + "','_self');</script>");
            else AddMsg("<script>window.open('" + UrlIn + "','_self');</script>");
    }

    protected Boolean RightCheck() //是否有權修改資料
    {
        string UserID = Session["UserID"].ToString().ToLower();
        string UserName = Session["UserName"].ToString();   //登入的UserName
        string UnitName = Session["UnitName"].ToString();   //登入的UnitName
        string Sw = SelMaintainor.SelectedValue;  //作業維護人員 
        string Dno = TextDevNo.Text; if (Dno == "") Dno = "0";
        string Hw = GetValue("IDMS", "select [維護人員] from [實體設備] where [設備編號]=" + Dno); //設備維護人員
        string Older = GetValue("IDMS", "select [維護人員] from [作業主機] where [作業編號]=" + Dno);

        if (UserID != "operator" & (InGroup(UserName, Sw) | InGroup(UserName, Older) | InGroup(UserName, Hw) | UnitName == "系統控制課" | UnitName == "電腦操作課" | InGroup(UserName, "軟體小組") | InGroup(UserName, "網管小組"))) return (true);
        else return (false);
    }

    protected string GetValue(string DB, string SQL)   //取得單一資料
    {
        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings[DB + "ConnectionString"].ConnectionString);
        Conn.Open();
        SqlCommand cmd = new SqlCommand(SQL, Conn);
        SqlDataReader dr = null;
        dr = cmd.ExecuteReader();
        string cfg = ""; if (dr.Read()) cfg=dr[0].ToString();
        cmd.Cancel(); cmd.Dispose(); dr.Close(); Conn.Close(); Conn.Dispose();

        return (cfg);
    }

    protected Boolean AskCheck() //是否尚有授權軟體
    {
        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
        Conn.Open();
        SqlCommand cmd = new SqlCommand("select * from [軟體授權] where [作業編號]=" + TextApNo.Text, Conn);
        SqlDataReader dr = null;
        dr = cmd.ExecuteReader();
        Boolean TF = false; if (dr.Read()) TF=true;
        cmd.Cancel(); cmd.Dispose(); dr.Close(); Conn.Close(); Conn.Dispose();

        return (TF);
    }

    protected Boolean InGroup(string ChkName,string ChkUnit) //檢查ChkName是否為ChkUnit成員或本身
    {
        Boolean TF = false;
        
        if (ChkName==ChkUnit) TF=true;  //是否同名
        else if(ChkUnit == "") TF=false;	//檢查單位必填
        else //是否為成員UN (課別與小組同義)
        {
            SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
            Conn.Open();
            SqlCommand cmd = new SqlCommand("select * from Config where (Kind='" + ChkUnit + "') and Item='" + ChkName + "'", Conn);
            SqlDataReader dr = null;
            dr = cmd.ExecuteReader();
            TF=true;
            cmd.Cancel(); cmd.Dispose(); dr.Close(); Conn.Close(); Conn.Dispose();
        }

        return (TF);
    }

    protected string GetDuty(string Maintainor) //取得使用課別
    {
        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
        Conn.Open();
        string SQL = "(select Kind from Config where Item='" + Maintainor + "' and Kind in (select Item from Config where Kind='資訊中心'))"
            + " UNION (select Config from Config where Item='" + Maintainor + "' and kind in (select Item from Config where Kind='維護群組') and Kind<>'資訊中心')"
            + " UNION (select Item from Config where Item='" + Maintainor + "' and Kind='資訊中心')";
        SqlCommand cmd = new SqlCommand(SQL, Conn);
        SqlDataReader dr = null;
        dr = cmd.ExecuteReader();
        string cfg = "";
        if (Maintainor=="") cfg="";
        else if (dr.Read()) cfg=dr[0].ToString();
        cmd.Cancel(); cmd.Dispose(); dr.Close(); Conn.Close(); Conn.Dispose();

        return (cfg);
    }

    protected void InsLifeLog(string SQL) //寫入生命履歷
    {
        string LifeNo = GetPKNo("履歷編號", "生命履歷").ToString(); //履歷編號
        string TblName = "作業主機";    //表格名稱
        string PKno = TextApNo.Text;   //主鍵編號
        string Mt = SelMaintainor.SelectedValue;    //維護人員
        string Dno = TextDevNo.Text; if (Dno == "") Dno = "0";
        string OldMt = GetValue("IDMS","select [維護人員] from [作業主機] where [作業編號]=" + PKno);    //原負責人
        string OldKeeper = GetValue("IDMS","select [保管人員] from [View_設備管理] where [設備編號]=" + Dno);    //原保管人
        string UN = Session["UserName"].ToString();   //登入的UserName：異動人員
        string LiftDT = DateTime.Now.ToString("yyyy/MM/dd HH:mm");  //異動日期
        string LifeIP = Request.ServerVariables["REMOTE_ADDR"].ToString();

        ExecDbSQL("Insert into [生命履歷] values(" + LifeNo + ",'" + TblName + "'," + PKno + ",'" + SQL.Replace("'", "''") + "','" + Mt + "','" + OldMt + "','" + OldKeeper + "','" + UN + "','" + LiftDT + "','" + LifeIP + "')");        
    }

    protected void LinkLcs_Click(object sender, EventArgs e)  //授權查詢
    {
        Session["AskSQL"] = "SELECT * FROM [View_軟體管理] WHERE [授權編號] is not null and [作業編號]=" + TextApNo.Text;
        OpenExecWindow("window.open('../Software/Ask.aspx?Search=SwEdit','_self');");
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

    protected void BtnLife_Click(object sender, EventArgs e) //查詢生命履歷
    {
        Session["LifeSQL"] = "select * from [生命履歷] where [表格名稱]='作業主機' and [主鍵編號]=" + TextApNo.Text + " order by [履歷編號] desc";
        OpenExecWindow("window.open('../Life/LifeLog.aspx?Search=ApEdit&Tbl=作業主機&PK=" + TextApNo.Text + "','_self');");
    }

    protected void SelHostType_SelectedIndexChanged(object sender, EventArgs e)   //虛擬主機負責人
    {
        if (SelHostType.SelectedValue == "-1") OpenExecWindow("window.open('../Device/Device.aspx?SelMode=Y','_blank');");
        else TextDevNo.Text = SelHostType.SelectedValue;
    }

    protected void SelSysNo_SelectedIndexChanged(object sender, EventArgs e)   //系統說明
    {
        lblSysAlias.Text = GetSysAlias(SelSysNo.SelectedValue,"");
        lblAssets.Text = GetValue("IDMS", "select [Config] from [Config],[系統資源] where [Config].[Item]=[系統資源].[資產編號] and [系統資源].[資源編號]=" + SelSysNo.SelectedValue);
    } 

    protected void BtnPowerOff_Click(object sender, EventArgs e) //更新關機順序
    {
        if (RightCheck())
        {
            string PowerOff=GetValue("IDMS","select [關機順序] from [實體設備] where [設備編號]=" + TextDevNo.Text);
            
            if (TextApNo.Text == "") AddMsg("<script>alert('請先建立作業主機資料再修改關機順序！');</script>");
            else if (PowerOff == "") AddMsg("<script>alert('作業主機對應不到實體設備，請檢查設備編號是否有誤！');</script>");
            else if (SelHostType.SelectedValue != "-1") AddMsg("<script>alert('一般主機才需設定關機順序，虛擬主機不用！');</script>");
            else if (PowerOff == SelPowerOff.SelectedValue) AddMsg("<script>alert('關機順序相同，不用更新！');</script>");            
            else
            {
                LblUpdateDT.Text = DateTime.Now.ToString("yyyy/MM/dd HH:mm");
                LblUpdater.Text = Session["UserName"].ToString();

                InsLifeLog("修改 [" + TextHostName.Text + "(" + TextIP.Text + ")] 作業主機的關機順序 ：： " + PowerOff + " -> " + SelPowerOff.SelectedValue);
                ExecDbSQL("Update [實體設備] set [關機順序]='" + SelPowerOff.SelectedValue + "' where [設備編號]=" + TextDevNo.Text);

                AddMsg("<script>alert('更新 [" + TextHostName.Text + "] 關機順序完成！(" + PowerOff + " -> " + SelPowerOff.SelectedValue + ")');</script>");
                ReadAp(int.Parse(TextApNo.Text));
            }
        }
        else AddMsg("<script>alert('您不是維護人員，沒有修改 [" + TextHostName.Text + "] 的權限！您可將維護人員設為自己或自己所屬的小組，但不能改為別人！若有需要，可請當事者將這筆資料的維護人員設為自己。');</script>");
    }

    protected void AddMsg(string strMsg)
    {
        Literal Msg = new Literal();
        Msg.Text = strMsg;
        Page.Controls.Add(Msg);
    }

    protected void LinkCopyDevName_Click(object sender, EventArgs e)  //複製設備名稱
    {
        if (TextDevNo.Text != "0" & TextDevNo.Text != "")
        {
            TextHostName.Text = GetValue("IDMS", "select [設備名稱] from [實體設備] where [設備編號]=" + TextDevNo.Text);
        }
        else AddMsg("<script>alert('請先指定您要複製名稱的設備編號！');</script>");
    }
    protected void LinkCopyDevFunc_Click(object sender, EventArgs e)  //複製設備功能
    {
        if (TextDevNo.Text != "0" & TextDevNo.Text != "")
        {
            TextFunctions.Text = GetValue("IDMS", "select [設備用途] from [實體設備] where [設備編號]=" + TextDevNo.Text);
        }
        else AddMsg("<script>alert('請先指定您要複製功能的設備編號！');</script>");
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
            for (int i = 0; i < Table1.Rows.Count; i++) Table1.Rows[i].Cells[2].Visible = false;
            LinkHideHelp.Text = "顯示說明";
        }
        else
        {
            for (int i = 0; i < Table1.Rows.Count; i++) Table1.Rows[i].Cells[2].Visible = true;
            LinkHideHelp.Text = "隱藏說明";
        }
    }

    protected void SelOS_SelectedIndexChanged(object sender, EventArgs e)   //產生平台核心
    {
        SelKernel.Items.Clear();        
        
        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
        Conn.Open();
        SqlCommand cmd = new SqlCommand("select distinct [核心版本] from [作業主機] where [作業平台]='" + SelOS.SelectedValue + "'", Conn);
        SqlDataReader dr = null;
        dr = cmd.ExecuteReader();
        while (dr.Read()) SelKernel.Items.Add(dr[0].ToString());
        cmd.Cancel(); cmd.Dispose(); dr.Close(); Conn.Close(); Conn.Dispose();
    }
    protected void SelKernel_SelectedIndexChanged(object sender, EventArgs e)   //平台核心填值
    {
        TextKernel.Text = SelKernel.SelectedValue;
    }
}