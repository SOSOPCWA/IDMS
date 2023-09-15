using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.Configuration;
using System.Data;
using System.Data.SqlClient;

public partial class SOS_ChkEdit : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        if (!IsPostBack)
        {
            ReadHelp(); //讀取欄位說明

            //填入預設值
            TextDevNo.Text = "0";
            TextPointerNo.Text = "0";
        }
    }

    protected void Page_PreRenderComplete(object sender, EventArgs e)   //放Page_Load無法顯示
    {
        TextChkYear.Text = Request["ChkYear"];
        TextDevNo.Text = Request["DevNo"];

        if (TextDevNo.Text == "")    //各按鈕啟用與否
        {
            BtnEdit.Enabled = false; BtnDel.Enabled = false;
        }
        else
        {
            BtnEdit.Enabled = true; BtnDel.Enabled = true;
        }

        if (!IsPostBack)   //顯示各下拉式選單(DB)，注意讀值先後順序
        {
            if (TextDevNo.Text != "") ReadCheck();        //讀取機器清查 
        }

        GetMt();   //顯示維護人員清單        
    }

    protected void GetMt()   //顯示維護人員清單
    {
        DropDownList Sel = (DropDownList)form1.FindControl("SelMt");     //放入Table物件後，必須用FindControl才能取值
        Label Lbl = (Label)form1.FindControl("LblMt");

        SqlCommand sqlQuery = new SqlCommand();
        sqlQuery.CommandText = "select [Item] from [Config] where [Kind]='" + Sel.SelectedValue + "' order by Mark";
        DataSet ds = RunQuery(sqlQuery);

        LblMt.Text = "";
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


    protected void ReadCheck()    //讀取機器清查
    {
        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
        Conn.Open();
        SqlCommand cmd = new SqlCommand("SELECT * FROM [View_機器清查] WHERE [清查年度]='" + TextChkYear.Text + "' and [設備編號]=" + TextDevNo.Text, Conn);
        SqlDataReader dr = null;
        dr = cmd.ExecuteReader();

        if (dr.Read())
        {
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
            
            lblHost.Text = "<br />[設備種類]=" + dr["設備種類"].ToString() + "<br />"
                + "[設備名稱]=" + dr["設備名稱"].ToString() + "<br />"
                + "[設備用途]=" + dr["設備用途"].ToString() + "<br />"
                + "[財產編號]=" + dr["財產編號"].ToString() + "<br />"
                + "[廠牌型式]=" + dr["廠牌"].ToString() + " " + dr["型式"].ToString() + "<br />"
                + "[放置地點]=" + dr["設備區域名稱"].ToString() + dr["設備定位名稱"].ToString() + "<br />"
                + "[維護人員]=" + dr["設備維護人員"].ToString();

            string HwUnit = GetUnit(dr["維護人員"].ToString());
            for (int i = 0; i < SelHwUnit.Items.Count; i++) if (SelHwUnit.Items[i].Value == HwUnit) SelHwUnit.SelectedIndex = i;
            SelMt.DataBind();  //連帶選項需要觸發
            for (int i = 0; i < SelMt.Items.Count; i++) if (SelMt.Items[i].Value == dr["維護人員"].ToString()) SelMt.SelectedIndex = i;            

            for (int i = 0; i < ChkChecks.Items.Count; i++)//清查結果
            {
                if (dr["清查結果"].ToString().IndexOf(ChkChecks.Items[i].Value) >= 0) ChkChecks.Items[i].Selected = true;
            }          

            TextMemo.Text = dr["備註說明"].ToString();

            for (int i = 0; i < SelStatus.Items.Count; i++) if (SelStatus.Items[i].Value == dr["清查狀態"].ToString()) SelStatus.SelectedIndex = i;
            
            LblCheckDT.Text = DateTime.Parse(dr["清查時間"].ToString()).ToString("yyyy/MM/dd HH:mm");
            LblChecker.Text = dr["清查人員"].ToString();

            LblUpdateDT.Text = DateTime.Parse(dr["更新時間"].ToString()).ToString("yyyy/MM/dd HH:mm");
            LblUpdater.Text = dr["更新人員"].ToString();
            
        }

        cmd.Cancel(); dr.Close(); Conn.Close(); Conn.Dispose();
    }

    protected string GetUnit(string Pname) //讀取人員單位或分組名稱
    {
        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
        Conn.Open();
        SqlCommand cmd = new SqlCommand("SELECT [課別] FROM [View_組織架構] WHERE [性質]='員工' and [成員]='" + Pname + "'", Conn);
        SqlDataReader dr = null;
        dr = cmd.ExecuteReader();
        if (dr.Read())  //先讀數值資訊組各課成員(各課優先)
        {
            return (dr[0].ToString());
        }
        else  //再讀其它中心或維護群組(分組次之)
        {
            cmd.Cancel(); dr.Close();
            cmd = new SqlCommand("SELECT Kind FROM Config WHERE Kind not in (select Item from Config where Kind='數值資訊組') and Item='" + Pname + "'", Conn);
            dr = cmd.ExecuteReader();
            if (dr.Read())
            {
                return (dr[0].ToString());
            }
            else
            {
                return ("");
            }
        }
        cmd.Cancel(); dr.Close(); Conn.Close(); Conn.Dispose();
    }

    protected void ReadHelp() //讀取欄位說明
    {
        SqlCommand sqlQuery = new SqlCommand();
        sqlQuery.CommandText = "select Item,Memo,Config from Config where Kind='機器清查' order by Mark";
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
        Literal Msg = new Literal();

        if (!RightCheck())
        {
            Msg.Text = "<script>alert('您不是環境小組的成員，沒有新增機器清查的權限！');</script>";
        }
        else
        {
            if (TextDevNo.Text == "") TextDevNo.Text = "0";

            LblCheckDT.Text = DateTime.Now.ToString("yyyy/MM/dd HH:mm");
            LblChecker.Text = Session["UserName"].ToString();
            LblUpdateDT.Text = LblCheckDT.Text;
            LblUpdater.Text = LblUpdater.Text;

            string DevNo = GetPKNo("設備編號", "機器清查").ToString();
            TextDevNo.Text = DevNo; //新增完成後，要賦予新取得之設備編號 

            string SQL = GetInsSQL(DevNo);
            //Response.Write(SQL);
            //Response.End();
            ExecDbSQL(SQL);

            Msg.Text = "<script>alert('新增機器清查資料[" + TextChkYear.Text + "] - [" + TextDevNo.Text + "] 完成！');window.open('ChkEdit.aspx?ChkYear=" + TextChkYear.Text + "&DevNo=" + TextDevNo.Text + "','_self');</script>";  //新增後關視窗是因無法改PK
        }

        Page.Controls.Add(Msg);
    }

    protected int GetPKNo(string PKfield, string PKtbl) //取得主鍵編號
    {
        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
        Conn.Open();
        SqlCommand cmd = new SqlCommand("select min(" + PKfield + ") from " + PKtbl, Conn);
        SqlDataReader dr = null;
        dr = cmd.ExecuteReader();

        int PK = -1;
        if (dr.Read())
        {
            if (int.Parse(dr[0].ToString()) < 0) PK = int.Parse(dr[0].ToString()) - 1;            
        }
        
        cmd.Cancel(); dr.Close(); Conn.Close(); Conn.Dispose();

        return (PK);        
    }

    protected void BtnDel_Click(object sender, EventArgs e)    //真刪除，請注意生命履歷要詳細記錄所有所刪資料之詳細所有欄位資料
    {
        Literal Msg = new Literal();

        if (!RightCheck())
        {
            Msg.Text = "<script>alert('您不是環境小組的成員，沒有刪除機器清查資料的權限！');</script>";
        }
        else
        {
            SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
            Conn.Open();
            SqlCommand cmd = new SqlCommand("select * from [實體設備] where [設備編號]=" + TextDevNo.Text, Conn);
            SqlDataReader dr = null;
            dr = cmd.ExecuteReader();
            if (dr.Read()) Msg.Text = "<script>alert('IDMS在 [實體設備] 內有的資料，[機器清查] 禁止刪除；須等負責人刪除或修正放置地點後，才可刪除之！');</script>";
            else
            {
                string SQL = "delete from [機器清查] where [清查年度]='" + TextChkYear.Text + "' and [設備編號]=" + TextDevNo.Text;
                ExecDbSQL(SQL);

                Msg.Text = "<script>alert('完成刪除 [" + TextChkYear.Text + "] - [" + TextDevNo.Text + "] 資料！');window.close();</script>";
            }
            cmd.Cancel(); dr.Close(); Conn.Close(); Conn.Dispose();
        }

        Page.Controls.Add(Msg);
    }

    protected void BtnEdit_Click(object sender, EventArgs e) //按下修改按鈕
    {
        Literal Msg = new Literal();

        if (!RightCheck())
        {
            Msg.Text = "<script>alert('您不是環境小組的成員，沒有異動機器清查資料的權限！');</script>";
        }
        else
        {
            LblUpdateDT.Text = DateTime.Now.ToString("yyyy/MM/dd HH:mm");
            LblUpdater.Text = Session["UserName"].ToString();

            ExecDbSQL("Update [機器清查] set " + GetUpdate(int.Parse(TextDevNo.Text), "SQL") + " where [清查年度]='" + TextChkYear.Text + "' and [設備編號]=" + TextDevNo.Text);

            Msg.Text = "<script>alert('更新資料 [" + TextChkYear.Text + "] - [" + TextDevNo.Text + "] 完成！');window.open('ChkEdit.aspx?ChkYear=" + TextChkYear.Text + "&DevNo=" + TextDevNo.Text + "','_self');</script>";
        }

        Page.Controls.Add(Msg);
    }

    protected string GetInsSQL(string DevNo) //取得新增資料的語法
    {
        string PointerNo = SelPointer.SelectedValue;
        if (TextPointerNo.Text != "") PointerNo = TextPointerNo.Text;
        
        return ("insert into [機器清查] values('"
            + TextChkYear.Text + "',"
            + DevNo + ","
            + PointerNo + ",'"
            + SelMt.SelectedValue + "','"
            + GetChecks() + "','"
            + TextMemo.Text + "','"
            + SelStatus.SelectedValue + "','"
            + Session["UserName"].ToString() + "','"
            + LblCheckDT.Text + "','"
            + Session["UserName"].ToString() + "','"
            + LblUpdateDT.Text + "'"              
            + ")");
    }

    protected string GetChecks() //取得清查結果
    {
        string Checks = "";
        for (int i = 0; i < ChkChecks.Items.Count; i++)
        {
            if (ChkChecks.Items[i].Selected)
            {
                if (Checks == "") Checks = ChkChecks.Items[i].Value;
                else Checks = Checks + "," + ChkChecks.Items[i].Value;
            }
        }

        return (Checks);
    }

    protected string GetUpdate(int DevNo, string SQLorLife) //取得修改資料的語法
    {
        string SQL = "";
        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
        Conn.Open();
        SqlCommand cmd = new SqlCommand("select * from [機器清查] where [清查年度]='" + TextChkYear.Text + "' and [設備編號]=" + DevNo, Conn);
        SqlDataReader dr = null;
        dr = cmd.ExecuteReader();

        string PointerNo = SelPointer.SelectedValue;
        if (TextPointerNo.Text != "") PointerNo = TextPointerNo.Text;

        if (dr.Read())
        {
            SQL = GetUpdateCol("定位編號", dr["定位編號"].ToString(), PointerNo, "integer", SQLorLife)
                + GetUpdateCol("維護人員", dr["維護人員"].ToString(), SelMt.SelectedValue, "string", SQLorLife)
                + GetUpdateCol("清查結果", dr["清查結果"].ToString(), GetChecks(), "string", SQLorLife)
                + GetUpdateCol("備註說明", dr["備註說明"].ToString(), TextMemo.Text, "string", SQLorLife)                ;
            if (dr["清查狀態"].ToString() == "待清查")
            {
                SQL = SQL + GetUpdateCol("清查狀態", dr["清查狀態"].ToString(), SelStatus.SelectedValue, "string", SQLorLife)
                    + GetUpdateCol("清查時間", DateTime.Parse(dr["清查時間"].ToString()).ToString("yyyy/MM/dd HH:mm"), DateTime.Now.ToString("yyyy/MM/dd HH:mm"), "datetime", SQLorLife)
                    + GetUpdateCol("清查人員", dr["清查人員"].ToString(), Session["UserName"].ToString(), "string", SQLorLife);
            }
            else
            {
                SQL = SQL + GetUpdateCol("清查狀態", dr["清查狀態"].ToString(), SelStatus.SelectedValue, "string", SQLorLife);
            }
            SQL=SQL + GetUpdateCol("更新時間", DateTime.Parse(dr["更新時間"].ToString()).ToString("yyyy/MM/dd HH:mm"), DateTime.Now.ToString("yyyy/MM/dd HH:mm"), "datetime", SQLorLife)
                + GetUpdateCol("更新人員", dr["更新人員"].ToString(), Session["UserName"].ToString(), "string", SQLorLife);

            if (SQL != "") SQL = SQL.Substring(1);
        }
        cmd.Cancel(); dr.Close(); Conn.Close(); Conn.Dispose();

        return SQL;
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
                    case "money":
                    case "null": SQL = SQL + ",[" + ColName + "]=" + Target; break;
                    default: SQL = SQL + ",[" + ColName + "]='" + Target + "'"; break;
                }
            }
            else if (SQLorLife == "Life")
            {
                SQL = SQL + ",[" + ColName + "]=" + Source + " -> " + Target;
            }
        }
        return (SQL);
    }    

    protected void ExecDbSQL(string SQL) //執行資料庫異動
    {
        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
        Conn.Open();
        SqlCommand cmd = new SqlCommand(SQL, Conn);
        //Response.Write(cmd.CommandText);
        //Response.End();
        cmd.ExecuteNonQuery();
        cmd.Cancel(); Conn.Close(); Conn.Dispose();
    }

    protected Boolean RightCheck() //是否有權修改資料
    {
        string UserName=Session["UserName"].ToString();
        if (InGroup(UserName, "環境小組")) return (true);
        else return (false);
    }

    protected Boolean InGroup(string ChkName, string ChkUnit) //檢查ChkName是否為ChkUnit成員或本身
    {
        if (ChkName == ChkUnit) return (true);  //是否同名
        else if (ChkUnit == "") return (false);	//檢查單位必填
        else //是否為成員UN (課別與小組同義)
        {
            SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
            Conn.Open();
            SqlCommand cmd = new SqlCommand("select * from Config where (Kind='" + ChkUnit + "') and Item='" + ChkName + "'", Conn);
            SqlDataReader dr = null;
            dr = cmd.ExecuteReader();
            if (dr.Read()) return (true);
            cmd.Cancel(); dr.Close(); Conn.Close(); Conn.Dispose();
        }

        return (false);
    }

    protected void OpenExecWindow(string strJavascript)    //選取後就另開視窗
    {
        if (ScriptManager.GetCurrent(Page) == null)
        {
            Page.ClientScript.RegisterStartupScript(Page.GetType(), "ChkEdit", strJavascript, true);
        }
        else
        {
            ScriptManager.RegisterStartupScript(Page, Page.GetType(), "ChkEdit", strJavascript, true);
        }
    }

    protected void BtnDev_Click(object sender, EventArgs e) //按下設備按鈕
    {
        OpenExecWindow("window.open('../Device/DevEdit.aspx?DevNo=" + TextDevNo.Text + "','_blank');");
    }

    protected void SelPointer_SelectedIndexChanged(object sender, EventArgs e)  //定位編號
    {
        DropDownList Sel = (DropDownList)form1.FindControl("SelPointer");    //放入Table物件後，必須用FindControl才能取值
        TextPointerNo.Text = ""; // Sel.SelectedValue;
    }
}