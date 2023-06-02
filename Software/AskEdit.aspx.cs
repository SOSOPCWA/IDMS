using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.Configuration;
using System.Data;
using System.Data.SqlClient;

public partial class Software_AskEdit : System.Web.UI.Page
{
    const int MinYear = 1998;   //起始年
    const int EverYear = 2099;   //永久年
    int NowYear = DateTime.Now.Year;
    const int OverYear = 10;    //年選單多產生幾年

    protected void Page_PreRenderComplete(object sender, EventArgs e)   //放Page_Load無法顯示
    {
        if (!IsPostBack)
        {
            if (!RightCheck())
            {
                Literal Msg = new Literal();
                Msg.Text = "<script>alert('您沒有檢視資料的權利！');window.close();</script>";
                Page.Controls.Add(Msg);
            }
            else
            {
                TextAskNo.Text = Request["AskNo"];

                if (TextAskNo.Text == "")    //各按鈕啟用與否
                {
                    BtnEdit.Enabled = false; BtnDel.Enabled = false;
                    LinkAskFormAdd.Visible = false;
                }
                else
                {
                    BtnEdit.Enabled = true; BtnDel.Enabled = true;
                    LinkAskFormAdd.Visible = true;
                }

                //顯示各下拉式選單(DB)，注意讀值先後順序
                SelAskFormNo.DataBind(); SelCause.DataBind(); SelExec.DataBind();

                if (TextAskNo.Text != "") ReadAsk();        //讀取軟體授權 

                SelVersAsk.DataBind(); for (int i = 0; i < SelVersAsk.Items.Count; i++) if (SelVersAsk.Items[i].Value == TextVersAsk.Text) SelVersAsk.SelectedIndex = i;
                SelLcsAsk.DataBind(); SelStyleLcs.DataBind();
                SelUserLcs.DataBind(); for (int i = 0; i < SelUserLcs.Items.Count; i++) if (SelUserLcs.Items[i].Value == TextUserLcs.Text) SelUserLcs.SelectedIndex = i;

                GetSw();    //於Label顯示軟體名稱
                GetAP();    //於Label顯示作業主機名稱        
                SelAskFormNo_SelectedIndexChanged(null, null);  //顯示申請表單資訊      
            }

            ShowHideHelp();  //顯示或隱藏欄位說明
        }
    }
    
    protected void Page_Load(object sender, EventArgs e)
    {
        if (!IsPostBack)
        {
            ReadHelp(); //讀取欄位說明

            //產生各種年月日下拉式選單
            AddSel(SelStatusYYYY, MinYear, NowYear + OverYear); AddSel(SelStatusMM, 1, 12); AddSel(SelStatusDD, 1, 31);
            AddSel(SelKeyYYYY, MinYear, NowYear + OverYear); AddSel(SelKeyMM, 1, 12); AddSel(SelKeyDD, 1, 31);
            AddSel(SelFormYYYY, MinYear, NowYear + OverYear); AddSel(SelFormMM, 1, 12); AddSel(SelFormDD, 1, 31);

            //填入預設值
            TextSwNo.Text = "0"; TextApNo.Text = "0"; TextIdentify.Text = "0";

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

    protected void AddYearSel(DropDownList Sel)   //產生永久年
    {
        ListItem Item = new ListItem();
        Item.Text = "(永久)";
        Item.Value = EverYear.ToString();

        Sel.Items.Add(Item);
    }    

    protected void GetSw()   //顯示軟體名稱
    {
        if (TextSwNo.Text != "")
        {
            SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
            Conn.Open();
            SqlCommand cmd = new SqlCommand("select [軟體名稱],[購買版本],[授權方式],[授權數量] FROM [軟體主檔] WHERE [軟體編號]=" + TextSwNo.Text, Conn);
            SqlDataReader dr = null;
            dr = cmd.ExecuteReader();

            if (dr.Read()) LblSwName.Text = dr[0].ToString() + " " + dr[1].ToString() + " " + dr[2].ToString() + " * " + dr[3].ToString();
            else LblSwName.Text = "(尚未設定)";

            cmd.Cancel(); cmd.Dispose(); dr.Close(); Conn.Close(); Conn.Dispose();
        }
        else
        {
            TextSwNo.Text = "0";
            LblSwName.Text = "(尚未設定)";
        }
    }

    protected void GetAP()   //顯示作業主機名稱
    {
        if (TextApNo.Text != "")
        {
            SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
            Conn.Open();
            SqlCommand cmd = new SqlCommand("select [系統全名],[主機名稱] FROM [View_通用設備] WHERE [作業編號]=" + TextApNo.Text, Conn);
            SqlDataReader dr = null;
            dr = cmd.ExecuteReader();

            if (dr.Read()) LblApNo.Text = dr[0].ToString() + "-" + dr[1].ToString();
            else LblApNo.Text = "(尚未設定)";

            cmd.Cancel(); cmd.Dispose(); dr.Close(); Conn.Close(); Conn.Dispose();
        }
        else
        {
            TextApNo.Text = "0";
            LblApNo.Text = "(尚未設定)";
        }
    }

    protected string GetChecks(string tbl, string PkNo)   //取得資料查核
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

    protected void ReadAsk()    //讀取軟體授權
    {
        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
        Conn.Open();
        SqlCommand cmd = new SqlCommand("SELECT * FROM [View_軟體管理] WHERE [授權編號]=" + TextAskNo.Text, Conn);
        SqlDataReader dr = null;
        dr = cmd.ExecuteReader();

        if (dr.Read())
        {
            TextSwNo.Text = dr["軟體編號"].ToString();
            TextFormNo.Text = dr["表單編號"].ToString();
            TextApNo.Text = dr["授權作業編號"].ToString();
            TextIdentify.Text = dr["授權識別"].ToString();

            lblVersBuy.Text = dr["購買版本"].ToString();
            lblVersCan.Text = dr["可用版本"].ToString();
            TextVersAsk.Text = dr["申請版本"].ToString();            

            lblLcsBuy.Text = dr["購買授權"].ToString();
            for (int i = 0; i < SelLcsAsk.Items.Count; i++) if (SelLcsAsk.Items[i].Value == dr["申請授權"].ToString()) SelLcsAsk.SelectedIndex = i;
            TextLcsMemo.Text = dr["授權說明"].ToString();

            if (RightCheck())   //有無權利看到序號
            {
                lblSnBuy.Text = dr["購買序號"].ToString();
                TextSnAsk.Text = dr["申請序號"].ToString();
            }

            lblAttach.Text = dr["軟體附件"].ToString();
            TextAttach.Text = dr["授權附件"].ToString();

            for (int i = 0; i < SelUnitLcs.Items.Count; i++) if (SelUnitLcs.Items[i].Value == dr["授權單位"].ToString()) SelUnitLcs.SelectedIndex = i;
            lblUnitUse.Text = dr["使用單位"].ToString();

            TextUserLcs.Text = dr["授權人員"].ToString();
            lblUserUse.Text = dr["使用人員"].ToString();

            TextHostLcs.Text = dr["授權主機"].ToString();
            lblHostUse.Text = dr["使用主機"].ToString();

            TextIPLcs.Text = dr["授權IP"].ToString();
            lblIPUse.Text = dr["使用IP"].ToString();

            TextPropNoALcs.Text = dr["授權財編A"].ToString();
            TextPropNoBLcs.Text = dr["授權財編B"].ToString();
            lblPropNoAUse.Text = dr["使用財編A"].ToString();
            lblPropNoBUse.Text = dr["使用財編B"].ToString();

            TextBrandLcs.Text = dr["授權廠牌"].ToString();
            TextStyleLcs.Text = dr["授權型式"].ToString();
            lblBrandUse.Text = dr["使用廠牌"].ToString();
            lblStyleUse.Text = dr["使用型式"].ToString();            

            for (int i = 0; i < SelStatus.Items.Count; i++) if (SelStatus.Items[i].Value == dr["授權狀態"].ToString()) SelStatus.SelectedIndex = i;
            lblSwStatus.Text = dr["軟體狀態"].ToString();

            for (int i = 0; i < SelCause.Items.Count; i++) if (SelCause.Items[i].Value == dr["授權減損原因"].ToString()) SelCause.SelectedIndex = i;
            TextCause.Text = dr["授權減損原因"].ToString();
            for (int i = 0; i < SelExec.Items.Count; i++) if (SelExec.Items[i].Value == dr["授權減損方式"].ToString()) SelExec.SelectedIndex = i;
            TextExec.Text = dr["授權減損方式"].ToString();
            if (!(dr["授權減損日期"] is DBNull))
            {
                int StatusYYYY = DateTime.Parse(dr["授權減損日期"].ToString()).Year;
                if (StatusYYYY - MinYear + 1 > SelStatusYYYY.Items.Count) SelStatusYYYY.SelectedIndex = SelStatusYYYY.Items.Count - 1;
                else SelStatusYYYY.SelectedIndex = StatusYYYY - MinYear + 1;

                int StatusMM = DateTime.Parse(dr["授權減損日期"].ToString()).Month; SelStatusMM.SelectedIndex = StatusMM;
                int StatusDD = DateTime.Parse(dr["授權減損日期"].ToString()).Day; SelStatusDD.SelectedIndex = StatusDD;
            }
            else
            {
                SelStatusYYYY.SelectedIndex = 0; SelStatusMM.SelectedIndex = 0; SelStatusDD.SelectedIndex = 0;
            }
            if (!(dr["授權填表日期"] is DBNull))
            {
                int KeyYYYY = DateTime.Parse(dr["授權填表日期"].ToString()).Year;
                if (KeyYYYY - MinYear + 1 > SelKeyYYYY.Items.Count) SelKeyYYYY.SelectedIndex = SelKeyYYYY.Items.Count - 1;
                else SelKeyYYYY.SelectedIndex = KeyYYYY - MinYear + 1;

                int KeyMM = DateTime.Parse(dr["授權填表日期"].ToString()).Month; SelKeyMM.SelectedIndex = KeyMM;
                int KeyDD = DateTime.Parse(dr["授權填表日期"].ToString()).Day; SelKeyDD.SelectedIndex = KeyDD;
            }
            else
            {
                SelKeyYYYY.SelectedIndex = 0; SelKeyMM.SelectedIndex = 0; SelKeyDD.SelectedIndex = 0;
            }

            TextKeyer.Text = dr["授權填表人員"].ToString();
            TextMemo.Text = dr["授權備註說明"].ToString();

            if (!(dr["授權建立日期"] is DBNull)) LblCreateDT.Text = DateTime.Parse(dr["授權建立日期"].ToString()).ToString("yyyy/MM/dd HH:mm");
            LblCreater.Text = dr["授權建立人員"].ToString();
            if (!(dr["授權修改日期"] is DBNull)) LblUpdateDT.Text = DateTime.Parse(dr["授權修改日期"].ToString()).ToString("yyyy/MM/dd HH:mm");
            LblUpdater.Text = dr["授權修改人員"].ToString();

            lblChecks.Text = GetChecks("軟體授權", TextAskNo.Text);   //取得資料查核
        }

        cmd.Cancel(); cmd.Dispose(); dr.Close(); Conn.Close(); Conn.Dispose();
    }

    protected void ReadHelp() //讀取欄位說明
    {
        SqlCommand sqlQuery = new SqlCommand();
        sqlQuery.CommandText = "select Item,Memo,Config from Config where Kind='軟體授權' order by Mark";
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
            Msg.Text = "<script>alert('您不是軟體小組的成員，沒有新增授權資料的權限！');</script>";
        }
        else if (TextSwNo.Text == "" | TextSwNo.Text == "0")
        {
            Msg.Text = "<script>alert('您未設定軟體編號！');</script>";
        }
        else
        {
            if (TextSwNo.Text == "") TextSwNo.Text = "0";
            if (TextApNo.Text == "") TextApNo.Text = "0";
            if (TextIdentify.Text == "") TextIdentify.Text = "0";

            LblCreateDT.Text = DateTime.Now.ToString("yyyy/MM/dd HH:mm");
            LblCreater.Text = Session["UserName"].ToString();
            LblUpdateDT.Text = LblCreateDT.Text;
            LblUpdater.Text = LblCreater.Text;

            string AskNo = GetPKNo("授權編號", "軟體授權").ToString();
            TextAskNo.Text = AskNo; //新增完成後，要賦予新取得之授權編號 
            if (TextIdentify.Text == "" | TextIdentify.Text == "0") TextIdentify.Text = AskNo;

            string SQL = GetInsSQL(AskNo);
            //Response.Write(SQL);
            //Response.End();
            ExecDbSQL(SQL);
            InsLifeLog(SQL);

            Msg.Text = "<script>alert('新增授權資料[" + LblSwName.Text + "] - [" + TextHostLcs.Text + "] 完成！');window.open('AskEdit.aspx?AskNo=" + TextAskNo.Text + "','_self');</script>";            
        }

        Page.Controls.Add(Msg);
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

    protected void BtnDel_Click(object sender, EventArgs e)    //真刪除，請注意生命履歷要詳細記錄所有所刪資料之詳細所有欄位資料
    {
        Literal Msg = new Literal();

        if (!RightCheck())
        {
            Msg.Text = "<script>alert('您不是軟體小組的成員，沒有刪除授權資料的權限！');</script>";
        }
        else
        {
            string SQL = "delete from [軟體授權] where [授權編號]=" + TextAskNo.Text;
            ExecDbSQL(SQL);

            InsLifeLog("刪除 [" + LblSwName.Text + "] - [" + TextHostLcs.Text + "] ： " + SQL + "，原本軟體授權SQL ： " + GetInsSQL(TextAskNo.Text));
            ExecDbSQL("delete from [申請表單] where [表單種類]='軟體申請' and [主鍵編號]=" + TextAskNo.Text);

            Msg.Text = "<script>alert('完成刪除 [" + LblSwName.Text + "] - [" + TextHostLcs.Text + "] 資料！');window.close();</script>";
        }

        Page.Controls.Add(Msg);
    }

    protected void BtnEdit_Click(object sender, EventArgs e) //按下修改按鈕
    {
        Literal Msg = new Literal();

        if (!RightCheck())
        {
            Msg.Text = "<script>alert('您不是軟體小組的成員，沒有異動授權資料的權限！');</script>";
        }
        else if (TextSwNo.Text == "" | TextSwNo.Text == "0")
        {
            Msg.Text = "<script>alert('您未設定軟體編號！');</script>";
        }
        else
        {
            if (TextSwNo.Text == "") TextSwNo.Text = "0";
            if (TextApNo.Text == "") TextApNo.Text = "0";
            if (TextIdentify.Text == "" | TextIdentify.Text == "0") TextIdentify.Text = TextAskNo.Text;
            LblUpdateDT.Text = DateTime.Now.ToString("yyyy/MM/dd HH:mm");
            LblUpdater.Text = Session["UserName"].ToString();

            InsLifeLog("修改 [" + LblSwName.Text + "] - [" + TextHostLcs.Text + "] ：： " + GetUpdate(int.Parse(TextAskNo.Text), "Life"));
            ExecDbSQL("Update [軟體授權] set " + GetUpdate(int.Parse(TextAskNo.Text), "SQL") + " where [授權編號]=" + TextAskNo.Text);

            Msg.Text = "<script>alert('更新資料[" + LblSwName.Text + "] - [" + TextHostLcs.Text + "] 完成！');</script>";
            ReadAsk();
        }

        Page.Controls.Add(Msg);
    }    

    protected string GetYMD(DropDownList SelY, DropDownList SelM, DropDownList SelD, string pm)   //取得年月日
    {
        string YMD = SelY.SelectedValue + "/" + SelM.SelectedValue + "/" + SelD.SelectedValue;

        if (pm == "p")
        {
            if (SelY.SelectedValue == EverYear.ToString()) return ("'" + EverYear.ToString() + "/12/31'");
            else if (YMD.Length == 10) return ("'" + YMD + "'");
            else return ("null");
        }
        else
        {
            if (SelY.SelectedValue == EverYear.ToString()) return (EverYear.ToString() + "/12/31");
            else if (YMD.Length == 10) return (YMD);
            else return ("null");
        }
    }

    protected string GetInsSQL(string AskNo) //取得新增資料的語法
    {
        string Identify = TextIdentify.Text;
        if (Identify == "" | Identify == "0") Identify = AskNo;

        return ("insert into [軟體授權] values("
            + AskNo + ","
            + TextSwNo.Text + ","
            + TextApNo.Text + ","
            + Identify + ",'"
            + TextVersAsk.Text + "','"
            + SelLcsAsk.SelectedValue + "','"
            + TextLcsMemo.Text + "','"
            + TextSnAsk.Text + "','"
            + SelUnitLcs.SelectedValue + "','"
            + TextUserLcs.Text + "','"
            + TextHostLcs.Text + "','"
            + TextIPLcs.Text + "','"
            + TextPropNoALcs.Text + "','"
            + TextPropNoBLcs.Text + "','"
            + TextBrandLcs.Text + "','"
            + TextStyleLcs.Text + "','"
            + TextAttach.Text + "','"
            + SelStatus.SelectedValue + "','"
            + TextCause.Text + "','"
            + TextExec.Text + "',"
            + GetYMD(SelStatusYYYY, SelStatusMM, SelStatusDD, "p") + ","
            + GetYMD(SelKeyYYYY, SelKeyMM, SelKeyDD, "p") + ",'"
            + TextKeyer.Text + "','"
            + LblCreateDT.Text + "','"
            + LblCreater.Text + "','"
            + LblUpdateDT.Text + "','"
            + LblUpdater.Text + "','"
            + TextMemo.Text + "'"
            + ")");
    }

    protected void BtnPre_Click(object sender, EventArgs e) //取得預覽值
    {
        Label1.Text = "表單編號/申請事項 -- " + TextFormNo.Text + " / " + TextAskFormMemo.Text + "<br />"
            + "授權編號 -- " + TextAskNo.Text + "<br />"
            + "軟體編號 -- " + TextSwNo.Text + "<br />"
            + "作業編號 -- " + TextApNo.Text + "<br />"
            + "授權識別 -- " + TextIdentify.Text + "<br />"
            + "申請版本 -- " + TextVersAsk.Text + "<br />"
            + "申請授權 -- " + SelLcsAsk.SelectedValue + "<br />"
            + "授權說明 -- " + TextLcsMemo.Text + "<br />"
            + "申請序號 -- " + TextSnAsk.Text + "<br />"
            + "授權單位 -- " + SelUnitLcs.SelectedValue + "<br />"
            + "授權人員 -- " + TextUserLcs.Text + "<br />"
            + "授權主機 -- " + TextHostLcs.Text + "<br />"
            + "授權IP -- " + TextIPLcs.Text + "<br />"
            + "授權財編A -- " + TextPropNoALcs.Text + "<br />"
            + "授權財編B -- " + TextPropNoBLcs.Text + "<br />"
            + "授權廠牌 -- " + TextBrandLcs.Text + "<br />"
            + "授權型式 -- " + TextStyleLcs.Text + "<br />"
            + "授權附件 -- " + TextAttach.Text + "<br />"
            + "授權狀態 -- " + SelStatus.SelectedValue + "<br />"
            + "減損原因 -- " + TextCause.Text + "<br />"
            + "減損方式 -- " + TextExec.Text + "<br />"
            + "減損日期 -- " + GetYMD(SelStatusYYYY, SelStatusMM, SelStatusDD, "m") + "<br />"
            + "填表日期 -- " + GetYMD(SelKeyYYYY, SelKeyMM, SelKeyDD, "m") + "<br />"
            + "填表人員 -- " + TextKeyer.Text + "<br />"
            + "建立日期 -- " + LblCreateDT.Text + "<br />"
            + "建立人員 -- " + LblCreater.Text + "<br />"
            + "修改日期 -- " + LblUpdateDT.Text + "<br />"
            + "修改人員 -- " + LblUpdater.Text + "<br />"
            + "備註說明 -- " + TextMemo.Text + "<br />";
    }

    protected string GetUpdate(int AskNo, string SQLorLife) //取得修改資料的語法
    {
        string SQL = "";
        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
        Conn.Open();
        SqlCommand cmd = new SqlCommand("select * from [軟體授權] where [授權編號]=" + AskNo, Conn);
        SqlDataReader dr = null;
        string SaveDT = "",StatusDay="",KeyDay="";
        dr = cmd.ExecuteReader();
        if (dr.Read())
        {
            if (!(dr["修改日期"] is DBNull)) SaveDT = DateTime.Parse(dr["修改日期"].ToString()).ToString("yyyy/MM/dd HH:mm");
            if (!(dr["減損日期"] is DBNull)) StatusDay = DateTime.Parse(dr["減損日期"].ToString()).ToString("yyyy/MM/dd");
            if (!(dr["填表日期"] is DBNull)) KeyDay = DateTime.Parse(dr["填表日期"].ToString()).ToString("yyyy/MM/dd");

            SQL = GetUpdateCol("軟體編號", dr["軟體編號"].ToString(), TextSwNo.Text, "integer", SQLorLife)
                + GetUpdateCol("作業編號", dr["作業編號"].ToString(), TextApNo.Text, "integer", SQLorLife)
                + GetUpdateCol("授權識別", dr["授權識別"].ToString(), TextIdentify.Text, "integer", SQLorLife)
                + GetUpdateCol("申請版本", dr["申請版本"].ToString(), TextVersAsk.Text, "string", SQLorLife)
                + GetUpdateCol("申請授權", dr["申請授權"].ToString(), SelLcsAsk.SelectedValue, "string", SQLorLife)
                + GetUpdateCol("授權說明", dr["授權說明"].ToString(), TextLcsMemo.Text, "string", SQLorLife)
                + GetUpdateCol("申請序號", dr["申請序號"].ToString(), TextSnAsk.Text, "string", SQLorLife)
                + GetUpdateCol("授權單位", dr["授權單位"].ToString(), SelUnitLcs.SelectedValue, "string", SQLorLife)
                + GetUpdateCol("授權人員", dr["授權人員"].ToString(), TextUserLcs.Text, "string", SQLorLife)
                + GetUpdateCol("授權主機", dr["授權主機"].ToString(), TextHostLcs.Text, "string", SQLorLife)
                + GetUpdateCol("授權IP", dr["授權IP"].ToString(), TextIPLcs.Text, "string", SQLorLife)
                + GetUpdateCol("授權財編A", dr["授權財編A"].ToString(), TextPropNoALcs.Text, "string", SQLorLife)
                + GetUpdateCol("授權財編B", dr["授權財編B"].ToString(), TextPropNoBLcs.Text, "string", SQLorLife)
                + GetUpdateCol("授權廠牌", dr["授權廠牌"].ToString(), TextBrandLcs.Text, "string", SQLorLife)
                + GetUpdateCol("授權型式", dr["授權型式"].ToString(), TextStyleLcs.Text, "string", SQLorLife)
                + GetUpdateCol("授權附件", dr["授權附件"].ToString(), TextAttach.Text, "string", SQLorLife)
                + GetUpdateCol("授權狀態", dr["授權狀態"].ToString(), SelStatus.SelectedValue, "string", SQLorLife)
                + GetUpdateCol("減損原因", dr["減損原因"].ToString(), TextCause.Text, "string", SQLorLife)
                + GetUpdateCol("減損方式", dr["減損方式"].ToString(), TextExec.Text, "string", SQLorLife)
                + GetUpdateCol("減損日期", StatusDay, GetYMD(SelStatusYYYY, SelStatusMM, SelStatusDD, "m"), "date", SQLorLife)
                + GetUpdateCol("填表日期", KeyDay, GetYMD(SelKeyYYYY, SelKeyMM, SelKeyDD, "m"), "date", SQLorLife)
                + GetUpdateCol("填表人員", dr["填表人員"].ToString(), TextKeyer.Text, "string", SQLorLife)
                + GetUpdateCol("修改日期", SaveDT, DateTime.Now.ToString("yyyy/MM/dd HH:mm"), "datetime", SQLorLife)
                + GetUpdateCol("修改人員", dr["修改人員"].ToString(), Session["UserName"].ToString(), "string", SQLorLife)
                + GetUpdateCol("備註說明", dr["備註說明"].ToString(), TextMemo.Text, "string", SQLorLife);

            if (SQL != "") SQL = SQL.Substring(1);
        }
        cmd.Cancel(); cmd.Dispose(); dr.Close(); Conn.Close(); Conn.Dispose();

        return (SQL);
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
                if (Source == "" | Source == "null") Source = "(空白)";
                if (Target == "" | Target == "null") Target = "(空白)";
                if (Source != Target) SQL = SQL + ",[" + ColName + "]：" + Source + " -> " + Target;
            }
        }
        return (SQL);
    }

    protected void InsLifeLog(string SQL) //寫入生命履歷
    {
        string LifeNo = GetPKNo("履歷編號", "生命履歷").ToString(); //履歷編號
        string TblName = "軟體授權";    //表格名稱
        string PKno = TextAskNo.Text;   //主鍵編號
        string Mt = "軟體小組";    //維護人員
        string UN = Session["UserName"].ToString();   //登入的UserName：異動人員
        string LiftDT = DateTime.Now.ToString("yyyy/MM/dd HH:mm");  //異動日期
        string LifeIP = Request.ServerVariables["REMOTE_ADDR"].ToString();

        ExecDbSQL("Insert into [生命履歷] values(" + LifeNo + ",'" + TblName + "'," + PKno + ",'" + SQL.Replace("'", "''") + "','" + Mt + "','','','" + UN + "','" + LiftDT + "','" + LifeIP + "')");
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

    protected Boolean RightCheck() //是否有權修改資料
    {
        if (InGroup(Session["UserName"].ToString(),"軟體小組") | InGroup(Session["UserName"].ToString(),"SSM小組")) return (true);
        else return (false);
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
            SqlCommand cmd = new SqlCommand("select * from Config where (Kind='" + ChkUnit + "') and Item='" + ChkName + "'", Conn);
            SqlDataReader dr = null;
            dr = cmd.ExecuteReader();
            if (dr.Read()) TF=true;
            cmd.Cancel(); cmd.Dispose(); dr.Close(); Conn.Close(); Conn.Dispose();
        }

        return (TF);
    }

    protected void OpenExecWindow(string strJavascript)    //選取後就另開視窗
    {
        if (ScriptManager.GetCurrent(Page) == null)
        {
            Page.ClientScript.RegisterStartupScript(Page.GetType(), "AskEdit", strJavascript, true);
        }
        else
        {
            ScriptManager.RegisterStartupScript(Page, Page.GetType(), "AskEdit", strJavascript, true);
        }
    }

    protected void SelCause_SelectedIndexChanged(object sender, EventArgs e)  //選擇減損原因
    {
        DropDownList Sel = (DropDownList)form1.FindControl("SelCause");    //放入Table物件後，必須用FindControl才能取值
        TextCause.Text = Sel.SelectedValue;
    }

    protected void SelExec_SelectedIndexChanged(object sender, EventArgs e)  //選擇減損方式
    {
        DropDownList Sel = (DropDownList)form1.FindControl("SelExec");    //放入Table物件後，必須用FindControl才能取值
        TextExec.Text = Sel.SelectedValue;
    }

    protected void LinkSwEdit_Click(object sender, EventArgs e)  //編輯Sw By SwNo
    {
        Literal Msg = new Literal();

        if (TextSwNo.Text == "0" | TextSwNo.Text == "")
        {
            Msg.Text = "<script>alert('您尚未輸入軟體編號！');</script>";
            Page.Controls.Add(Msg);
        }
        else OpenExecWindow("window.open('SwEdit.aspx?SwNo=" + TextSwNo.Text + "','_self');");
    }
    protected void LinkFormEdit_Click(object sender, EventArgs e)  //編輯Sw By FormNo
    {
        Literal Msg = new Literal();

        if (TextFormNo.Text.Length != 9)
        {
            Msg.Text = "<script>alert('您尚未輸入表單編號！');</script>";
            Page.Controls.Add(Msg);
        }
        else
        {
            SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
            Conn.Open();
            SqlCommand cmd = new SqlCommand("select [軟體編號] from [軟體主檔] where [表單編號]='" + TextFormNo.Text + "'", Conn);
            SqlDataReader dr = null;
            dr = cmd.ExecuteReader();

            if (dr.Read())
            {

                OpenExecWindow("window.open('SwEdit.aspx?SwNo=" + dr[0].ToString() + "','_self');");
            }
            else
            {
                Msg.Text = "<script>alert('查無此表單編號(" + TextFormNo.Text + ")！');</script>";
                Page.Controls.Add(Msg);
            }
            cmd.Cancel(); cmd.Dispose(); dr.Close(); Conn.Close(); Conn.Dispose();
        }
    }
    protected void LinkSwIn_Click(object sender, EventArgs e)  //帶入Sw By SwNo
    {
        Sw_In("[軟體編號]=" + TextSwNo.Text);
    }
    protected void LinkFormIn_Click(object sender, EventArgs e)  //帶入Sw By FormNo
    {
        Literal Msg = new Literal();

        if (TextFormNo.Text.Length != 9)
        {
            Msg.Text = "<script>alert('您尚未輸入表單編號！');</script>";
            Page.Controls.Add(Msg);
        }
        else Sw_In("[表單編號]='" + TextFormNo.Text + "'");
    }
    protected void Sw_In(string PK)  //帶入Sw By SwNo or FormNo
    {
        Literal Msg = new Literal();

        if (PK == "[軟體編號]=0" | PK == "[軟體編號]=")
        {
            ClearSw();
            Msg.Text = "<script>alert('您尚未輸入軟體或表單編號，軟體資料已清除！');</script>";
        }
        else
        {
            SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
            Conn.Open();
            SqlCommand cmd = new SqlCommand("select [軟體名稱],[購買版本],[可用版本],[授權方式],[購買序號],[軟體附件],[軟體編號],[表單編號],[軟體狀態] from [軟體主檔] where " + PK, Conn);
            SqlDataReader dr = null;
            dr = cmd.ExecuteReader();
            if (dr.Read())
            {
                lblSwStatus.Text = dr["軟體狀態"].ToString();
                LblSwName.Text = dr["軟體名稱"].ToString() + " " + dr["購買版本"].ToString() + " " + dr["授權方式"].ToString();

                lblVersBuy.Text = dr["購買版本"].ToString(); SelVersAsk.DataBind();
                lblVersCan.Text = dr["可用版本"].ToString();
                lblLcsBuy.Text = dr["授權方式"].ToString();

                for (int i = 0; i < SelVersAsk.Items.Count; i++) if (SelVersAsk.Items[i].Value == dr["購買版本"].ToString()) SelVersAsk.SelectedIndex = i;
                TextVersAsk.Text = dr["購買版本"].ToString();
                for (int i = 0; i < SelLcsAsk.Items.Count; i++) if (SelLcsAsk.Items[i].Value == dr["授權方式"].ToString()) SelLcsAsk.SelectedIndex = i;

                if (RightCheck())   //有無權利看到序號
                {
                    lblSnBuy.Text = dr["購買序號"].ToString();
                    TextSnAsk.Text = dr["購買序號"].ToString();
                }

                lblAttach.Text = dr["軟體附件"].ToString();

                TextSwNo.Text = dr["軟體編號"].ToString();
                TextFormNo.Text = dr["表單編號"].ToString();

                Msg.Text = "<script>alert('軟體資料已帶入！');</script>";
            }
            else
            {
                Msg.Text = "<script>alert('查無此軟體保管單！');</script>";
            }
            cmd.Cancel(); cmd.Dispose(); dr.Close(); Conn.Close(); Conn.Dispose();
        }
        Page.Controls.Add(Msg);
    }
    protected void ClearSw()  //清空Sw
    {
        
        TextSwNo.Text = "0";
        TextFormNo.Text = "";
        lblSwStatus.Text = "";
        LblSwName.Text = "";

        lblVersBuy.Text = ""; 
        lblVersCan.Text = "";
        lblLcsBuy.Text = "";

        lblSnBuy.Text = "";
        lblAttach.Text = "";  
    }

    protected void LinkApIn_Click(object sender, EventArgs e)  //帶入AP
    {
        Literal Msg = new Literal();

        if (TextApNo.Text == "0" | TextApNo.Text == "")
        {
            ClearAsk();
            Msg.Text = "<script>alert('您尚未選取要授權的作業主機，申請資料已清除！');</script>";
        }
        else
        {
            SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
            Conn.Open();
            SqlCommand cmd = new SqlCommand("select * from [View_通用設備] where [作業編號]=" + TextApNo.Text, Conn);
            SqlDataReader dr = null;
            dr = cmd.ExecuteReader();
            if (dr.Read())
            {
                LblApNo.Text = dr["系統全名"].ToString() + "-" + dr["主機名稱"].ToString();
                
                for (int i = 0; i < SelUnitLcs.Items.Count; i++) if (SelUnitLcs.Items[i].Value == dr["使用課別"].ToString()) SelUnitLcs.SelectedIndex = i;
                for (int i = 0; i < SelUserLcs.Items.Count; i++) if (SelUserLcs.Items[i].Value == dr["使用人員"].ToString()) SelUserLcs.SelectedIndex = i;  //本行其實無作用
                TextUserLcs.Text = dr["使用人員"].ToString();
                TextHostLcs.Text = dr["主機名稱"].ToString();
                TextIPLcs.Text = dr["IP"].ToString();
                TextPropNoALcs.Text = dr["財產編號A"].ToString();
                TextPropNoBLcs.Text = dr["財產編號B"].ToString();
                TextBrandLcs.Text = dr["廠牌"].ToString();
                TextStyleLcs.Text = dr["型式"].ToString();

                lblUnitUse.Text = dr["使用課別"].ToString();
                lblUserUse.Text = dr["使用人員"].ToString();
                lblHostUse.Text = dr["主機名稱"].ToString();
                lblIPUse.Text = dr["IP"].ToString();
                lblPropNoAUse.Text = dr["財產編號A"].ToString();
                lblPropNoBUse.Text = dr["財產編號B"].ToString();
                lblBrandUse.Text = dr["廠牌"].ToString();
                lblStyleUse.Text = dr["型式"].ToString();

                Msg.Text = "<script>alert('主機資料已帶入！');</script>";
            }
            else
            {
                Msg.Text = "<script>alert('無此作業編號(" + TextApNo.Text + ")！');</script>";
            }
            cmd.Cancel(); cmd.Dispose(); dr.Close(); Conn.Close(); Conn.Dispose();
        }
        Page.Controls.Add(Msg);
    }
    protected void LinkApClear_Click(object sender, EventArgs e)  //清除AP
    {
        ClearAsk();
    }
    protected void ClearAsk()  //當授權狀態為可申請或已減損時，作業編號、申請單位、申請人員、主機名稱、IP...自動清空
    {
        TextApNo.Text = "0"; LblApNo.Text = ""; LblApNo.Text = "(尚未設定)";
        SelUnitLcs.SelectedIndex = 0; lblUnitUse.Text = "";
        TextUserLcs.Text = ""; lblUserUse.Text = "";
        TextHostLcs.Text = ""; lblHostUse.Text = "";
        TextIPLcs.Text = ""; lblIPUse.Text = "";
        TextPropNoALcs.Text = ""; lblPropNoAUse.Text = "";
        TextPropNoBLcs.Text = ""; lblPropNoBUse.Text = "";
        TextBrandLcs.Text = ""; lblBrandUse.Text = "";
        TextStyleLcs.Text = ""; lblStyleUse.Text = "";
    }

    protected void SelAskFormNo_SelectedIndexChanged(object sender, EventArgs e)   //選擇申請表單
    {
        if (TextAskNo.Text != "")
        {
            ListBox Sel = (ListBox)form1.FindControl("SelAskFormNo");    //放入Table物件後，必須用FindControl才能取值

            SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
            Conn.Open();
            SqlCommand cmd = new SqlCommand("select * from [申請表單] where [表單種類]='軟體申請' and [主鍵編號]=" + TextAskNo.Text + " and [表單編號]='" + Sel.SelectedValue + "'", Conn);
            SqlDataReader dr = null;
            dr = cmd.ExecuteReader();
            if (dr.Read())
            {
                TextAskFormNo.Text = dr["表單編號"].ToString();
                TextAskFormMemo.Text = dr["申請事項"].ToString();

                if (!(dr["填表日期"] is DBNull))
                {
                    int FormYYYY = DateTime.Parse(dr["填表日期"].ToString()).Year;
                    if (FormYYYY - MinYear + 1 > SelFormYYYY.Items.Count) SelFormYYYY.SelectedIndex = SelFormYYYY.Items.Count - 1;
                    else SelFormYYYY.SelectedIndex = FormYYYY - MinYear + 1;

                    int FormMM = DateTime.Parse(dr["填表日期"].ToString()).Month; SelFormMM.SelectedIndex = FormMM;
                    int FormDD = DateTime.Parse(dr["填表日期"].ToString()).Day; SelFormDD.SelectedIndex = FormDD;
                }
                else
                {
                    SelFormYYYY.SelectedIndex = 0; SelFormMM.SelectedIndex = 0; SelFormDD.SelectedIndex = 0;
                }
            }
            cmd.Cancel(); cmd.Dispose(); dr.Close(); Conn.Close(); Conn.Dispose();
        }
    }
    protected void LinkAskFormAdd_Click(object sender, EventArgs e)  //新增申請表單
    {
        ListBox Sel = (ListBox)form1.FindControl("SelAskFormNo");    //放入Table物件後，必須用FindControl才能取值

        if (TextAskFormNo.Text != "" & TextAskNo.Text != "" & TextAskNo.Text != "0" & TextAskFormMemo.Text != "" & SelFormYYYY.SelectedValue != "" & SelFormMM.SelectedValue != "" & SelFormDD.SelectedValue != "")
        {
            InsLifeLog("新增 [" + LblSwName.Text + "] - [" + TextHostLcs.Text + "] 之[申請表單]： " + TextAskFormNo.Text + " " + TextAskFormMemo.Text + " " + SelFormYYYY.SelectedValue + "/" + SelFormMM.SelectedValue + "/" + SelFormDD.SelectedValue);
            InsAskForm(false);
            Sel.DataBind();

            AddMsg("完成新增申請單(" + TextAskFormNo.Text + ")！");
        }
        else AddMsg("您申請單資料不完整或尚未新增授權資料！");
    }
    protected void LinkAskFormDel_Click(object sender, EventArgs e)  //刪除申請表單
    {
        ListBox Sel = (ListBox)form1.FindControl("SelAskFormNo");    //放入Table物件後，必須用FindControl才能取值

        if (Sel.SelectedValue != "")
        {
            InsLifeLog("刪除 [" + LblSwName.Text + "] - [" + TextHostLcs.Text + "] 之[申請表單]： " + Sel.SelectedItem.Text + " " + SelFormYYYY.SelectedValue + "/" + SelFormMM.SelectedValue + "/" + SelFormDD.SelectedValue);
            ExecDbSQL("delete from [申請表單] where [表單種類]='軟體申請' and [主鍵編號]=" + TextAskNo.Text + " and [表單編號]='" + Sel.SelectedValue + "'");            
            Sel.DataBind();
            ClearForm();

            AddMsg("完成刪除申請單(" + Sel.SelectedValue + ")！");
        }
        else AddMsg("您尚未選取申請單！");
    }
    protected void InsAskForm(Boolean DelYN)   //新增申請表單
    {
        if (DelYN) ExecDbSQL("delete from [申請表單] where [表單種類]='軟體申請' and [主鍵編號]=" + TextAskNo.Text + " and [表單編號]='" + SelAskFormNo.SelectedValue + "'");
        ExecDbSQL("insert into [申請表單] values('軟體申請'," + TextAskNo.Text + ",'" + TextAskFormNo.Text + "','" + TextAskFormMemo.Text + "','" + SelFormYYYY.SelectedValue + "/" + SelFormMM.SelectedValue + "/" + SelFormDD.SelectedValue + "')");
        SelAskFormNo.DataBind();
        ClearForm();
    }
    protected void ClearForm()  //清空申請單
    {
        TextAskFormNo.Text = ""; TextAskFormMemo.Text = ""; SelFormYYYY.SelectedIndex = 0; SelFormMM.SelectedIndex = 0; SelFormDD.SelectedIndex = 0;
    }

    protected void SelVersAsk_SelectedIndexChanged(object sender, EventArgs e)   //選擇軟體版本
    {
        DropDownList Sel = (DropDownList)form1.FindControl("SelVersAsk");    //放入Table物件後，必須用FindControl才能取值
        TextVersAsk.Text = Sel.SelectedValue;
    }

    protected void SelUnitLcs_SelectedIndexChanged(object sender, EventArgs e)  //選擇授權單位
    {
        SelUserLcs.DataBind();
        TextUserLcs.Text = SelUserLcs.Items[0].Value;
    }
    protected void SelUserLcs_SelectedIndexChanged(object sender, EventArgs e)   //選擇授權人員
    {
        DropDownList Sel = (DropDownList)form1.FindControl("SelUserLcs");    //放入Table物件後，必須用FindControl才能取值
        TextUserLcs.Text = Sel.SelectedValue;
    }

    protected void SelKeyerUnit_SelectedIndexChanged(object sender, EventArgs e)  //選擇填表單位
    {
        SelKeyer.DataBind();
        TextKeyer.Text = SelKeyer.Items[0].Value;
    }
    protected void SelKeyer_SelectedIndexChanged(object sender, EventArgs e)   //選擇填表人員
    {
        DropDownList Sel = (DropDownList)form1.FindControl("SelKeyer");    //放入Table物件後，必須用FindControl才能取值
        TextKeyer.Text = Sel.SelectedValue;
    }

    protected void SelBrandLcs_SelectedIndexChanged(object sender, EventArgs e)  //選擇授權廠牌
    {
        DropDownList Sel = (DropDownList)form1.FindControl("SelBrandLcs");    //放入Table物件後，必須用FindControl才能取值
        TextBrandLcs.Text = Sel.SelectedValue;
        SelStyleLcs.DataBind();
        TextStyleLcs.Text = SelStyleLcs.Items[0].Value;
    }
    protected void SelStyleLcs_SelectedIndexChanged(object sender, EventArgs e)  //選擇授權型式
    {
        DropDownList Sel = (DropDownList)form1.FindControl("SelStyleLcs");    //放入Table物件後，必須用FindControl才能取值
        TextStyleLcs.Text = Sel.SelectedValue;
    }
    
    protected void SelStatus_SelectedIndexChanged(object sender, EventArgs e)  //狀態改變
    {
        if (SelStatus.SelectedValue == "可申請" | SelStatus.SelectedValue == "已減損") ClearAsk();
    }

    protected void BtnLife_Click(object sender, EventArgs e) //查詢生命履歷
    {
        Session["LifeSQL"] = "select * from [生命履歷] where [表格名稱]='軟體授權' and [主鍵編號]=" + TextAskNo.Text + " order by [履歷編號] desc";
        OpenExecWindow("window.open('../Life/LifeLog.aspx?Search=AskEdit&Tbl=軟體授權&PK=" + TextAskNo.Text + "','_self');");
    }

    protected void AddMsg(string MsgText)
    {
        Literal Msg = new Literal();
        Msg.Text = "<script>alert('" + MsgText + "');</script>";
        Page.Controls.Add(Msg);
    }

    protected void BtnPrint_Click(object sender, EventArgs e) //列印申請單
    {
        OpenExecWindow("window.open('AskForm.aspx?AskNo=" + TextAskNo.Text + "','_blank');");
    }

    protected void LinkIdentify_Click(object sender, EventArgs e)  //識別查詢
    {
        if (TextIdentify.Text != "" & TextIdentify.Text != "0")
        {
            Session["AskSQL"] = "SELECT * FROM [View_軟體管理] WHERE [軟體編號]=" + TextSwNo.Text + " and [授權識別]=" + TextIdentify.Text;
            OpenExecWindow("window.open('Ask.aspx?Search=AskEdit','_self');");
        }
        else
        {
            Literal Msg = new Literal();
            Msg.Text = "<script>alert('請先設定授權識別號碼！');</script>";
            Page.Controls.Add(Msg);
        }
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
}