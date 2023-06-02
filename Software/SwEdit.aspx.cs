using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.Configuration;
using System.Data;
using System.Data.SqlClient;

public partial class Software_SwEdit : System.Web.UI.Page
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
                TextSwNo.Text = Request["SwNo"];

                if (TextSwNo.Text == "")    //各按鈕啟用與否
                {
                    BtnEdit.Enabled = false; BtnDel.Enabled = false;
                }
                else
                {
                    BtnEdit.Enabled = true; BtnDel.Enabled = true;
                }

                //顯示各下拉式選單(DB)，注意讀值先後順序
                SelSwName.DataBind(); SelVersBuy.DataBind(); SelVersCan.DataBind(); SelLcsKind.DataBind(); SelSource.DataBind(); SelBrand.DataBind();
                SelStyle.DataBind(); SelLife.DataBind(); SelSupply.DataBind(); SelCause.DataBind(); SelExec.DataBind(); SelUnit.DataBind(); SelAsker.DataBind();

                if (TextSwNo.Text != "") ReadSw();        //讀取軟體主檔 

                GetLicense();    //於Label顯示各種申請授權方式統計
                GetStatus();    //於Label顯示各種授權狀態統計    
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
            AddSel(SelFormYYYY, MinYear, NowYear + OverYear); SelFormYYYY.SelectedIndex = NowYear - MinYear + 1;
            AddSel(SelBuyYYYY, MinYear, NowYear + OverYear); AddSel(SelBuyMM, 1, 12); AddSel(SelBuyDD, 1, 31);
            AddSel(SelUseYYYY, MinYear, NowYear + OverYear); AddSel(SelUseMM, 1, 12); AddSel(SelUseDD, 1, 31); AddYearSel(SelUseYYYY);
            AddSel(SelUpYYYY, MinYear, NowYear + OverYear); AddSel(SelUpMM, 1, 12); AddSel(SelUpDD, 1, 31); AddYearSel(SelUpYYYY);
            AddSel(SelStatusYYYY, MinYear, NowYear + OverYear); AddSel(SelStatusMM, 1, 12); AddSel(SelStatusDD, 1, 31);
            AddSel(SelKeyYYYY, MinYear, NowYear + OverYear); AddSel(SelKeyMM, 1, 12); AddSel(SelKeyDD, 1, 31);

            //填入預設值
            TextLcsNum.Text = "0"; TextPrice.Text = "0"; TextTotal.Text = "0";
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

    protected void GetStatus()   //授權狀態統計
    {
        lblStatus.Text = "";

        if (TextSwNo.Text != "")
        {
            SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
            Conn.Open();
            SqlCommand cmd = new SqlCommand("select [授權狀態],count(*) FROM [軟體授權] WHERE [軟體編號]=" + TextSwNo.Text + " group by [授權狀態] order by [授權狀態]", Conn);
            SqlDataReader dr = null;
            dr = cmd.ExecuteReader();

            while (dr.Read())
            {
                lblStatus.Text = lblStatus.Text + dr[0].ToString() + "：" + dr[1].ToString() + "　";
            }

            cmd.Cancel(); cmd.Dispose(); dr.Close(); Conn.Close(); Conn.Dispose();
        }
    }

    protected void GetLicense()   //申請授權方式統計
    {
        lblLicense.Text = "";

        if (TextSwNo.Text != "")
        {
            SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
            Conn.Open();
            SqlCommand cmd = new SqlCommand("select [申請授權],count(*) FROM [軟體授權] WHERE [軟體編號]=" + TextSwNo.Text + " group by [申請授權] order by [申請授權]", Conn);
            SqlDataReader dr = null;
            dr = cmd.ExecuteReader();

            string LcsKind = "";

            while (dr.Read())
            {
                LcsKind = dr[0].ToString(); if (LcsKind == "") LcsKind = "(無)";
                if (lblLicense.Text == "") lblLicense.Text = LcsKind + "*" + dr[1].ToString();
                else lblLicense.Text = lblLicense.Text + " + " + LcsKind + "*" + dr[1].ToString();
            }

            cmd.Cancel(); cmd.Dispose(); dr.Close(); Conn.Close(); Conn.Dispose();
        }
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

    protected void ReadSw()    //讀取軟體主檔
    {
        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
        Conn.Open();
        SqlCommand cmd = new SqlCommand("SELECT * FROM [軟體主檔] WHERE [軟體編號]=" + TextSwNo.Text, Conn);
        SqlDataReader dr = null;
        dr = cmd.ExecuteReader();

        if (dr.Read())
        {
            SelFormYYYY.SelectedIndex = int.Parse(dr["表單編號"].ToString().Substring(0, 4)) - MinYear + 1;
            TextFormXXXX.Text = dr["表單編號"].ToString().Substring(5);

            for (int i = 0; i < SelUnit.Items.Count; i++) if (SelUnit.Items[i].Value == dr["軟體單位"].ToString()) SelUnit.SelectedIndex = i;
            TextAsker.Text = dr["軟體人員"].ToString();

            for (int i = 0; i < SelSwName.Items.Count; i++) if (SelSwName.Items[i].Value == dr["軟體名稱"].ToString()) SelSwName.SelectedIndex = i;
            TextSwName.Text = dr["軟體名稱"].ToString();
            SelVersBuy.DataBind();  //連帶選項需要觸發
            for (int i = 0; i < SelVersBuy.Items.Count; i++) if (SelVersBuy.Items[i].Value == dr["購買版本"].ToString()) SelVersBuy.SelectedIndex = i;
            SelVersCan.DataBind();  //連帶選項需要觸發
            for (int i = 0; i < SelVersCan.Items.Count; i++) if (SelVersCan.Items[i].Value == dr["可用版本"].ToString()) SelVersCan.SelectedIndex = i;
            TextVersBuy.Text = dr["購買版本"].ToString();
            TextVersCan.Text = dr["可用版本"].ToString();

            for (int i = 0; i < SelLcsKind.Items.Count; i++) if (SelLcsKind.Items[i].Value == dr["授權方式"].ToString()) SelLcsKind.SelectedIndex = i;
            TextLcsNum.Text = dr["授權數量"].ToString();
            TextLcsMemo.Text = dr["授權說明"].ToString();

            for (int i = 0; i < SelSource.Items.Count; i++) if (SelSource.Items[i].Value == dr["軟體來源"].ToString()) SelSource.SelectedIndex = i;
            TextSource.Text = dr["軟體來源"].ToString();
            for (int i = 0; i < SelFunctions.Items.Count; i++) if (SelFunctions.Items[i].Value == dr["軟體功能"].ToString()) SelFunctions.SelectedIndex = i;
            TextFunctions.Text = dr["軟體功能"].ToString();

            for (int i = 0; i < SelBrand.Items.Count; i++) if (SelBrand.Items[i].Value == dr["適用廠牌"].ToString()) SelBrand.SelectedIndex = i;
            SelStyle.DataBind();  //連帶選項需要觸發
            for (int i = 0; i < SelStyle.Items.Count; i++) if (SelStyle.Items[i].Value == dr["適用型式"].ToString()) SelStyle.SelectedIndex = i;
            TextBrand.Text = dr["適用廠牌"].ToString();
            TextStyle.Text = dr["適用型式"].ToString();

            if (RightCheck())   //有無權利看到序號
            {
                TextSnBuy.Text = dr["購買序號"].ToString();
                TextSnDown.Text = dr["降級序號"].ToString();
            }

            for (int i = 0; i < SelLife.Items.Count; i++) if (SelLife.Items[i].Value == dr["期限說明"].ToString()) SelLife.SelectedIndex = i;
            TextLife.Text = dr["期限說明"].ToString();

            if (!(dr["購買日期"] is DBNull))
            {
                int BuyYYYY = DateTime.Parse(dr["購買日期"].ToString()).Year;
                if (BuyYYYY - MinYear + 1 > SelBuyYYYY.Items.Count) SelBuyYYYY.SelectedIndex = SelBuyYYYY.Items.Count - 1;
                else SelBuyYYYY.SelectedIndex = BuyYYYY - MinYear + 1;

                int BuyMM = DateTime.Parse(dr["購買日期"].ToString()).Month; SelBuyMM.SelectedIndex = BuyMM;
                int BuyDD = DateTime.Parse(dr["購買日期"].ToString()).Day; SelBuyDD.SelectedIndex = BuyDD;
            }
            else
            {
                SelBuyYYYY.SelectedIndex = 0; SelBuyMM.SelectedIndex = 0; SelBuyDD.SelectedIndex = 0;
            }
            if (!(dr["使用期限"] is DBNull))
            {
                int UseYYYY = DateTime.Parse(dr["使用期限"].ToString()).Year;
                if (UseYYYY - MinYear + 1 > SelUseYYYY.Items.Count) SelUseYYYY.SelectedIndex = SelUseYYYY.Items.Count - 1;
                else SelUseYYYY.SelectedIndex = UseYYYY - MinYear + 1;

                int UseMM = DateTime.Parse(dr["使用期限"].ToString()).Month; SelUseMM.SelectedIndex = UseMM;
                int UseDD = DateTime.Parse(dr["使用期限"].ToString()).Day; SelUseDD.SelectedIndex = UseDD;
            }
            else
            {
                SelUseYYYY.SelectedIndex = 0; SelUseMM.SelectedIndex = 0; SelUseDD.SelectedIndex = 0;
            }
            if (!(dr["更新期限"] is DBNull))
            {
                int UpYYYY = DateTime.Parse(dr["更新期限"].ToString()).Year;
                if (UpYYYY - MinYear + 1 > SelUpYYYY.Items.Count) SelUpYYYY.SelectedIndex = SelUpYYYY.Items.Count - 1;
                else SelUpYYYY.SelectedIndex = UpYYYY - MinYear + 1;

                int UpMM = DateTime.Parse(dr["更新期限"].ToString()).Month; SelUpMM.SelectedIndex = UpMM;
                int UpDD = DateTime.Parse(dr["更新期限"].ToString()).Day; SelUpDD.SelectedIndex = UpDD;
            }
            else
            {
                SelUpYYYY.SelectedIndex = 0; SelUpMM.SelectedIndex = 0; SelUpDD.SelectedIndex = 0;
            }

            TextPrice.Text = String.Format("{0:F0}", dr["軟體單價"]);
            TextTotal.Text = String.Format("{0:F0}", dr["軟體總價"]);
            TextPriceMemo.Text = dr["價格說明"].ToString();

            for (int i = 0; i < SelSupply.Items.Count; i++) if (SelSupply.Items[i].Value == dr["提供單位"].ToString()) SelSupply.SelectedIndex = i;
            TextSupply.Text = dr["提供單位"].ToString();

            TextMedia.Text = dr["存放媒體"].ToString();
            TextBookNo.Text = dr["圖書編號"].ToString();
            TextAttach.Text = dr["軟體附件"].ToString();

            for (int i = 0; i < SelStatus.Items.Count; i++) if (SelStatus.Items[i].Value == dr["軟體狀態"].ToString()) SelStatus.SelectedIndex = i;
            for (int i = 0; i < SelCause.Items.Count; i++) if (SelCause.Items[i].Value == dr["減損原因"].ToString()) SelCause.SelectedIndex = i;
            TextCause.Text = dr["減損原因"].ToString();
            for (int i = 0; i < SelExec.Items.Count; i++) if (SelExec.Items[i].Value == dr["減損方式"].ToString()) SelExec.SelectedIndex = i;
            TextExec.Text = dr["減損方式"].ToString();
            if (!(dr["減損日期"] is DBNull))
            {
                int StatusYYYY = DateTime.Parse(dr["減損日期"].ToString()).Year;
                if (StatusYYYY - MinYear + 1 > SelStatusYYYY.Items.Count) SelStatusYYYY.SelectedIndex = SelStatusYYYY.Items.Count - 1;
                else SelStatusYYYY.SelectedIndex = StatusYYYY - MinYear + 1;

                int StatusMM = DateTime.Parse(dr["減損日期"].ToString()).Month; SelStatusMM.SelectedIndex = StatusMM;
                int StatusDD = DateTime.Parse(dr["減損日期"].ToString()).Day; SelStatusDD.SelectedIndex = StatusDD;
            }
            else
            {
                SelStatusYYYY.SelectedIndex = 0; SelStatusMM.SelectedIndex = 0; SelStatusDD.SelectedIndex = 0;
            }
            if (!(dr["填表日期"] is DBNull))
            {
                int KeyYYYY = DateTime.Parse(dr["填表日期"].ToString()).Year;
                if (KeyYYYY - MinYear + 1 > SelKeyYYYY.Items.Count) SelKeyYYYY.SelectedIndex = SelKeyYYYY.Items.Count - 1;
                else SelKeyYYYY.SelectedIndex = KeyYYYY - MinYear + 1;

                int KeyMM = DateTime.Parse(dr["填表日期"].ToString()).Month; SelKeyMM.SelectedIndex = KeyMM;
                int KeyDD = DateTime.Parse(dr["填表日期"].ToString()).Day; SelKeyDD.SelectedIndex = KeyDD;
            }
            else
            {
                SelKeyYYYY.SelectedIndex = 0; SelKeyMM.SelectedIndex = 0; SelKeyDD.SelectedIndex = 0;
            }

            TextMemo.Text = dr["備註說明"].ToString();

            if (!(dr["建立日期"] is DBNull)) LblCreateDT.Text = DateTime.Parse(dr["建立日期"].ToString()).ToString("yyyy/MM/dd HH:mm");
            else LblCreateDT.Text = "";            
            LblCreater.Text = dr["建立人員"].ToString();            
            if (!(dr["修改日期"] is DBNull)) LblCreateDT.Text = DateTime.Parse(dr["修改日期"].ToString()).ToString("yyyy/MM/dd HH:mm");
            else LblCreateDT.Text = "";
            LblUpdater.Text = dr["修改人員"].ToString();

            lblChecks.Text = GetChecks("軟體主檔", TextSwNo.Text);   //取得資料查核
        }

        cmd.Cancel(); cmd.Dispose(); dr.Close(); Conn.Close(); Conn.Dispose();
    }

    protected void ReadHelp() //讀取欄位說明
    {
        SqlCommand sqlQuery = new SqlCommand();
        sqlQuery.CommandText = "select Item,Memo,Config from Config where Kind='軟體主檔' order by Mark";
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
            Msg.Text = "<script>alert('您不是軟體小組的成員，沒有新增軟體資料的權限！');</script>";
        }
        else if (TextSwName.Text == "")
        {
            Msg.Text = "<script>alert('您未填軟體名稱！');</script>";
        }
        else if (FormCheck("Ins"))
        {
            Msg.Text = "<script>alert('表單編號重複！');</script>";
        }
        else
        {
            if (TextLcsNum.Text == "") TextLcsNum.Text = "0";
            if (TextPrice.Text == "") TextPrice.Text = "0";
            if (TextTotal.Text == "") TextTotal.Text = "0";

            LblCreateDT.Text = DateTime.Now.ToString("yyyy/MM/dd HH:mm");
            LblCreater.Text = Session["UserName"].ToString();
            LblUpdateDT.Text = LblCreateDT.Text;
            LblUpdater.Text = LblCreater.Text;

            string SwNo = GetPKNo("軟體編號", "軟體主檔").ToString();
            TextSwNo.Text = SwNo; //新增完成後，要賦予新取得之軟體編號 

            string SQL = GetInsSQL(SwNo);
            //Response.Write(SQL);
            //Response.End();
            ExecDbSQL(SQL);
            InsLifeLog(SQL);

            Msg.Text = "<script>alert('新增資料[" + TextSwName.Text + "] 完成！');window.open('SwEdit.aspx?SwNo=" + TextSwNo.Text + "','_self');</script>";
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
        int PkNo = 1; if (dr.Read()) PkNo = int.Parse(dr[0].ToString()) + 1;
        cmd.Cancel(); cmd.Dispose(); dr.Close(); Conn.Close(); Conn.Dispose();

        return (PkNo);
    }

    protected void BtnDel_Click(object sender, EventArgs e)    //真刪除，請注意生命履歷要詳細記錄所有所刪資料之詳細所有欄位資料
    {
        Literal Msg = new Literal();

        if (!RightCheck())
        {
            Msg.Text = "<script>alert('您不是軟體小組的成員，沒有刪除軟體資料的權限！');</script>";
        }
        else if (AskCheck())
        {
            Msg.Text = "<script>alert('請注意需先刪除所有相依的授權申請，才能刪除本資料！');</script>";
        }
        else
        {
            string SQL = "delete from [軟體主檔] where [軟體編號]=" + TextSwNo.Text;
            ExecDbSQL(SQL);

            InsLifeLog("刪除 [" + TextSwName.Text + "] ： " + SQL + "，原本軟體主檔SQL ： " + GetInsSQL(TextSwNo.Text));

            Msg.Text = "<script>alert('完成刪除 [" + TextSwName.Text + "] 資料！');window.close();</script>";
        }

        Page.Controls.Add(Msg);
    }

    protected void BtnEdit_Click(object sender, EventArgs e) //按下修改按鈕
    {
        Literal Msg = new Literal();

        if (!RightCheck())
        {
            Msg.Text = "<script>alert('您不是軟體小組的成員，沒有異動軟體資料的權限！');</script>";
        }
        else if (FormCheck("Edit"))
        {
            Msg.Text = "<script>alert('表單編號重複！');</script>";
        }
        else if (TextSwName.Text == "")
        {
            Msg.Text = "<script>alert('您未填軟體名稱！');</script>";
        }
        else
        {
            if (TextLcsNum.Text == "") TextLcsNum.Text = "0";
            if (TextPrice.Text == "") TextPrice.Text = "0";
            if (TextTotal.Text == "") TextTotal.Text = "0";
            LblUpdateDT.Text = DateTime.Now.ToString("yyyy/MM/dd HH:mm");
            LblUpdater.Text = Session["UserName"].ToString();

            InsLifeLog("修改 [" + TextSwName.Text + "] ：： " + GetUpdate(int.Parse(TextSwNo.Text), "Life"));
            ExecDbSQL("Update [軟體主檔] set " + GetUpdate(int.Parse(TextSwNo.Text), "SQL") + " where [軟體編號]=" + TextSwNo.Text);

            Msg.Text = "<script>alert('更新資料[" + TextSwName.Text + "] 完成！');</script>";
            ReadSw();
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

    protected string GetInsSQL(string SwNo) //取得新增資料的語法
    {
        return ("insert into [軟體主檔] values("
            + SwNo + ",'"
            + SelFormYYYY.SelectedValue + "-" + TextFormXXXX.Text + "','"
            + SelUnit.SelectedValue + "','"
            + TextAsker.Text + "','"
            + TextSwName.Text + "','"
            + TextVersBuy.Text + "','"
            + TextVersCan.Text + "','"
            + SelLcsKind.SelectedValue + "',"
            + TextLcsNum.Text + ",'"
            + TextLcsMemo.Text + "','"
            + TextSource.Text + "','"
            + TextFunctions.Text + "','"
            + TextBrand.Text + "','"
            + TextStyle.Text + "','"
            + TextSnBuy.Text + "','"
            + TextSnDown.Text + "','"
            + TextLife.Text + "',"
            + GetYMD(SelBuyYYYY, SelBuyMM, SelBuyDD, "p") + ","
            + GetYMD(SelUseYYYY, SelUseMM, SelUseDD, "p") + ","
            + GetYMD(SelUpYYYY, SelUpMM, SelUpDD, "p") + ","
            + TextPrice.Text + ","
            + TextTotal.Text + ",'"
            + TextPriceMemo.Text + "','"
            + TextSupply.Text + "','"
            + TextMedia.Text + "','"
            + TextBookNo.Text + "','"
            + TextAttach.Text + "','"
            + SelStatus.SelectedValue + "','"
            + TextCause.Text + "','"
            + TextExec.Text + "',"
            + GetYMD(SelStatusYYYY, SelStatusMM, SelStatusDD, "p") + ","
            + GetYMD(SelKeyYYYY, SelKeyMM, SelKeyDD, "p") + ",'"
            + LblCreateDT.Text + "','"
            + LblCreater.Text + "','"
            + LblUpdateDT.Text + "','"
            + LblUpdater.Text + "','"
            + TextMemo.Text + "'"
            + ")");
    }

    protected void BtnPre_Click(object sender, EventArgs e) //取得預覽值
    {
        Label1.Text ="軟體編號 -- " + TextSwNo.Text + "<br />"
            + "表單編號 -- " + SelFormYYYY.SelectedValue + "-" + TextFormXXXX.Text + "<br />"
            + "軟體單位 -- " + SelUnit.SelectedValue + "<br />"
            + "軟體人員 -- " + TextAsker.Text + "<br />"
            + "軟體名稱 -- " + TextSwName.Text + "<br />"
            + "購買版本 -- " + TextVersBuy.Text + "<br />"
            + "可用版本 -- " + TextVersCan.Text + "<br />"
            + "授權方式 -- " + SelLcsKind.SelectedValue + "<br />"
            + "授權數量 -- " + TextLcsNum.Text + "<br />"
            + "授權說明 -- " + TextLcsMemo.Text + "<br />"
            + "軟體來源 -- " + TextSource.Text + "<br />"
            + "軟體功能 -- " + TextFunctions.Text + "<br />"
            + "適用廠牌 -- " + TextBrand.Text + "<br />"
            + "適用型式 -- " + TextStyle.Text + "<br />"
            + "購買序號 -- " + TextSnBuy.Text + "<br />"
            + "降級序號 -- " + TextSnDown.Text + "<br />"
            + "期限說明 -- " + TextLife.Text + "<br />"
            + "購買日期 -- " + GetYMD(SelBuyYYYY, SelBuyMM, SelBuyDD, "m") + "<br />"
            + "使用期限 -- " + GetYMD(SelUseYYYY, SelUseMM, SelUseDD, "m") + "<br />"
            + "更新期限 -- " + GetYMD(SelUpYYYY, SelUpMM, SelUpDD, "m") + "<br />"
            + "軟體單價 -- " + TextPrice.Text + "<br />"
            + "軟體總價 -- " + TextTotal.Text + "<br />"
            + "價格說明 -- " + TextPriceMemo.Text + "<br />"
            + "提供單位 -- " + TextSupply.Text + "<br />"
            + "存放媒體 -- " + TextMedia.Text + "<br />"
            + "圖書編號 -- " + TextBookNo.Text + "<br />"
            + "軟體附件 -- " + TextAttach.Text + "<br />"
            + "軟體狀態 -- " + SelStatus.SelectedValue + "<br />"
            + "減損原因 -- " + TextCause.Text + "<br />"
            + "減損方式 -- " + TextExec.Text + "<br />"
            + "減損日期 -- " + GetYMD(SelStatusYYYY, SelStatusMM, SelStatusDD, "m") + "<br />"
            + "填表日期 -- " + GetYMD(SelKeyYYYY, SelKeyMM, SelKeyDD, "m") + "<br />"
            + "建立日期 -- " + LblCreateDT.Text + "<br />"
            + "建立人員 -- " + LblCreater.Text + "<br />"
            + "修改日期 -- " + LblUpdateDT.Text + "<br />"
            + "修改人員 -- " + LblUpdater.Text + "<br />"
            + "備註說明 -- " + TextMemo.Text + "<br />";
    }

    protected string GetUpdate(int SwNo, string SQLorLife) //取得修改資料的語法
    {
        string SQL = "";
        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
        Conn.Open();
        SqlCommand cmd = new SqlCommand("select * from [軟體主檔] where [軟體編號]=" + SwNo, Conn);
        SqlDataReader dr = null;
        string SaveDT = "", BuyDay = "", UseDay = "", UpDay = "", StatusDay = "", KeyDay = "";
        dr = cmd.ExecuteReader();
        if (dr.Read())
        {
            if (!(dr["修改日期"] is DBNull)) SaveDT = DateTime.Parse(dr["修改日期"].ToString()).ToString("yyyy/MM/dd HH:mm");
            if (!(dr["購買日期"] is DBNull)) BuyDay = DateTime.Parse(dr["購買日期"].ToString()).ToString("yyyy/MM/dd");
            if (!(dr["使用期限"] is DBNull)) UseDay = DateTime.Parse(dr["使用期限"].ToString()).ToString("yyyy/MM/dd");
            if (!(dr["更新期限"] is DBNull)) UpDay = DateTime.Parse(dr["更新期限"].ToString()).ToString("yyyy/MM/dd");
            if (!(dr["減損日期"] is DBNull)) StatusDay = DateTime.Parse(dr["減損日期"].ToString()).ToString("yyyy/MM/dd");
            if (!(dr["填表日期"] is DBNull)) KeyDay = DateTime.Parse(dr["填表日期"].ToString()).ToString("yyyy/MM/dd");
            
            SQL = GetUpdateCol("表單編號", dr["表單編號"].ToString(), SelFormYYYY.SelectedValue + "-" + TextFormXXXX.Text, "string", SQLorLife)
                + GetUpdateCol("軟體單位", dr["軟體單位"].ToString(), SelUnit.SelectedValue, "string", SQLorLife)
                + GetUpdateCol("軟體人員", dr["軟體人員"].ToString(), TextAsker.Text, "string", SQLorLife)
                + GetUpdateCol("軟體名稱", dr["軟體名稱"].ToString(), TextSwName.Text, "string", SQLorLife)
                + GetUpdateCol("購買版本", dr["購買版本"].ToString(), TextVersBuy.Text, "string", SQLorLife)
                + GetUpdateCol("可用版本", dr["可用版本"].ToString(), TextVersCan.Text, "string", SQLorLife)
                + GetUpdateCol("授權方式", dr["授權方式"].ToString(), SelLcsKind.SelectedValue, "string", SQLorLife)
                + GetUpdateCol("授權數量", dr["授權數量"].ToString(), TextLcsNum.Text, "integer", SQLorLife)
                + GetUpdateCol("授權說明", dr["授權說明"].ToString(), TextLcsMemo.Text, "string", SQLorLife)
                + GetUpdateCol("軟體來源", dr["軟體來源"].ToString(), TextSource.Text, "string", SQLorLife)
                + GetUpdateCol("軟體功能", dr["軟體功能"].ToString(), TextFunctions.Text, "string", SQLorLife)
                + GetUpdateCol("適用廠牌", dr["適用廠牌"].ToString(), TextBrand.Text, "string", SQLorLife)
                + GetUpdateCol("適用型式", dr["適用型式"].ToString(), TextStyle.Text, "string", SQLorLife)
                + GetUpdateCol("購買序號", dr["購買序號"].ToString(), TextSnBuy.Text, "string", SQLorLife)
                + GetUpdateCol("降級序號", dr["降級序號"].ToString(), TextSnDown.Text, "string", SQLorLife)
                + GetUpdateCol("期限說明", dr["期限說明"].ToString(), TextLife.Text, "string", SQLorLife)
                + GetUpdateCol("購買日期", BuyDay, GetYMD(SelBuyYYYY, SelBuyMM, SelBuyDD, "m"), "date", SQLorLife)
                + GetUpdateCol("使用期限", UseDay, GetYMD(SelUseYYYY, SelUseMM, SelUseDD, "m"), "date", SQLorLife)
                + GetUpdateCol("更新期限", UpDay, GetYMD(SelUpYYYY, SelUpMM, SelUpDD, "m"), "date", SQLorLife)
                + GetUpdateCol("軟體單價", Int64.Parse(dr["軟體單價"].ToString(),System.Globalization.NumberStyles.AllowThousands | System.Globalization.NumberStyles.AllowDecimalPoint).ToString(), TextPrice.Text, "money", SQLorLife)
                + GetUpdateCol("軟體總價", Int64.Parse(dr["軟體總價"].ToString(),System.Globalization.NumberStyles.AllowThousands | System.Globalization.NumberStyles.AllowDecimalPoint).ToString(), TextTotal.Text, "money", SQLorLife)
                + GetUpdateCol("價格說明", dr["價格說明"].ToString(), TextPriceMemo.Text, "string", SQLorLife)
                + GetUpdateCol("提供單位", dr["提供單位"].ToString(), TextSupply.Text, "string", SQLorLife)
                + GetUpdateCol("存放媒體", dr["存放媒體"].ToString(), TextMedia.Text, "string", SQLorLife)
                + GetUpdateCol("圖書編號", dr["圖書編號"].ToString(), TextBookNo.Text, "string", SQLorLife)
                + GetUpdateCol("軟體附件", dr["軟體附件"].ToString(), TextAttach.Text, "string", SQLorLife)
                + GetUpdateCol("軟體狀態", dr["軟體狀態"].ToString(), SelStatus.SelectedValue, "string", SQLorLife)
                + GetUpdateCol("減損原因", dr["減損原因"].ToString(), TextCause.Text, "string", SQLorLife)
                + GetUpdateCol("減損方式", dr["減損方式"].ToString(), TextExec.Text, "string", SQLorLife)
                + GetUpdateCol("減損日期", StatusDay, GetYMD(SelStatusYYYY, SelStatusMM, SelStatusDD, "m"), "date", SQLorLife)
                + GetUpdateCol("填表日期", KeyDay, GetYMD(SelKeyYYYY, SelKeyMM, SelKeyDD, "m"), "date", SQLorLife)
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
        string TblName = "軟體主檔";    //表格名稱
        string PKno = TextSwNo.Text;   //主鍵編號
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

    protected Boolean AskCheck() //請注意需先刪除所有相依的授權申請，才能刪除本資料！
    {
        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
        Conn.Open();
        SqlCommand cmd = new SqlCommand("select * from [軟體授權] where [軟體編號]=" + TextSwNo.Text, Conn);
        SqlDataReader dr = null;
        dr = cmd.ExecuteReader();
        Boolean TF = false; if (dr.Read()) TF = true;
        cmd.Cancel(); cmd.Dispose(); dr.Close(); Conn.Close(); Conn.Dispose();

        return (TF);
    }

    protected Boolean SwCheck() //是否有軟體編號記錄！
    {
        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
        Conn.Open();
        SqlCommand cmd = new SqlCommand("select * from [軟體主檔] where [軟體編號]=" + TextNewSwNo.Text, Conn);
        SqlDataReader dr = null;
        dr = cmd.ExecuteReader();
        Boolean TF = false; if (dr.Read()) TF = true;
        cmd.Cancel(); cmd.Dispose(); dr.Close(); Conn.Close(); Conn.Dispose();

        return (TF);
    }

    protected Boolean FormCheck(string DBaction) //檢查表單編號是否重複！
    {
        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
        Conn.Open();
        SqlCommand cmd;
        if (DBaction == "Ins") cmd = new SqlCommand("select [表單編號] from [軟體主檔] where [表單編號]='" + SelFormYYYY.SelectedValue + "-" + TextFormXXXX.Text + "'", Conn);
        else cmd = new SqlCommand("select [表單編號] from [軟體主檔] where [軟體編號]<>" + TextSwNo.Text + " and [表單編號]='" + SelFormYYYY.SelectedValue + "-" + TextFormXXXX.Text + "'", Conn);
        SqlDataReader dr = null;
        dr = cmd.ExecuteReader();
        Boolean TF = false; if (dr.Read()) TF = true;
        cmd.Cancel(); cmd.Dispose(); dr.Close(); Conn.Close(); Conn.Dispose();

        return (TF);
    }

    protected Boolean RightCheck() //是否有權修改資料
    {
        if (InGroup(Session["UserName"].ToString(), "軟體小組") | InGroup(Session["UserName"].ToString(), "SSM小組")) return (true);
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

    protected void LinkLcs_Click(object sender, EventArgs e)  //授權查詢
    {
        Session["AskSQL"] = "SELECT * FROM [View_軟體管理] WHERE [軟體編號]=" + TextSwNo.Text + " and [授權編號] is not null";
        OpenExecWindow("window.open('Ask.aspx?Search=SwEdit','_self');");
    }

    protected void LinkNewSwNo_Click(object sender, EventArgs e)  //批次修改
    {
        Literal Msg = new Literal();

        if (!RightCheck())
        {
            Msg.Text = "<script>alert('您不是軟體小組的成員，沒有異動軟體資料的權限！');</script>";
        }
        else if (TextSwNo.Text == "" | TextSwNo.Text == "0" | TextNewSwNo.Text == "")
        {
            Msg.Text = "<script>alert('新或舊軟體編號空白，無法執行！');</script>";
        }
        else if (!SwCheck())
        {
            Msg.Text = "<script>alert('軟體編號為 " + TextNewSwNo.Text + " 的軟體保管單尚未建檔，請先建立再執行此功能！');</script>";
        }
        else
        {
            string SQL = "Update [軟體授權] set [軟體編號]=" + TextNewSwNo.Text + " where [軟體編號]=" + TextSwNo.Text;

            ExecDbSQL(SQL);
            InsLifeLog(SelFormYYYY.SelectedValue + "-" + TextFormXXXX.Text + "軟體編號批次修改：" + SQL);

            Msg.Text = "<script>alert('已將[軟體授權]主檔中[軟體編號]為" + TextSwNo.Text + "的所有資料均改為" + TextNewSwNo.Text + "！');</script>";
        }

        Page.Controls.Add(Msg);
    }

    protected void SelUnit_SelectedIndexChanged(object sender, EventArgs e)  //選擇軟體單位
    {
        SelAsker.DataBind();
        TextAsker.Text = SelAsker.Items[0].Value;
    }
    protected void SelAsker_SelectedIndexChanged(object sender, EventArgs e)   //選擇軟體人員
    {
        DropDownList Sel = (DropDownList)form1.FindControl("SelAsker");    //放入Table物件後，必須用FindControl才能取值
        TextAsker.Text = Sel.SelectedValue;
    }

    protected void SelSwName_SelectedIndexChanged(object sender, EventArgs e)  //選擇軟體來源
    {
        DropDownList Sel = (DropDownList)form1.FindControl("SelSwName");    //放入Table物件後，必須用FindControl才能取值
        TextSwName.Text = Sel.SelectedValue;

        SelVersBuy.DataBind();
        TextVersBuy.Text = SelVersBuy.Items[0].Value;
        SelVersCan.DataBind();
        TextVersCan.Text = SelVersCan.Items[0].Value;
    }

    protected void SelVersBuy_SelectedIndexChanged(object sender, EventArgs e)   //選擇軟體版本
    {
        DropDownList Sel = (DropDownList)form1.FindControl("SelVersBuy");    //放入Table物件後，必須用FindControl才能取值
        TextVersBuy.Text = Sel.SelectedValue;
    }

    protected void SelVersCan_SelectedIndexChanged(object sender, EventArgs e)    //選擇可用版本
    {
        DropDownList Sel = (DropDownList)form1.FindControl("SelVersCan");    //放入Table物件後，必須用FindControl才能取值
        TextVersCan.Text = Sel.SelectedValue;
    }

    protected void SelSource_SelectedIndexChanged(object sender, EventArgs e)  //選擇軟體來源
    {
        DropDownList Sel = (DropDownList)form1.FindControl("SelSource");    //放入Table物件後，必須用FindControl才能取值
        TextSource.Text = Sel.SelectedValue;
    }

    protected void SelFunctions_SelectedIndexChanged(object sender, EventArgs e)  //選擇軟體功能
    {
        DropDownList Sel = (DropDownList)form1.FindControl("SelFunctions");    //放入Table物件後，必須用FindControl才能取值
        TextFunctions.Text = Sel.SelectedValue;
    }

    protected void SelBrand_SelectedIndexChanged(object sender, EventArgs e)  //選擇適用廠牌
    {
        DropDownList Sel = (DropDownList)form1.FindControl("SelBrand");    //放入Table物件後，必須用FindControl才能取值
        TextBrand.Text = Sel.SelectedValue;
        SelStyle.DataBind();
        TextStyle.Text = SelStyle.Items[0].Value;
    }

    protected void SelStyle_SelectedIndexChanged(object sender, EventArgs e)  //選擇適用型式
    {
        DropDownList Sel = (DropDownList)form1.FindControl("SelStyle");    //放入Table物件後，必須用FindControl才能取值
        TextStyle.Text = Sel.SelectedValue;
    }

    protected void SelLife_SelectedIndexChanged(object sender, EventArgs e)  //選擇期限說明
    {
        DropDownList Sel = (DropDownList)form1.FindControl("SelLife");    //放入Table物件後，必須用FindControl才能取值
        TextLife.Text = Sel.SelectedValue;
    }

    protected void SelSupply_SelectedIndexChanged(object sender, EventArgs e)  //選擇提供單位
    {
        DropDownList Sel = (DropDownList)form1.FindControl("SelSupply");    //放入Table物件後，必須用FindControl才能取值
        TextSupply.Text = Sel.SelectedValue;
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

    protected void BtnLife_Click(object sender, EventArgs e) //查詢生命履歷
    {
        Session["LifeSQL"] = "select * from [生命履歷] where [表格名稱]='軟體主檔' and [主鍵編號]=" + TextSwNo.Text + " order by [履歷編號] desc";
        OpenExecWindow("window.open('../Life/LifeLog.aspx?Search=SwEdit&Tbl=軟體主檔&PK=" + TextSwNo.Text + "','_self');");
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