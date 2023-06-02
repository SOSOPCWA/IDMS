using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.Configuration;
using System.Data;
using System.Data.SqlClient;

public partial class SoftWare_AskForm : System.Web.UI.Page
{
    string AskNo = "";

    protected void Page_Load(object sender, EventArgs e)
    {
        if (Request["AskNo"] != null & Request["AskNo"] != "")
        {
            AskNo = Request["AskNo"].ToString();
            ReadForm(AskNo);
        }
    }

    protected void Page_PreRenderComplete(object sender, EventArgs e)   //放Page_Load無法顯示
    {
        Literal Msg = new Literal();
        Msg.Text = "<script>alert('[建議]：\\n"
                                + "1. 請先預覽列印調整設定，使能用一張紙就好\\n"
                                + "2. 請取消所有頁首頁尾設定\\n"
                                + "3. 預設為安裝，除非最新申請單[申請事項]有`移除`或`異動`字眼\\n"
                                + "4. 可攜式電腦使用微軟軟體：單機版(NB)\\n"
                                + "   a.[設備種類]非筆記型電腦，則匯入主使用\\n"
                                + "   b.[設備種類]為筆記型電腦，則匯入筆電\\n"                                
                                + "5. 微軟授權軟體使用評估申請單：試用版\\n"                                
                                + "');</script>";
        Page.Controls.Add(Msg);
    }

    protected void ReadForm(string AskNo)    //讀取授權資料
    {
        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
        Conn.Open();
        SqlCommand cmd = new SqlCommand("SELECT * FROM [View_軟體管理] WHERE [授權編號]=" + AskNo, Conn);
        SqlDataReader dr = null;
        dr = cmd.ExecuteReader();

        if (dr.Read())
        {
            //--------------------------保管單資訊--------------------------------------------------------------------------------------------------------------
            //lblAskNo.Text = AskNo; 
            if (dr["授權單位"].ToString() != "") lblAskUnit.Text = dr["授權單位"].ToString();
            if (!(dr["授權填表日期"] is DBNull)) lblKeyDay.Text = DateTime.Parse(dr["授權填表日期"].ToString()).ToString("yyyy 年 MM 月 dd 日");
            lblSwName.Text = "YYYY-NNNN"; ; if (dr["軟體名稱"].ToString() != "") lblSwName.Text = dr["軟體名稱"].ToString();
            if (dr["購買版本"].ToString() != "") lblVer.Text = dr["購買版本"].ToString();
            if (dr["授權主機"].ToString() != "") lblHost.Text = dr["授權主機"].ToString();
            if (dr["授權IP"].ToString() != "") lblIP.Text = dr["授權IP"].ToString();

            string Ask = "安裝";
            ChkIns.Checked = true;
            //--------------------------申請單資訊--------------------------------------------------------------------------------------------------------------
            SqlConnection Conn1 = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
            Conn1.Open();
            SqlCommand cmd1 = new SqlCommand("select * from [申請表單] where [表單種類]='軟體申請' and [主鍵編號]=" + AskNo + " order by [填表日期] desc", Conn1);
            SqlDataReader dr1 = null;
            dr1 = cmd1.ExecuteReader();
            if (dr1.Read())
            {
                lblFormNo.Text = "YYYY-NNNN"; lblFormNo.ForeColor = System.Drawing.Color.LightGray; lblFormNo.Font.Bold = false;
                if (dr1["表單編號"].ToString() != "")
                {
                    lblFormNo.Text = dr1["表單編號"].ToString();
                    if (lblFormNo.Text.Length > 9) lblFormNo.Text = lblFormNo.Text.Substring(0, 9);
                    lblFormNo.ForeColor = System.Drawing.Color.Black; lblFormNo.Font.Bold = true;
                }

                lblKeyDay.Text = "　年　月　日";
                if (!(dr1["填表日期"] is DBNull)) lblKeyDay.Text = DateTime.Parse(dr1["填表日期"].ToString()).ToString("yyyy 年 MM 月 dd 日");

                Ask = dr1["申請事項"].ToString();
            }
            cmd1.Cancel(); cmd1.Dispose(); dr1.Close(); Conn1.Close(); Conn1.Dispose();

            //授權方式：單機版、單機版(NB)、開發粄、隨機版、網路版、Server版、Server版(V1)、Server版(V4)、Server版(V∞)、Client版、Server & Client版、升級版、其它，欄位取自[軟體主檔]，不用輸入
            switch (dr["申請授權"].ToString())
            {
                case "單機版(NB)": //一單機一NB(Select 大量授權)
                    {
                        ChkNB.Checked = true;
                        if (!(dr["設備種類"] is DBNull))
                        {
                            if (dr["設備種類"] == "筆記型電腦")
                            {
                                if (dr["授權主機"].ToString() != "") lblNBHost.Text = dr["授權主機"].ToString() + "　";
                                if (dr["授權IP"].ToString() != "") lblNBIP.Text = dr["授權IP"].ToString() + "　";
                            }
                            else
                            {
                                if (dr["授權主機"].ToString() != "") lblSelHost.Text = dr["授權主機"].ToString() + "　";
                                if (dr["授權IP"].ToString() != "") lblSelIP.Text = dr["授權IP"].ToString() + "　";
                            }
                        }
                        break;
                    }
                case "試用版": //60日軟體使用評估(Select)
                    {
                        ChkTest.Checked = true;
                        if (dr["授權主機"].ToString() != "") lblTestHost.Text = dr["授權主機"].ToString() + "　";
                        if (dr["授權IP"].ToString() != "") lblTestIP.Text = dr["授權IP"].ToString() + "　";
                        break;
                    }
            }

            if (Ask.IndexOf("移除") >= 0)
            {
                ChkIns.Checked = false;
                ChkDel.Checked = true;
            }

            if (Ask.IndexOf("異動") >= 0)
            {
                ChkIns.Checked = false;
                ChkUpd.Checked = true;
                if (dr["使用主機"].ToString() != "") lblOldHost.Text = dr["使用主機"].ToString() + "　";
                if (dr["使用IP"].ToString() != "") lblOldIP.Text = dr["使用IP"].ToString() + "　";
                if (dr["授權主機"].ToString() != "") lblNewHost.Text = dr["授權主機"].ToString() + "　";
                if (dr["授權IP"].ToString() != "") lblNewIP.Text = dr["授權IP"].ToString() + "　";
            }
        }

        cmd.Cancel(); cmd.Dispose(); dr.Close(); Conn.Close(); Conn.Dispose();
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
}