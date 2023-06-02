using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.Configuration;
using System.Data;
using System.Data.SqlClient;

public partial class Life_LifeLog : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        Session.Timeout = 720;
        if (!IsPostBack)
        {
            string x = "";

            ListItem ItemY = new ListItem(); ItemY.Text = ""; ItemY.Value = ""; SelYYYY.Items.Add(ItemY);
            int YYYY = int.Parse((DateTime.Now).ToString("yyyy"));
            for (int i = YYYY; i >= YYYY - 5; i--)   //產生異動日期年之選項
            {
                x = i.ToString(); if (i < 10) x = "0" + x;
                ListItem ItemX = new ListItem(); ItemX.Text = x; ItemX.Value = x; if (i == YYYY) ItemX.Selected = true;
                SelYYYY.Items.Add(ItemX);
            }

            ListItem ItemM = new ListItem(); ItemM.Text = ""; ItemM.Value = ""; SelMM.Items.Add(ItemM);
            int MM = int.Parse((DateTime.Now).ToString("MM"));
            for (int i = 1; i <= 12; i++)   //產生異動日期月之選項
            {
                x = i.ToString(); if (i < 10) x = "0" + x;
                ListItem ItemX = new ListItem(); ItemX.Text = x; ItemX.Value = x; if (i == MM) ItemX.Selected = true;
                SelMM.Items.Add(ItemX);
            }

            ListItem ItemD = new ListItem(); ItemD.Text = ""; ItemD.Value = ""; SelDD.Items.Add(ItemD);
            int DD = int.Parse(DateTime.Now.ToString("dd"));
            for (int i = 1; i <= 31; i++)   //產生異動日期月之選項
            {
                x = i.ToString(); if (i < 10) x = "0" + x;
                ListItem ItemX = new ListItem(); ItemX.Text = x; ItemX.Value = x;
                SelDD.Items.Add(ItemX);
            }

            if (RightCheck())   //有無權利看到軟體異動記錄
            {
                ListItem ItemX = new ListItem(); ItemX.Text = "軟體主檔"; ItemX.Value = "軟體主檔"; SelTbl.Items.Add(ItemX);
                ListItem ItemZ = new ListItem(); ItemZ.Text = "軟體授權"; ItemZ.Value = "軟體授權"; SelTbl.Items.Add(ItemZ);
            }

            if (Request["Search"] != null)  //外部查詢
            {
                if (Session["LifeSQL"] != null) 
                {
                    ViewState["SQL"] = Session["LifeSQL"];
                    SelYYYY.SelectedIndex = 0; SelMM.SelectedIndex = 0;
                    for (int i = 0; i < SelTbl.Items.Count; i++) if (SelTbl.Items[i].Value == Request["Tbl"].ToString()) SelTbl.SelectedIndex = i;
                    TextPK.Text = Request["PK"].ToString();
                }
            }
            else
            {
                ViewState["SQL"] = "select * from [生命履歷] where [異動日期]>='" + DateTime.Now.ToString("yyyy/MM/01 00:00:00") + "' order by [履歷編號] desc";
            }

            SqlDataSource3.SelectCommand = ViewState["SQL"].ToString();
            GridView1.DataBind();
        }
    }

    protected void Page_PreRenderComplete(object sender, EventArgs e)   //放Page_Load無法顯示
    {
        SqlDataSource3.SelectCommand = ViewState["SQL"].ToString(); //查詢後，SQL要保持住，否則會用預設值
    }

    protected void SelUnit_SelectedIndexChanged(object sender, EventArgs e)
    {
        SelMt.DataBind();
    }

    protected void Button1_Click(object sender, EventArgs e)
    {
        if (ChkPage.Checked) GridView1.AllowPaging = true;
        else GridView1.AllowPaging = false;
        
        string SQL = "select * from [生命履歷] where 1=1";

        string UnitSQL = "select [成員] from [View_組織架構] where [課別]='" + SelUnit.SelectedValue + "' or [成員] in (select [Item] from [Config] where [Kind]='" + SelUnit.SelectedValue + "')";
        if (ChkUnit.Checked) SQL = SQL + " and ([維護人員] in (" + UnitSQL + ") or [原負責人] in (" + UnitSQL + ") or [異動人員] in (" + UnitSQL + ") or [原保管人] in (" + UnitSQL + ")"
            + " or [表格名稱]='實體設備' and [主鍵編號] in (select [設備編號] from [View_設備管理] where [保管人員] in (" + UnitSQL + ")))";
        if (ChkMt.Checked) SQL = SQL + " and ([維護人員]='" + SelMt.SelectedValue + "' or [維護人員] in (select Kind from Config where Item='" + SelMt.SelectedValue + "')"
            + " or [原負責人]='" + SelMt.SelectedValue + "' or [原負責人] in (select Kind from Config where Item='" + SelMt.SelectedValue + "')"
            + " or [異動人員]='" + SelMt.SelectedValue + "' or [原保管人]='" + SelMt.SelectedValue + "'"
            + " or [表格名稱]='實體設備' and [主鍵編號] in (select [設備編號] from [View_設備管理] where [保管人員]='" + SelMt.SelectedValue + "'))";

        string YYYY1, YYYY2, MM1, MM2, DD1, DD2;
        if (SelYYYY.SelectedValue != "") { YYYY1 = SelYYYY.SelectedValue; YYYY2 = YYYY1; }
        else { YYYY1 = "1900"; YYYY2 = "9999"; }
        if (SelMM.SelectedValue != "") { MM1 = SelMM.SelectedValue; MM2 = MM1; }
        else { MM1 = "01"; MM2 = "12"; }
        if (SelDD.SelectedValue != "") { DD1 = SelDD.SelectedValue; DD2 = DD1; }
        else { DD1 = "01"; DD2 = DateTime.DaysInMonth(int.Parse(YYYY2), int.Parse(MM2)).ToString(); }

        SQL = SQL + " and [異動日期] between '" + YYYY1 + "/" + MM1 + "/" + DD1 + " 00:00:00' and '" + YYYY2 + "/" + MM2 + "/" + DD2 + " 23:59:59'";

        if (SelTbl.SelectedValue != "") SQL = SQL + " and [表格名稱]='" + SelTbl.SelectedValue + "'";

        
        if (TextPK.Text != "") SQL = SQL + " and [主鍵編號]=" + TextPK.Text;

        if (TextLife.Text != "")
        {
            string[] KeyA = TextLife.Text.Trim().Split(',');
            for (int i = 0; i < KeyA.GetLength(0); i++) SQL = SQL + " and [生命履歷] like '%" + KeyA[i] + "%'";
        }

        SQL = SQL + " order by [履歷編號] desc";

        int n; Literal Msg = new Literal();
        if (TextPK.Text == "" | int.TryParse(TextPK.Text.ToString(), out n))
        {
            ViewState["SQL"] = SQL;
            //Response.Write(SQL);
            //Response.End();
            SqlDataSource3.SelectCommand = SQL;
            GridView1.DataBind();
        }
        else
        {
            Msg.Text = "<script>alert('請檢查輸入的主鍵編號是否為數字！');</script>";
            Page.Controls.Add(Msg);
        }
    }

    protected void GridView1_SelectedIndexChanging(object sender, GridViewSelectEventArgs e)    //選取後就另開編輯視窗或帶入設備編號
    {
        Literal Msg = new Literal();
        string LifeNo = GridView1.Rows[e.NewSelectedIndex].Cells[1].Text;

        if (TextLife.Text != "#delete#")
        {
            string tblName = GridView1.Rows[e.NewSelectedIndex].Cells[2].Text; if (tblName == "秘總財產" | tblName == "秘總財產") tblName = "實體設備";
            string PKno = GridView1.Rows[e.NewSelectedIndex].Cells[3].Text;
            string PKstr = "[設備編號]=" + PKno;

            if (tblName == "Config" | tblName == "定位設定" | tblName == "資產清冊") Msg.Text = "<script>alert('系統設定未設計直接帶出資料機制，請自行至該介面查詢!');</script>";
            else if (tblName == "設備迴路") OpenExecWindow("window.open('../Device/TreeEdit.aspx?DevNo=" + PKno + "','_blank');");
            else if (tblName == "系統迴路") OpenExecWindow("window.open('../AP/SysTree.aspx?SysNo=" + PKno + "','_self');");
            else 
            {
                if (tblName == "作業主機") PKstr = "[作業編號]=" + PKno;
                if (tblName == "系統資源") PKstr = "[資源編號]=" + PKno;
                if (tblName == "軟體主檔") PKstr = "[軟體編號]=" + PKno;
                if (tblName == "軟體授權") PKstr = "[授權編號]=" + PKno;

                SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
                Conn.Open();
                SqlCommand cmd = new SqlCommand("select * from [" + tblName + "] where " + PKstr, Conn);
                SqlDataReader dr = null;
                dr = cmd.ExecuteReader();
                if (dr.Read())
                {
                    string url;

                    switch (tblName)
                    {
                        case "實體設備": url = "../Device/DevEdit.aspx?DevNo=" + PKno; break;
                        case "作業主機": url = "../AP/ApEdit.aspx?ApNo=" + PKno; break;
                        case "系統資源": url = "../AP/SysEdit.aspx?SysNo=" + PKno; break;
                        case "軟體主檔": url = "../SoftWare/SwEdit.aspx?SwNo=" + PKno; break;
                        case "軟體授權": url = "../SoftWare/AskEdit.aspx?AskNo=" + PKno; break;
                        default: url = ""; break;
                    }

                    OpenExecWindow("window.open('" + url + "','_blank');");
                }
                else
                {

                    Msg.Text = "<script>alert('該筆資料已經刪除，無法顯示!');</script>";

                }
                cmd.Cancel(); cmd.Dispose(); dr.Close(); Conn.Close(); Conn.Dispose();
            }
        }
        else
        {
            SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
            Conn.Open();
            SqlCommand cmd = new SqlCommand("delete from [生命履歷] where [履歷編號]=" + LifeNo, Conn);
            cmd.ExecuteNonQuery();
            cmd.Cancel(); cmd.Dispose(); Conn.Close(); Conn.Dispose();
        }

        if(Msg.Text != "") Page.Controls.Add(Msg);
    }

    protected void OpenExecWindow(string strJavascript)    //選取後就另開視窗
    {
        if (ScriptManager.GetCurrent(Page) == null)
        {
            Page.ClientScript.RegisterStartupScript(Page.GetType(), "LifeLog", strJavascript, true);
        }
        else
        {
            ScriptManager.RegisterStartupScript(Page, Page.GetType(), "LifeLog", strJavascript, true);
        }
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
}