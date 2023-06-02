using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.Configuration;
using System.Data.SqlClient;

public partial class Config_Config : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        Session.Timeout = 720;
        if (!Page.IsPostBack)
        {
            ListBox1.DataSourceID = "SqlDataSource2";
        }
    }

    protected void Page_PreRenderComplete(object sender, EventArgs e)   //放Page_Load無法顯示
    {
        lblKind.Text = GetValue("IDMS", "select [Memo] from [Config] where [Kind]='" + Request.QueryString["kind"] + "' and Item='" + SelKind.SelectedValue + "'").Replace("\r\n", "<br />");
    }

    protected void SelKind_SelectedIndexChanged(object sender, EventArgs e)
    {
        ListBox1.DataSourceID = "SqlDataSource2";
        ClearData();
        lblKind.Text = GetValue("IDMS", "select [Memo] from [Config] where [Kind]='" + Request.QueryString["kind"] + "' and Item='" + SelKind.SelectedValue + "'");
    }

    protected void ListBox1_SelectedIndexChanged(object sender, EventArgs e)
    {
        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
        Conn.Open();
        SqlCommand cmd = new SqlCommand("select * from Config where Kind='" + SelKind.SelectedValue + "' and Item='" + ListBox1.SelectedValue + "'", Conn);
        SqlDataReader dr = null;
        dr = cmd.ExecuteReader();
        if (dr.Read())
        {
            TextItem.Text = dr["Item"].ToString();
            TextConfig.Text = dr["Config"].ToString();
            TextMark.Text = dr["Mark"].ToString();
            TextMemo.Text = dr["Memo"].ToString();
        }
        cmd.Cancel(); cmd.Dispose(); dr.Close(); Conn.Close(); Conn.Dispose();
    }

    protected Boolean HasData(string Kind, string Item)
    {
        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
        Conn.Open();
        SqlCommand cmd = new SqlCommand("select * from Config where Kind='" + Kind + "' and Item='" + Item + "'", Conn);
        SqlDataReader dr = null;
        dr = cmd.ExecuteReader();
        Boolean TF = false; if (dr.Read()) TF=true;
        cmd.Cancel(); cmd.Dispose(); dr.Close(); Conn.Close(); Conn.Dispose();

        return (TF);
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

    protected void InsLifeLog(string SQL,string oldMt) //寫入生命履歷
    {
        string LifeNo = GetPKNo("履歷編號", "生命履歷").ToString(); //履歷編號
        string TblName = "Config";    //表格名稱
        string PKno = "0";   //主鍵編號
        string Mt = "SSM小組";    //維護人員
        string UN = Session["UserName"].ToString();   //登入的UserName：異動人員
        string LiftDT = DateTime.Now.ToString("yyyy/MM/dd HH:mm");  //異動日期
        string LifeIP = Request.ServerVariables["REMOTE_ADDR"].ToString();

        oldMt = Mt; //異動課別或群組成員，亦通知該課或群組之功能暫不啟用

        ExecDbSQL("Insert into [生命履歷] values(" + LifeNo + ",'" + TblName + "'," + PKno + ",'" + SQL.Replace("'", "''") + "','" + Mt + "','" + oldMt + "','" + Mt + "','" + UN + "','" + LiftDT + "','" + LifeIP + "')");
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

    protected void BtnAdd_Click(object sender, EventArgs e)
    {
        Literal Msg = new Literal();

        if (!RightCheck()) Msg.Text = "<script>alert('您沒有權限異動資料，若有需求，請洽機房！');</script>";
        else
        {
            if (HasData(SelKind.SelectedValue, TextItem.Text))
            {
                Msg.Text = "<script>alert('該筆資料已存在，無法再新增！');</script>";
            }
            else
            {
                string SQL = "insert into Config values('" + SelKind.SelectedValue + "','" + TextItem.Text + "','" + TextConfig.Text + "','" + TextMark.Text + "','" + TextMemo.Text.Replace("'", "''") + "')";
                string oldMt = ""; if (GetValue("IDMS", "select * from [Config] where [Item]='" + SelKind.SelectedValue + "' and [Kind] in (select [Item] from [Config] where [Kind] in ('組織架構','維護群組'))") != "") oldMt = SelKind.SelectedValue;

                SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
                Conn.Open();
                SqlCommand cmd = new SqlCommand(SQL, Conn);
                cmd.ExecuteNonQuery();
                cmd.Cancel(); cmd.Dispose(); Conn.Close(); Conn.Dispose();

                ListBox1.ClearSelection();
                ListBox1.Items.Add(TextItem.Text);
                ListBox1.Items[ListBox1.Items.Count - 1].Selected = true;               

                Msg.Text = "<script>alert('資料已新增！');</script>";

                InsLifeLog(SQL,oldMt);
            }
        }
        Page.Controls.Add(Msg);
    }

    protected void BtnEdit_Click(object sender, EventArgs e)
    {
        Literal Msg = new Literal();

        if (!RightCheck()) Msg.Text = "<script>alert('您沒有權限異動資料，若有需求，請洽機房！');</script>";
        else
        {
            if (ListBox1.SelectedValue != "")
            {                
                string UpdateSQL = GetUpdate("SQL");
                if (UpdateSQL != "")
                {
                    string oldMt = ""; if (GetValue("IDMS", "select * from [Config] where [Item]='" + SelKind.SelectedValue + "' and [Kind] in (select [Item] from [Config] where [Kind] in ('組織架構','維護群組'))") != "") oldMt = SelKind.SelectedValue;
                    InsLifeLog("修改 [" + SelKind.SelectedValue + "." + ListBox1.SelectedValue + "]  ：： " + GetUpdate("Life"),oldMt);
                    ExecDbSQL("Update [Config] set " + UpdateSQL + " where Kind='" + SelKind.SelectedValue + "' and Item='" + ListBox1.SelectedValue + "'");
                    Msg.Text = "<script>alert('更新資料 [" + SelKind.SelectedValue + "." + ListBox1.SelectedValue + "] 完成！');</script>";
                }
                else Msg.Text = "<script>alert('您未改變任何資料！');</script>";
            }
            else
            {
                Msg.Text = "<script>alert('您尚未點選欲修改之資料！');</script>";
            }
        }
        Page.Controls.Add(Msg);
    }

    protected string GetUpdate(string SQLorLife) //取得修改資料的SQL語法
    {
        string SQL = "";
        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
        Conn.Open();
        SqlCommand cmd = new SqlCommand("select * from [Config] where [Kind]='" + SelKind.SelectedValue + "' and [Item]='" + ListBox1.SelectedValue + "'", Conn);
        SqlDataReader dr = null;
        dr = cmd.ExecuteReader();
        if (dr.Read())
        {
            SQL = GetUpdateCol("Item", dr["Item"].ToString(), TextItem.Text, "string", SQLorLife)
                + GetUpdateCol("Config", dr["Config"].ToString(), TextConfig.Text, "string", SQLorLife)
                + GetUpdateCol("Mark", dr["Mark"].ToString(), TextMark.Text, "string", SQLorLife)
                + GetUpdateCol("Memo", dr["Memo"].ToString(), TextMemo.Text.Replace("'", "''"), "string", SQLorLife);                

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
                    case "money": SQL = SQL + ",[" + ColName + "]=" + Target; break;
                    case "null": SQL = SQL + ",[" + ColName + "]=" + null; break;
                    default: SQL = SQL + ",[" + ColName + "]='" + Target + "'"; break;
                }
            }
            else if (SQLorLife == "Life")
            {
                if (Source == "") Source = "(空白)";
                if (Target == "") Target = "(空白)";
                SQL = SQL + ",[" + ColName + "]：" + Source + " -> " + Target;
            }
        }
        return (SQL);
    }

    protected void BtnDel_Click(object sender, EventArgs e)
    {
        Literal Msg = new Literal();

        if (ListBox1.SelectedIndex != -1)
        {
            if (!RightCheck()) Msg.Text = "<script>alert('您沒有權限異動資料，若有需求，請洽機房！');</script>";
            else
            {
                string SQL = "delete Config where Kind='" + SelKind.SelectedValue + "' and Item='" + ListBox1.SelectedValue + "'";

                SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
                Conn.Open();
                SqlCommand cmd = new SqlCommand(SQL, Conn);
                cmd.ExecuteNonQuery();
                cmd.Cancel(); cmd.Dispose(); Conn.Close(); Conn.Dispose();

                string oldMt = ""; if (GetValue("IDMS", "select * from [Config] where [Item]='" + SelKind.SelectedValue + "' and [Kind] in (select [Item] from [Config] where [Kind] in ('組織架構','維護群組'))") != "") oldMt = SelKind.SelectedValue;
                InsLifeLog(SQL + ",[Config]=" + TextConfig.Text + ",[Mark]=" + TextMark.Text,oldMt);

                ListBox1.Items.Remove(ListBox1.SelectedItem);
                ClearData();
                Msg.Text = "<script>alert('刪除資料完成！');</script>";
            }
        }
        else
        {
            Msg.Text = "<script>alert('您尚未點選欲刪除之資料！');</script>";
        }
        Page.Controls.Add(Msg);
    }

    protected void BtnPin_Click(object sender, EventArgs e) //把人從別的課移到這個課
    {
        Literal Msg = new Literal();

        if (TextItem.Text != "")
        {
            string oldClass = GetValue("IDMS", "select [Kind] from [Config] where [Item]='" + TextItem.Text + "' and [Kind]<>'" + SelKind.SelectedValue + "' and [Kind] in (select [成員] from [View_課別] where [成員]<>'')");
            string newClass = GetValue("IDMS", "select [成員] from [View_課別] where [成員]='" + SelKind.SelectedValue + "'");

            if (!RightCheck()) Msg.Text = "<script>alert('您沒有權限異動資料，若有需求，請洽機房！');</script>";
            else if (oldClass == "") Msg.Text = "<script>alert('查不到" + TextItem.Text + "的原來課別！');</script>";
            else if (newClass == "") Msg.Text = "<script>alert('" + SelKind.SelectedValue + "不是合法的課別項目！');</script>";
            else
            {
                string oldMt = ""; if (GetValue("IDMS", "select * from [Config] where [Item]='" + SelKind.SelectedValue + "' and [Kind] in (select [Item] from [Config] where [Kind] in ('組織架構','維護群組'))") != "") oldMt = SelKind.SelectedValue;
                InsLifeLog("修改 [" + oldClass + "." + TextItem.Text + "]  ：： [Kind]：" + oldClass + " -> " + SelKind.SelectedValue,oldMt);
                ExecDbSQL("Update [Config] set [Kind]='" + SelKind.SelectedValue + "' where Kind='" + oldClass + "' and Item='" + TextItem.Text + "'");
                SelKind_SelectedIndexChanged(null,null);

                Msg.Text = "<script>alert('更新資料 [" + SelKind.SelectedValue + "." + TextItem.Text + "] 完成！');</script>";
            }
        }
        else
        {
            Msg.Text = "<script>alert('您尚未填入欲更新之人員！');</script>";
        }
        Page.Controls.Add(Msg);
    }

    protected void ClearData()
    {
        TextItem.Text = "";
        TextConfig.Text = "";
        TextMark.Text = "";
        TextMemo.Text = "";
    }

    protected Boolean RightCheck() //是否有權修改資料
    {
        string UserID = Session["UserID"].ToString().ToLower();
        string UserName = Session["UserName"].ToString();   //登入的UserName
        string UnitName = Session["UnitName"].ToString();   //登入的UnitName
        string ChkUnit=""; //是否有權利異動Config

        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
        Conn.Open();
        SqlCommand cmd = new SqlCommand("select Config from Config where Kind='開放設定' and Item='" + SelKind.SelectedValue + "'", Conn);
        SqlDataReader dr = null;
        dr = cmd.ExecuteReader();
        if (dr.Read()) ChkUnit=dr[0].ToString();
        cmd.Cancel(); cmd.Dispose(); dr.Close(); Conn.Close(); Conn.Dispose();

        if (UserID != "operator" & (UnitName == "電腦操作課" | UnitName == "系統控制課" | UnitName == "技術支援課" | ChkUnit == "資訊中心" | InGroup(UserName, ChkUnit))) return (true);
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

    protected void BtnSearch_Click(object sender, EventArgs e)  //關鍵字查詢
    {
        string[] KeyA = TextKey.Text.Trim().Split(',');
        string SQL = "";

        for (int i = 0; i < KeyA.GetLength(0); i++)
        {
            if (i > 0) SQL = SQL + " and ";

            SQL = SQL + "[Kind]+','+[Item]+','+[Config]+','+[Mark]+','+[Memo] like '%" + KeyA[i] + "%'";
        }
        
        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
        Conn.Open();
        SqlCommand cmd = new SqlCommand("select * from [Config] where " + SQL, Conn);
        SqlDataReader dr = null;
        dr = cmd.ExecuteReader();
        lblKey.Text = "";
        while (dr.Read())
        {
            lblKey.Text = lblKey.Text + dr[0].ToString() + " " + dr[1].ToString() + " " + dr[2].ToString() + " " + dr[3].ToString() + " " + dr[4].ToString() + "<br/><br/>";
        }
        cmd.Cancel(); cmd.Dispose(); dr.Close(); Conn.Close(); Conn.Dispose();
    }
}