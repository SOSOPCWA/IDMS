using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.Configuration;
using System.Data;
using System.Data.SqlClient;

public partial class Life_Repeat : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        Session.Timeout = 720;
        if (!IsPostBack)
        {

        }
    }

    protected void Page_PreRenderComplete(object sender, EventArgs e)   //放Page_Load無法顯示
    {
        if (!IsPostBack)
        {
            string Unit = "",Mt = "";
            if (Request["Mt"] != null)  //外部人員帶入查詢
            {
                Mt = Request["Mt"].ToString();
                Unit = GetValue("IDMS", "select [課別] from [View_員工] where [成員]='" + Mt + "'");
                for (int i = 0; i < SelUnit.Items.Count; i++) 
                {
                    if (SelUnit.Items[i].Value != Unit) SelUnit.Items[i].Selected = false;
                    else SelUnit.Items[i].Selected = true;
                }
            }
            SelMt.DataBind();
            for (int i = 0; i < SelMt.Items.Count; i++)
            {
                if (SelMt.Items[i].Value != Mt) SelMt.Items[i].Selected = false;
                else SelMt.Items[i].Selected = true;
            }

            SearchClick();
        }
    }

    protected void Button1_Click(object sender, EventArgs e)
    {
        SearchClick();
    }

    protected void SearchClick()
    {
        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
        SqlConnection Conn1 = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
        Conn.Open(); Conn1.Open();
        SqlCommand cmd = new SqlCommand("", Conn);
        SqlCommand cmd1 = new SqlCommand("", Conn1);
        SqlDataReader dr = null, dr1 = null;
        Boolean flag = false;
        TextCheck.Text = "";
        string SQL = "";
        string SqlMt = " and ([維護人員]='" + SelMt.SelectedValue + "' or [維護人員] in (select Kind from Config where Item='" + SelMt.SelectedValue + "'))";
        //--------------------------------------------------------------------------------------------------------------------------
        if (SelRepeat.SelectedValue == "全部檢查" | SelRepeat.SelectedValue == "設備名稱")
        {
            TextCheck.Text = TextCheck.Text + "檢查 [實體設備] 中有重複的 [設備名稱] (僅檢查 系統設定>設備型態>設備種類 中設定為 唯一 ，非 重複 者)\n";
            TextCheck.Text = TextCheck.Text + "--------------------------------------------------------------------------------- \n";
            cmd.CommandText = "select [設備名稱],count(*) as CountXXX from [View_設備管理] where [設備名稱]<>'' group by [設備名稱]";
            dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                if (dr["CountXXX"].ToString() != "1")
                {
                    SQL = "select * from [View_設備管理] where [設備名稱] ='" + dr["設備名稱"]
                        + "' and [設備種類] not in (select [Item] from [Config] where [Kind] in (select [Item] from [Config] where [Kind]='設備型態') and [Config]='重複')";
                    if (SelMt.SelectedValue != "") SQL = SQL + SqlMt;
                    cmd1.CommandText = SQL;
                    dr1 = cmd1.ExecuteReader();
                    flag = false;
                    while (dr1.Read())
                    {
                        TextCheck.Text = TextCheck.Text + dr1["設備編號"].ToString() + ". "
                            + dr1["設備型態"] + "　"
                            + dr1["設備種類"] + "　"
                            + dr1["設備名稱"] + "　"
                            + dr1["財產編號"] + "　"
                            + dr1["廠牌"] + "　"
                            + dr1["型式"] + "　"
                            + dr1["維護人員"] + "\n";
                        flag = true;
                    }
                    cmd1.Cancel(); cmd1.Dispose(); dr1.Close();
                    if (flag) TextCheck.Text = TextCheck.Text + "\n";
                }
            }
            cmd.Cancel(); cmd.Dispose(); dr.Close();
            TextCheck.Text = TextCheck.Text + "\n\n";
        }
        //--------------------------------------------------------------------------------------------------------------------------
        if (SelRepeat.SelectedValue == "全部檢查" | SelRepeat.SelectedValue == "財產編號")
        {
            TextCheck.Text = TextCheck.Text + "檢查 [實體設備] 中有重複的 [財產編號] (財產編號A<>00000000000)\n";
            TextCheck.Text = TextCheck.Text + "--------------------------------------------------------------------------------- \n";
            cmd.CommandText = "select [財產編號],count(*) as CountXXX from [View_設備管理] where [財產編號A]<>'' and [財產編號B]<>'' and [財產編號] not like '00000000000%' group by [財產編號]";
            dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                if (dr["CountXXX"].ToString() != "1")
                {
                    SQL = "select * from [View_設備管理] where [財產編號] ='" + dr["財產編號"] + "'";
                    if (SelMt.SelectedValue != "") SQL = "select * from [View_設備管理] where [財產編號] in (select distinct [財產編號] from [View_設備管理] where [財產編號]='" + dr["財產編號"] + "'" + SqlMt + ")";
                    cmd1.CommandText = SQL;
                    dr1 = cmd1.ExecuteReader();
                    flag = false;
                    while (dr1.Read())
                    {
                        TextCheck.Text = TextCheck.Text + dr1["設備編號"].ToString() + ". "
                            + dr1["設備型態"] + "　"
                            + dr1["設備種類"] + "　"
                            + dr1["設備名稱"] + "　"
                            + dr1["財產編號"] + "　"
                            + dr1["廠牌"] + "　"
                            + dr1["型式"] + "　"
                            + dr1["維護人員"] + "\n";
                        flag = true;
                    }
                    cmd1.Cancel(); cmd1.Dispose(); dr1.Close();
                    if (flag) TextCheck.Text = TextCheck.Text + "\n";
                }
            }
            cmd.Cancel(); cmd.Dispose(); dr.Close();
            TextCheck.Text = TextCheck.Text + "\n\n";
        }
        //--------------------------------------------------------------------------------------------------------------------------
        if (SelRepeat.SelectedValue == "全部檢查" | SelRepeat.SelectedValue == "重複安裝")
        {
            TextCheck.Text = TextCheck.Text + "檢查 [作業主機] 中有重複的 [設備編號] (不含虛擬主機)\n";
            TextCheck.Text = TextCheck.Text + "--------------------------------------------------------------------------------- \n";
            cmd.CommandText = "SELECT [設備編號],COUNT([設備編號]) as [CountXXX] FROM [作業主機]"
                + " where [設備編號]>0 and [設備編號] not in (select [設備編號] from [實體設備] where [設備種類]='虛擬主機')"
                + " group by [設備編號] having COUNT([設備編號])>1";
            dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                SQL = "select * from [作業主機] where [設備編號] =" + dr["設備編號"];
                if (SelMt.SelectedValue != "") SQL = "select * from [作業主機] where [設備編號] in (select distinct [設備編號] from [作業主機] where [設備編號]=" + dr["設備編號"] + SqlMt + ")";
                cmd1.CommandText = SQL;
                dr1 = cmd1.ExecuteReader();
                flag = false;
                while (dr1.Read())
                {
                    TextCheck.Text = TextCheck.Text + dr1["作業編號"].ToString() + ". "
                        + dr1["主機名稱"] + "　"
                        + dr1["IP"] + "　"
                        + GetSysAlias(dr1["系統編號"].ToString(), "SysName") + "　"
                        + dr1["主要功能"] + "　"
                        + dr1["維護人員"] + "\n";
                    flag = true;
                }
                cmd1.Cancel(); cmd1.Dispose(); dr1.Close();
                if (flag) TextCheck.Text = TextCheck.Text + "\n";
            }
            cmd.Cancel(); cmd.Dispose(); dr.Close();
            TextCheck.Text = TextCheck.Text + "\n\n";
        }
        //--------------------------------------------------------------------------------------------------------------------------
        if (SelRepeat.SelectedValue == "全部檢查" | SelRepeat.SelectedValue == "主機名稱")
        {
            TextCheck.Text = TextCheck.Text + "檢查 [作業主機] 中不同安內外有重複的 [主機名稱] \n";
            TextCheck.Text = TextCheck.Text + "--------------------------------------------------------------------------------- \n";
            cmd.CommandText = "select [主機名稱],[安內外],count(*) as CountXXX from [作業主機] where [主機名稱]<>''"
                + " group by [主機名稱],[安內外]";
            dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                if (dr["CountXXX"].ToString() != "1")
                {
                    SQL = "select * from [作業主機] where [主機名稱] ='" + dr["主機名稱"] + "'";
                    if (SelMt.SelectedValue != "") SQL = "select * from [作業主機] where [主機名稱] in (select distinct [主機名稱] from [作業主機] where [主機名稱]='" + dr["主機名稱"] + "'" + SqlMt + ")";
                    cmd1.CommandText = SQL;
                    dr1 = cmd1.ExecuteReader();
                    flag = false;
                    while (dr1.Read())
                    {
                        TextCheck.Text = TextCheck.Text + dr1["作業編號"].ToString() + ". "
                            + dr1["主機名稱"] + "　"
                            + dr1["IP"] + "　" + dr1["安內外"] + "　" + dr1["作業狀態"] + "　"
                            + GetSysAlias(dr1["系統編號"].ToString(), "SysName") + "　"
                            + dr1["主要功能"] + "　"
                            + dr1["維護人員"] + "\n";
                        flag = true;
                    }
                    cmd1.Cancel(); cmd1.Dispose(); dr1.Close();
                    if (flag) TextCheck.Text = TextCheck.Text + "\n";
                }
            }
            cmd.Cancel(); cmd.Dispose(); dr.Close();
            TextCheck.Text = TextCheck.Text + "\n\n";
        }
        //--------------------------------------------------------------------------------------------------------------------------
        if (SelRepeat.SelectedValue == "全部檢查" | SelRepeat.SelectedValue == "IP位址")
        {
            TextCheck.Text = TextCheck.Text + "檢查 [作業主機] 中有重複的 [IP] \n";
            TextCheck.Text = TextCheck.Text + "--------------------------------------------------------------------------------- \n";
            cmd.CommandText = "select [IP],count(*) as CountXXX from [作業主機] where [IP]<>'' group by [IP]";
            dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                if (dr["CountXXX"].ToString() != "1")
                {
                    SQL = "select * from [作業主機] where [IP] ='" + dr["IP"] + "'";
                    if (SelMt.SelectedValue != "") SQL = "select * from [作業主機] where [IP] in (select distinct [IP] from [作業主機] where [IP]='" + dr["IP"] + "'" + SqlMt + ")";
                    cmd1.CommandText = SQL;
                    dr1 = cmd1.ExecuteReader();
                    flag = false;
                    while (dr1.Read())
                    {
                        TextCheck.Text = TextCheck.Text + dr1["作業編號"].ToString() + ". "
                            + dr1["主機名稱"] + "　"
                            + dr1["IP"] + "　"
                            + GetSysAlias(dr1["系統編號"].ToString(), "SysName") + "　"
                            + dr1["主要功能"] + "　"
                            + dr1["維護人員"] + "\n";
                        flag = true;
                    }
                    cmd1.Cancel(); cmd1.Dispose(); dr1.Close();
                    if (flag) TextCheck.Text = TextCheck.Text + "\n";
                }
            }
            cmd.Cancel(); cmd.Dispose(); dr.Close();
            TextCheck.Text = TextCheck.Text + "\n\n";
        }
        //------------------------------------------------------------------------------------------
        Conn.Close(); Conn.Dispose(); Conn1.Close(); Conn1.Dispose();
    }

    protected void SelUnit_SelectedIndexChanged(object sender, EventArgs e)
    {
        SelMt.DataBind();
    }

    protected string GetSysAlias(string SysNo, string Kind)   //取得系統相關資訊
    {
        if (Kind == "SysName") return (GetValue("IDMS", "select [系統全名] from [View_系統資源] where [資源編號]=" + SysNo));
        else return (GetValue("IDMS", "select [資源功能]" + " -- " + "[備註說明] from [系統資源] where [資源編號]=" + SysNo));
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