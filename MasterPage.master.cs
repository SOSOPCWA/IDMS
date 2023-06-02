using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.Configuration;
using System.Data;
using System.Data.SqlClient;

public partial class MasterPage : System.Web.UI.MasterPage
{
    protected void Page_Load(object sender, EventArgs e)
    {
        if (!Page.IsPostBack)
        {
            if (Session["UserID"] == null | Session["UserID"] == "" | Session["UserID"] == "None")
            {
                Response.Redirect("~/IDMS");
            }
            else
            {
                LblUserName.Text = Session["UserName"] + "(" + Session["UserID"] + ")" + " / " + Session["UnitName"] + "(" + Session["UnitID"] + ")";
                AddMenu("設備管理", "設備型態", 0,0);
                AddMenu("系統設定", "組織架構", 7, 0);
                AddMenu("系統設定", "維護群組", 7, 1);

                if (Request.Url.Port == 12345 | Request.ServerVariables["LOCAL_ADDR"].ToString() != "10.6.1.11")
                {
                    lblTest.Visible = true; //標題顯示測試機
                }
            }
        }
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

    protected void AddMsg(string strMsg)
    {
        Literal Msg = new Literal();
        Msg.Text = strMsg;
        Page.Controls.Add(Msg);
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

    protected void AddMenu(string SubSys, string kind, int idx1, int idx2)
    {
        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
        Conn.Open();
        SqlCommand cmd = new SqlCommand("Select [Item],[Config],[Memo] from [Config] where [Kind]='" + kind + "' order by [mark]", Conn);
        SqlDataReader dr = null;
        dr = cmd.ExecuteReader();
        while (dr.Read())
        {
            MenuItem itemx = new MenuItem();
            itemx.Text = dr[0].ToString();
            itemx.Value = dr[0].ToString();
            //itemx.ToolTip = dr[2].ToString();
            switch (SubSys)   //按子系統做不同Menu之處理
            {
                case "設備管理":
                    itemx.NavigateUrl = "Device/device.aspx?DevType=" + dr[0].ToString();
                    Menu1.Items[idx1].ChildItems.Add(itemx);
                    break;
                case "系統設定":
                    itemx.NavigateUrl = "Config/Config.aspx?Kind=" + dr[0].ToString();
                    Menu1.Items[idx1].ChildItems[idx2].ChildItems.Add(itemx);
                    break;
            }
        }
        cmd.Cancel(); cmd.Dispose(); dr.Close(); Conn.Close(); Conn.Dispose();
    }

    protected void Menu_MenuItemClick(Object sender, MenuEventArgs e)   //導入詞彙
    {
        if (e.Item.Value == "機房平面圖") Session["DevSQL"] = "SELECT * FROM [View_設備管理] WHERE 1=1";
        else if (e.Item.Text == "欄位說明") AddMsg("<script>alert('" + e.Item.ToolTip + "');</script>");
    }
}
