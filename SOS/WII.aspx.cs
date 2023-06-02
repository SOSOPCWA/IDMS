using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.Configuration;
using System.Data;
using System.Data.SqlClient;

public partial class SOS_WII : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        if (!IsPostBack)
        {
            GetMap();
            string IP = Request.ServerVariables["REMOTE_ADDR"].ToString();
            if (IP.IndexOf("172.") != 0 & IP.IndexOf("61.60") != 0)
            {
                //Img1.Visible = false;
                //Img2.Visible = false;
                //Img3.Visible = false;
            }
        }
    }

    protected void Page_PreRenderComplete(object sender, EventArgs e)   //放Page_Load無法顯示
    {
        if (!IsPostBack)
        {
            if (Request["SC"] != null)
            {
                GetSC(Request["SC"].ToString());
            }
        }
    }

    protected void GetSC(string SC) //取得測站連結點影像地圖
    {
        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
        Conn.Open();
        SqlCommand cmd = new SqlCommand("select [廠牌],[IP],[型式],[設備用途],[設備備註說明],[設備名稱],[主機名稱],[規格],[設備編號] from [View_通用設備] where [設備種類]='局屬網路' and [主機名稱]='" + SC + "'", Conn);
        SqlDataReader dr = null;
        dr = cmd.ExecuteReader();
        while (dr.Read())
        {
            lblBno.Text = dr[0].ToString();
            lblIP.Text = dr[1].ToString();
            lblOno.Text = dr[2].ToString();
            lblPno.Text = dr[3].ToString();
            lblTel.Text = dr[4].ToString();

            lblHost.Text = dr[5].ToString() + "(" + dr[6].ToString() +")";
            lblHost.ToolTip = dr[7].ToString();

            lnHost.NavigateUrl = "~/IDMS/Device/DevEdit.aspx?DevNo=" + dr[8].ToString();
        }
        cmd.Cancel(); cmd.Dispose(); dr.Close(); Conn.Close(); Conn.Dispose();
    }

    protected void GetMap() //取得測站連結點影像地圖
    {
        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
        Conn.Open();
        SqlCommand cmd = new SqlCommand("select [主機名稱],[規格] from [View_通用設備] where [設備種類]='局屬網路' and [規格]<>''", Conn);
        SqlDataReader dr = null;
        dr = cmd.ExecuteReader();
        while (dr.Read())
        {
            string XYmap = dr[1].ToString();
            int pos1 = XYmap.IndexOf(",");
            int pos2 = XYmap.IndexOf(",", pos1 + 1);
            string SC = dr[0].ToString();
            if (pos1 > 0 & pos2 > 0)
            {
                CircleHotSpot chs = new CircleHotSpot();
                //chs.AlternateText = "newly created hotspot";
                
                chs.NavigateUrl = "WII.aspx?SC=" + SC;

                chs.X = int.Parse( XYmap.Substring(0,pos1));
                chs.Y = int.Parse(XYmap.Substring(pos1+1, pos2-pos1-1));
                chs.Radius = 15;
                ImgWII.HotSpots.Add(chs);
            } 
        }
        cmd.Cancel(); cmd.Dispose(); dr.Close(); Conn.Close(); Conn.Dispose();
    }
    
    protected void OpenExecWindow(string strJavascript)    //選取後就另開視窗
    {
        if (ScriptManager.GetCurrent(Page) == null)
        {
            Page.ClientScript.RegisterStartupScript(Page.GetType(), "WII", strJavascript, true);
        }
        else
        {
            ScriptManager.RegisterStartupScript(Page, Page.GetType(), "WII", strJavascript, true);
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

    protected void BtnKey_Click(object sender, EventArgs e)
    {
        string key = txtKey.Text.Trim();
        lblKey.Text = GetValue("IDMS", "select [設備名稱] from [View_通用設備] where [廠牌] like '%" + key + "%' or [IP] like '%" + key + "%' or [型式] like '%" + key + "%' or [設備用途] like '%" + key + "%' or [設備備註說明] like '%" + key + "%'");
    }
}