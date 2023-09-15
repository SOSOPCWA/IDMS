using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.Configuration;
using System.Data.SqlClient;

public partial class _Default : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        string UserID = "", UserName = "", UnitID = "", UnitName = "";

        try
        {
            UserID = Request.Cookies["UserID"].Value.ToString();    //自SSM首頁取得
            if (UserID == "") ResponseLogin(UserID);
            else
            {
                SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
                Conn.Open();
                SqlCommand cmd = new SqlCommand("Select Kind,Item from Config where Mark='" + UserID + "'", Conn);
                SqlDataReader dr = null;
                dr = cmd.ExecuteReader();
                if (dr.Read())
                {
                    UserName = dr[1].ToString(); UnitName = dr[0].ToString(); dr.Close();
                    cmd.CommandText = "Select Config from Config where Kind='數值資訊組' and Item='" + UnitName + "'";
                    dr = cmd.ExecuteReader();
                    if (dr.Read())
                    {
                        UnitID = dr[0].ToString();
                        Session["UserID"] = UserID;
                        Session["UserName"] = UserName;
                        Session["UnitID"] = UnitID;
                        Session["UnitName"] = UnitName;
                    }
                    else ResponseLogin(UserID);
                }
                else
                {
                    ResponseLogin(UserID);
                }
                cmd.Cancel(); cmd.Dispose(); dr.Close(); Conn.Close(); Conn.Dispose();
            }            
        }
        catch 
        {            
            ResponseLogin(UserID);
        }
    }

    protected void ResponseLogin(string UserID)
    {
        Response.Write("無法取得認證帳號(" + UserID + ")，請先<a href='/' target='_top'>登入Windows網域</a>！若無法解決，請重啟瀏覽器。");
        Response.End();
    }
}