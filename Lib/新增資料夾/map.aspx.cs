using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.Configuration;
using System.Data;
using System.Data.SqlClient;

public partial class Lib_map : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        
        if (!IsPostBack)
        {
            Literal Msg = new Literal();
            string PointerNo = Request["PointerNo"];    //Request要在!IsPostBack之內才讀得到            

            if (PointerNo == null & Session["DevSQL"] != null)
            {
                int pos = Session["DevSQL"].ToString().IndexOf("FROM");
                string SQL = "Select * " + Session["DevSQL"].ToString().Substring(pos);
                string Mt = "維護人員"; if (SQL.IndexOf("View_通用設備") > 0) Mt = "設備維護人員";    //設備管理與進階查詢的欄位名稱略有不同
                
                SqlCommand sqlQuery = new SqlCommand(SQL + " and [坐標Y]>-1 order by [坐標X],[坐標Y],[定位方式],[高度] desc");
                DataSet ds = RunQuery(sqlQuery);

                if (ds.Tables.Count > 0)
                {
                    string X = "-1", Y = "-1", Alt = "", strPlace = "";

                    foreach (DataRow row in ds.Tables[0].Rows)
                    {
                        if (X != row["坐標X"].ToString() | Y != row["坐標Y"].ToString())
                        {
                            Msg.Text = Msg.Text + "<img tagName='" + X + "," + Y + "' src='../images/green.gif"
                                + "' style='position:absolute;z-index:2;left:"
                                + (13 * (int.Parse(X) + 3) + 2).ToString() + ";top:" + (14 * int.Parse(Y) + 2).ToString()
                                + "' width='10' height='12' border='0' alt='" + strPlace + "\r\n" + Alt + "' onclick='MapClick()' onMouseMove='MouseMove()'>";
                            X = row["坐標X"].ToString(); Y = row["坐標Y"].ToString();
                            Alt = row["設備編號"].ToString() + ". " + row["設備名稱"].ToString() + "(" + row["設備狀態"].ToString() + ")"
                                + row["設備用途"].ToString() + "-" + row[Mt].ToString();
                            strPlace = row["放置地點"].ToString(); //if (SQL.IndexOf("View_設備管理") > 0) strPlace = row["放置地點"].ToString();
                        }
                        else
                        {
                            Alt = Alt + "\r\n" + row["設備編號"].ToString() + ". " + row["設備名稱"].ToString() + "(" + row["設備狀態"].ToString() + ")"
                                + row["設備用途"].ToString() + "-" + row[Mt].ToString();
                        }
                    }
                    Msg.Text = Msg.Text + "<img tagName='" + X + "," + Y + "' src='../images/green.gif"
                                + "' style='position:absolute;z-index:2;left:"
                                + (13 * (int.Parse(X) + 3) + 2).ToString() + ";top:" + (14 * int.Parse(Y) + 2).ToString()
                                + "' width='10' height='12' border='0' alt='" + strPlace + "\r\n" + Alt
                                + "' alt='" + strPlace + "\r\n" + Alt + "' onclick='MapClick()' onMouseMove='MouseMove()'>";
                    Page.Controls.Add(Msg);
                }
                sqlQuery.Cancel(); ds.Dispose();
            }
        }
    }

    protected void Page_PreRenderComplete(object sender, EventArgs e)   //放Page_Load無法顯示
    {
        //Response.Write(Request["SQL"] + "<br/><br/>");
        //Response.Write(Request["Key"] + "<br/><br/>");
        //Response.End();

        if (!IsPostBack)
        {
            string IO = Request["IO"]; if (IO == null | IO=="") IO = "I";
            string PointerNo = Request["PointerNo"];
            int X = -1, Y = -1;

            if (PointerNo != null & PointerNo !="")
            {
                SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
                Conn.Open();
                SqlCommand cmd = new SqlCommand("select * from [定位設定] where [定位編號]=" + PointerNo, Conn);
                SqlDataReader dr = null;
                dr = cmd.ExecuteReader();
                if (dr.Read())
                {
                    X = int.Parse(dr["坐標X"].ToString());
                    Y = int.Parse(dr["坐標Y"].ToString());
                    if (IO == "I")
                    {
                        pos1.Style.Value = "z-index:2;position:absolute;left:" + ((X + 3) * 13 + 2).ToString() + ";top:" + (Y * 14 + 2).ToString();                        
                    }
                    else
                    {
                        pos2.Style.Value = "z-index:2;position:absolute;left:" + ((X + 3) * 13 + 2).ToString() + ";top:" + (Y * 14 + 2).ToString();
                    }
                }
                cmd.Cancel(); cmd.Dispose(); dr.Close(); Conn.Close(); Conn.Dispose();
            }
        }
    }

    protected DataSet RunQuery(SqlCommand sqlQuery) //讀取節點資訊
    {
        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
        SqlDataAdapter dbAdapter = new SqlDataAdapter();
        dbAdapter.SelectCommand = sqlQuery;
        sqlQuery.Connection = Conn;
        DataSet QueryDataSet = new DataSet();
        //Response.Write(sqlQuery.CommandText);
        //Response.End();
        dbAdapter.Fill(QueryDataSet);
        dbAdapter.Dispose(); Conn.Close(); Conn.Dispose();
        return (QueryDataSet);   
    }
}
