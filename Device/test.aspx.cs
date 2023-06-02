using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Text;
public partial class Device_test : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        Get_list();
        if (!IsPostBack)
        {

        }
    }

    protected void Get_list()
    {
        HashSet<Tuple<string, int>> ups550 = new HashSet<Tuple<string, int>>();
        //Link_UPS("83", ups550, l1);

        //l1.Text += "550 -- " + ups550.Count() + "<br>";
        HashSet<Tuple<string, int>> ups550N = new HashSet<Tuple<string, int>>();
        //Link_UPS("2587", ups550N, l2);
        //l2.Text += "550N -- " + ups550N.Count() + "<br>";
        HashSet<Tuple<string, int>> ups600 = new HashSet<Tuple<string, int>>();
        //Link_UPS("84", ups600, l3);
        HashSet<Tuple<string, int>> ups100 = new HashSet<Tuple<string, int>>();
        //Link_UPS("3603", ups100, l3);
        //l3.Text += "600 -- " + ups600.Count() + "<br>";
        ups550.ExceptWith(ups550N);
        ups550.ExceptWith(ups600);
     
        //l1.Text += "<br>ONLY 550---" + ups550.Count() + "<br>";

    }

    protected void Link_UPS(string UPS, HashSet<Tuple<string, int>> re, object sender)
    {
        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
        Conn.Open();
        string sql = "SELECT * FROM 接電";
        SqlCommand cmd = new SqlCommand(sql, Conn);

        HashSet<Tuple<string, string>> fi = new HashSet<Tuple<string, string>>();
        SqlDataReader dr = cmd.ExecuteReader();
        while (dr.Read())
        {
            //上 下
            fi.Add(Tuple.Create(dr[0].ToString(), dr[1].ToString()));
        }


        Stack<Tuple<string, int>> un_def = new Stack<Tuple<string, int>>();
        un_def.Push(Tuple.Create(UPS, 1));
        while (un_def.Any())
        {
            var tar = un_def.Pop();
            foreach (var item in fi)
            {
                if (item.Item1 == tar.Item1)
                { //TOP == tar
                    re.Add(Tuple.Create(item.Item2, tar.Item2 + 1));
                    ((Label)sender).Text += "<br>(" + item.Item1 + "," + item.Item2 + ")" + tar.Item2;
                    un_def.Push(Tuple.Create(item.Item2, tar.Item2 + 1));
                }
            }
        }
        
        cmd.Cancel(); cmd.Dispose(); dr.Close(); Conn.Close(); Conn.Dispose();
    }

    protected static void Link_UPS(string UPS, HashSet<Tuple<string, int>> re)
    {
        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
        Conn.Open();
        string sql = "SELECT * FROM 接電";
        SqlCommand cmd = new SqlCommand(sql, Conn);

        HashSet<Tuple<string, string>> fi = new HashSet<Tuple<string, string>>();
        SqlDataReader dr = cmd.ExecuteReader();
        while (dr.Read())
        {
            fi.Add(Tuple.Create(dr[0].ToString(), dr[1].ToString()));
        }


        Stack<Tuple<string, int>> un_def = new Stack<Tuple<string, int>>();
        un_def.Push(Tuple.Create(UPS, 1)); re.Add(Tuple.Create(UPS, 1));
        while (un_def.Any())
        {
            var tar = un_def.Pop();
            foreach (var item in fi)
            {
                if (item.Item1 == tar.Item1)
                { //TOP == tar
                    re.Add(Tuple.Create(item.Item2, tar.Item2 + 1));
                    //((Label)sender).Text += "<br>" + item.Item2 + "," + tar.Item2;                 
                    un_def.Push(Tuple.Create(item.Item2, tar.Item2 + 1));
                }
            }
        }
        
        cmd.Cancel(); cmd.Dispose(); dr.Close(); Conn.Close(); Conn.Dispose();
    }

    protected static void Link_UPS_withname(string UPS, HashSet<Tuple<string, string,string>> re)
    {
        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
        Conn.Open();
        string sql = @"select 上游編號, 下游編號, 設備名稱 from 接電 t1, View_設備管理 t2
                            where t1.下游編號 = t2.設備編號";
        SqlCommand cmd = new SqlCommand(sql, Conn);

        HashSet<Tuple<string, string, string>> fi = new HashSet<Tuple<string, string, string>>();
        SqlDataReader dr = cmd.ExecuteReader();
        while (dr.Read())
        {
            //上 下 下名
            fi.Add(Tuple.Create(dr[0].ToString(), dr[1].ToString(), dr[2].ToString() ));
        }


        Stack<Tuple<string>> un_def = new Stack<Tuple<string>>();
        //UPS先放入
        un_def.Push(Tuple.Create(UPS)); 
        while (un_def.Any())
        {
            var tar = un_def.Pop();
            foreach (var item in fi)
            {
                if (item.Item1 == tar.Item1)
                { //TOP == tar
                    re.Add(Tuple.Create(item.Item1, item.Item2, item.Item3));
                    //((Label)sender).Text += "<br>" + item.Item2 + "," + tar.Item2;                 
                    un_def.Push(Tuple.Create(item.Item2));
                }
            }
        }

        cmd.Cancel(); cmd.Dispose(); dr.Close(); Conn.Close(); Conn.Dispose();
    }

    public static string Get_Json()
    {
        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
        Conn.Open();
        string sql = @"select COUNT(*) from 接電 t1, View_設備管理 t2
                            where t1.下游編號 = t2.設備編號";
        SqlCommand cmd = new SqlCommand(sql, Conn);
        SqlDataReader dr = null;
        dr = cmd.ExecuteReader();
        int count = 0;
        if (dr.Read())
        {
            count = int.Parse(dr[0].ToString());
        }
        cmd.CommandText = @"select 上游編號,下游編號,設備名稱 from 接電 t1, View_設備管理 t2
                            where t1.下游編號 = t2.設備編號";
        dr.Close();
        dr = cmd.ExecuteReader();


        StringBuilder myStringBuilder = new StringBuilder("[");

        while (dr.Read() & count > 0)
        {

            myStringBuilder.Append("{\"TOP\":\"" + dr[0].ToString() +

                "\",\"BOT\":\"" + dr[1].ToString() +
                "\",\"Name\":\"" + dr[2].ToString()
                );
            if (count == 1)
            {
                myStringBuilder.Append("\"}");

            }
            else
            {
                myStringBuilder.Append("\"},");
            }
            count -= 1;
        }
        cmd.Cancel(); cmd.Dispose(); dr.Close(); Conn.Close(); Conn.Dispose();
        return myStringBuilder.ToString();
    }

    public static string Get_Json_UPS()
    {
        HashSet<Tuple<string, int>> ups550 = new HashSet<Tuple<string, int>>();
        Link_UPS("83", ups550);
        HashSet<Tuple<string, int>> ups550N = new HashSet<Tuple<string, int>>();
        Link_UPS("2587", ups550N);       
        HashSet<Tuple<string, int>> ups600 = new HashSet<Tuple<string, int>>();
        Link_UPS("84", ups600);
        ups550.ExceptWith(ups550N);
        ups550.ExceptWith(ups600);
        int count = ups550.Count();

        StringBuilder myStringBuilder = new StringBuilder("[");

        foreach(var i in ups550)
        {

            myStringBuilder.Append("{\"TOP\":\"" + i.Item1.ToString() +

                "\",\"BOT\":\"" + i.Item1.ToString() +
                "\",\"Name\":\"" + i.Item2.ToString()
                );
            if (count == 1)
            {
                myStringBuilder.Append("\"}");

            }
            else
            {
                myStringBuilder.Append("\"},");
            }
            count -= 1;
        }
        
        return myStringBuilder.ToString();
    }

    public static string Write_json()
    {
        HashSet<Tuple<string, int>> ups550 = new HashSet<Tuple<string, int>>();
        Link_UPS("83", ups550);
        int count = ups550.Count();
        StringBuilder myStringBuilder = new StringBuilder("[");

        foreach (var item in ups550)
        {
            myStringBuilder.Append("{\"dev\":\"" + item.Item1 +

                "\",\"level\":\"" + item.Item2);
            if (count == 1)
            {
                myStringBuilder.Append("\"}");
            }
            else
            {
                myStringBuilder.Append("\"},");
            }
            count -= 1;
        }
        

        return myStringBuilder.ToString();
    }

    public static string Get_Json2(int s) //{top, bot, name}
    {
        
        HashSet<Tuple<string, string, string>> ups550 = new HashSet<Tuple<string, string, string>>();
        Link_UPS_withname("83", ups550);

        StringBuilder myStringBuilder = new StringBuilder("[");
        myStringBuilder.Append(
            "{\"id\":\"83\"" +
            ",\"description\":\"UPS550\"" +
            ",\"type\":\"Root\"},"
            );
        int count = ups550.Count();
       foreach(var dev in ups550)
        {
            myStringBuilder.Append(
                "{\"id\":\"" + dev.Item2 +
                "\",\"parentId\":\"" + dev.Item1 +
                "\",\"type\":\"" + dev.Item2 +
                "\",\"description\":\"" + dev.Item3
                
                );
            if (count == 1)
            {
                myStringBuilder.Append("\"}");
            }
            else
            {
                myStringBuilder.Append("\"},");
            }
            count -= 1;
        }
        
        return myStringBuilder.ToString();
    }

}