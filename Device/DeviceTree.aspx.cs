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
        //Get_list();
		
        if (!IsPostBack)
        {
			ViewState["UPS"] = Request["UPS"];
			if(Request["UPS"]==null) ViewState["UPS"] ="UPS 825";
        }
    }

    public string Get_list()
    {
        string ans = "";
        HashSet<Tuple<string, int>> ups550 = new HashSet<Tuple<string, int>>();
        Link_UPS("83", ups550);

        var c = ups550.Count().ToString();
        //l1.Text += "550 -- " + ups550.Count() + "<br>";
        HashSet<Tuple<string, int>> ups550N = new HashSet<Tuple<string, int>>();
        Link_UPS("2587", ups550N);
        //l2.Text += "550N -- " + ups550N.Count() + "<br>";
        HashSet<Tuple<string, int>> ups600 = new HashSet<Tuple<string, int>>();
        Link_UPS("84", ups600);

        //l3.Text += "600 -- " + ups600.Count() + "<br>";
        ups550.ExceptWith(ups550N);
        ups550.ExceptWith(ups600);
        return c + "(550)/"+ups550.Count().ToString() +"(550O)/"+ ups550N.Count().ToString() +"(550N)/"+ ups600.Count().ToString()+"(600)";
        //l1.Text += "<br>ONLY 550---" + ups550.Count() + "<br>";
        
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
        un_def.Push(Tuple.Create(UPS, 1)); 
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

    protected static void Link_UPS_withname2(string UPS, HashSet<Tuple<string,string>> re)
    {
        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
        Conn.Open();
        string sql = @"select 上游編號, 下游編號 from 接電 t1";
        SqlCommand cmd = new SqlCommand(sql, Conn);

        HashSet<Tuple<string, string>> fi = new HashSet<Tuple<string, string>>();
        SqlDataReader dr = cmd.ExecuteReader();
        while (dr.Read())
        {
            //上 下 下名
            fi.Add(Tuple.Create(dr[0].ToString(), dr[1].ToString() ));
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
                    re.Add(Tuple.Create(item.Item1, item.Item2));                              
                    un_def.Push(Tuple.Create(item.Item2));
                }
            }
        }

        cmd.Cancel(); cmd.Dispose(); dr.Close(); Conn.Close(); Conn.Dispose();
    }
    protected static void Link_UPS_withname3(string UPS, HashSet<Tuple<string, string,string,string>> re)
    {
        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
        Conn.Open();
        string sql = @"select 上游編號, 下游編號, 設備名稱, 放置地點 from 接電 t1, View_設備管理 t2
                            where t1.下游編號 = t2.設備編號";
        SqlCommand cmd = new SqlCommand(sql, Conn);
        //所有設備放入set
        HashSet<Tuple<string, string, string, string>> fi = new HashSet<Tuple<string, string, string, string>>();
        
        SqlDataReader dr = cmd.ExecuteReader();
        while (dr.Read())
        {
            //所有設備----{上 this  this.name 地點}
            fi.Add(Tuple.Create(dr[0].ToString(), dr[1].ToString(), dr[2].ToString(), dr[3].ToString() ));
        }

        //未處理完的設備
        Stack<Tuple<string>> un_def = new Stack<Tuple<string>>();
        //重複出現{ID,次數}
        HashSet<Tuple<string,int>> ID_exist = new HashSet<Tuple<string,int>>();
        //UPS先放入
        un_def.Push(Tuple.Create(UPS)); 
        while (un_def.Any())
        {
            var tar = un_def.Pop();
            var Name = "";
            foreach (var item in fi)//巡視所有設備
            {
                if (item.Item1 == tar.Item1) //巡到的設備上游 == tar
                {
                    //確認有無重複
                    var tag_check_rep = true;
                    foreach(var exis in ID_exist){
                        if(exis.Item1 == item.Item2){ //重複set裡有出現過this.ID
                            Name = item.Item2 + "-" + exis.Item2;
                            ID_exist.Add(Tuple.Create(item.Item2,exis.Item2+1));
                            ID_exist.Remove(exis);
                            tag_check_rep = false;
                            break;
                        }
                    }
                    if(tag_check_rep){
                        Name = item.Item2;
                        ID_exist.Add(Tuple.Create(item.Item2,1));
                    }
                    
                                    
                    //{上游 this.ID this.Name this.location }
                    re.Add(Tuple.Create(item.Item1, Name, item.Item3, item.Item4));
                    //只推進stack this.ID
                    un_def.Push(Tuple.Create(item.Item2));
                }
                
                
            }
        }

        cmd.Cancel(); cmd.Dispose(); dr.Close(); Conn.Close(); Conn.Dispose();
    }
    public string TestF(){
        return ViewState["UPS"].ToString();
    }
    protected static void Link_UPS_nested(string UPS, HashSet<Tuple<string, string,string,string>> re)
    {
        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
        Conn.Open();
        string sql = @"select 上游編號, 下游編號, 設備名稱, 放置地點 from 接電 t1, View_設備管理 t2
                            where t1.下游編號 = t2.設備編號";
        SqlCommand cmd = new SqlCommand(sql, Conn);

        HashSet<Tuple<string, string, string, string>> fi = new HashSet<Tuple<string, string, string, string>>();
        SqlDataReader dr = cmd.ExecuteReader();
        while (dr.Read())
        {
            //上 下 下名
            fi.Add(Tuple.Create(dr[0].ToString(), dr[1].ToString(), dr[2].ToString(), dr[3].ToString() ));
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
                    re.Add(Tuple.Create(item.Item1, item.Item2, item.Item3, item.Item4));                              
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

    public static string NewWay()
    {

        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
        Conn.Open();
        string sql = "SELECT * FROM 接電";
        SqlCommand cmd = new SqlCommand(sql, Conn);
        //全部物件set
        List<Tuple<string, string>> fi = new List<Tuple<string, string>>();
        SqlDataReader dr = cmd.ExecuteReader();
        while (dr.Read())
        {
            //上 下
            fi.Add(Tuple.Create(dr[0].ToString(), dr[1].ToString()));
        }
        //上下游用的關係組合
        HashSet<int> LiUPS550 = new HashSet<int>();
        string UPS550_s = "83";
        HashSet<int> LiUPS550N = new HashSet<int>();
        string UPS550N_s = "2587";
        HashSet<int> LiUPS600 = new HashSet<int>();
        string UPS600_s = "84";
        NewWay_UPS( UPS550_s, LiUPS550, fi);
        NewWay_UPS( UPS550N_s, LiUPS550N, fi);
        NewWay_UPS(UPS600_s, LiUPS600, fi);
        LiUPS550.ExceptWith(LiUPS550N);
        LiUPS550.ExceptWith(LiUPS600);
        
        
        SqlConnection Conn2 = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
        Conn2.Open();
        string sql2 = @"select 上游編號, 下游編號, 設備名稱, 放置地點 from 接電 t1, View_設備管理 t2
                            where t1.下游編號 = t2.設備編號 AND t1.下游編號=@dno";
        SqlCommand cmd2 = new SqlCommand(sql2, Conn2);
        HashSet<Tuple<string, string,string,string>> UPS550 = new HashSet<Tuple<string, string,string,string>>();
        SqlDataReader dr2;
        foreach (var item in LiUPS550)
        {
            cmd2.Parameters.Clear();
            cmd2.Parameters.AddWithValue("@dno", item);
            dr2 = cmd2.ExecuteReader();
            if (dr2.Read())
            UPS550.Add(Tuple.Create(dr2[0].ToString(), dr2[1].ToString(), dr2[2].ToString(), dr2[3].ToString()));

            dr2.Close();
        }
        
        int count = UPS550.Count();
        StringBuilder myStringBuilder = new StringBuilder("[");
        myStringBuilder.Append(
            "{\"id\":\"" + UPS550_s +"\"" +
            ",\"description\":\"" + "UPS550" +"\"" +
            ",\"type\":\"Root\"" +
            ",\"loca\":\"B1\" },"
            );
        foreach(var item in UPS550)
        {
            myStringBuilder.Append(
                "{\"id\":\"" + item.Item2 +
                "\",\"parentId\":\"" + item.Item1 +
                "\",\"type\":\"" + item.Item2 +
                "\",\"description\":\"" + item.Item3 +
                "\",\"loca\":\"" + item.Item4 
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

    protected static void NewWay_UPS(string UPS, HashSet<int> LiUPS,  List<Tuple<string, string>> fi)
    {
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
                    LiUPS.Add(Int32.Parse(item.Item2));
                    un_def.Push(Tuple.Create(item.Item2));
                }
            }
        }
    }



    public string Get_Json3() //{top, bot, name}
    {
        //寫成json的set
        HashSet<Tuple<string, string, string,string>> ups = new HashSet<Tuple<string, string, string,string>>();
        
        string which = ViewState["UPS"].ToString();
        //起始設備編號
        var UUU = "";var temp="";
        bool only = false;
        switch (which)
        {
            case "UPS 825":
                UUU = "4216";
                break;
            case "550O": //550單接 ---- no use
                UUU = "83";
                return NewWay();
                break;
            case "UPS 550(B)":
                UUU = "4217";
                break;
            case "UPS 550(C)":
                UUU = "4222";
                break;
            case "UPS 550(D)":
                UUU = "2587";
                break;
			case "UPS 160":
                UUU = "3742";
                break;
			case "UPS 175":
                UUU = "3741";
                break;
			case "UPS 60":
                UUU = "3106";
                break;
            case "UPS 60(O+)":
                UUU = "3743";
                break;
            case "市電":
                UUU = "1299";
                break;
            case "發電機":
                UUU = "1298";
                break;
            case "總電源":
                UUU = "-2";
                break;
            default:
                return "[";
                break;
        }
        //從ROOT往下建立set
        Link_UPS_withname3(UUU, ups);
        if(only){ //unuse
            HashSet<Tuple<string, string, string,string>> ups550N = new HashSet<Tuple<string, string, string,string>>();
            HashSet<Tuple<string, string, string,string>> ups600 = new HashSet<Tuple<string, string, string,string>>();
            Link_UPS_withname3("2587", ups550N);
            temp += ups.Count().ToString() +"(550)/";
            temp += ups550N.Count().ToString() +"(550N)/";
            Link_UPS_withname3("84", ups600);
            temp += ups600.Count().ToString() +"(600)/";
            ups.ExceptWith(ups550N);
            ups.ExceptWith(ups600);
            temp += ups.Count().ToString() +"(550O)/";
        }
        
        
        
        StringBuilder myStringBuilder = new StringBuilder("[");
        myStringBuilder.Append(
            "{\"id\":\"" + UUU +"\"" +
            ",\"description\":\"" + which +"\"" +
            ",\"type\":\"Root\"" +
            ",\"loca\":\"B1\" }"
            );
        
        int count = ups.Count();
        
        foreach(var dev in ups)
        {            
            myStringBuilder.Append(
                ",{\"id\":\"" + dev.Item2 +
                "\",\"parentId\":\"" + dev.Item1 +
                "\",\"type\":\"" + dev.Item2 +
                "\",\"description\":\"" + dev.Item3.Replace("\"", " ") +
                "\",\"loca\":\"" + dev.Item4 +
                "\"}"              
                );
        }
        myStringBuilder.Replace("	", ""); //去除空白
        return myStringBuilder.ToString();
    }
    public  string Get_Json2() //{top, bot, name}
    {
        
        HashSet<Tuple<string, string>> ups = new HashSet<Tuple<string, string>>();
        string which = ViewState["UPS"].ToString();
        var UUU = "";var temp="";
        bool only = false;
        switch (which)
        {
            case "550":
                UUU = "83";
                break;
            case "550O":
                UUU = "83";only=true;
                break;
            case "550N":
                UUU = "2587";
                break;
            case "600":
                UUU = "84";
                break;
			case "100":
                UUU = "3104";
                break;
			case "175":
                UUU = "3603";
                break;
			case "60":
                UUU = "3106";
                break;
            case "-2":
                UUU = "-2";
                break;
            default:
                return "[";
                break;
        }
       
        Link_UPS_withname2(UUU, ups);
        HashSet<Tuple<string, string>> ups550N = new HashSet<Tuple<string, string>>();
        HashSet<Tuple<string, string>> ups600 = new HashSet<Tuple<string, string>>();
        Link_UPS_withname2("2587", ups550N);
        temp += ups.Count().ToString() +"(550)/";
        temp += ups550N.Count().ToString() +"(550N)/";
        Link_UPS_withname2("84", ups600);
        temp += ups600.Count().ToString() +"(600)/";
        ups.ExceptWith(ups550N);
        ups.ExceptWith(ups600);
        temp += ups.Count().ToString() +"(550O)/";
        return temp;
        //Check_Tree(ups);
        StringBuilder myStringBuilder = new StringBuilder("[");
        myStringBuilder.Append(
            "{\"id\":\"" + UUU +"\"" +
            ",\"description\":\"" + "UPS"+which +"\"" +
            ",\"type\":\"Root\"},"
            );
        int count = ups.Count();
        
        foreach(var dev in ups)
        {
            
            myStringBuilder.Append(
                "{\"id\":\"" + dev.Item2 +
                "\",\"parentId\":\"" + dev.Item1 +
                "\",\"type\":\"" + dev.Item2 +
                "\",\"description\":\"" + dev.Item2
                
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



    public  string Get_dev(int s) {
        //l1.Text = Request.QueryString["UPS"] ;
        return  Request.QueryString["UPS"] ;
    }

	public string Get_OneDev(){
		
		
		string tar = DevID.Text;
     
		
		return tar;
	}
    
	
	
	protected void Selection_Change(object sender, EventArgs e){
			ViewState["UPS"] = UPSList.SelectedItem.Value;
	}

    protected void BtnTest_Click(object sender, EventArgs e){
        Literal Msg = new Literal();
        Msg.Text = "<script>alert('TESDDDT');</script>";

        Page.Controls.Add(Msg);
    }
    
    public string Get_Devinfo(){
        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
        Conn.Open();
        string sql = "SELECT COUNT(*) FROM View_設備管理";

		SqlCommand cmd = new SqlCommand(sql,Conn);
        SqlDataReader dr = null;
        dr = cmd.ExecuteReader();
        int count = 0;
        if(dr.Read()){
            count = int.Parse(dr[0].ToString());
        }
        cmd.CommandText = @"SELECT 設備編號,
                                設備名稱,
                                設備種類,
                                設備用途,
                                放置地點,
                                用電電壓,
                                額定電流,
                                維護人員,
                                設備狀態,
                                備註說明
                                FROM [View_設備管理] ";
       
        dr.Close();
        dr = cmd.ExecuteReader();
        StringBuilder myStringBuilder = new StringBuilder("[");
        while(dr.Read())
        {          
            
            myStringBuilder.Append( "{\"ID\":\"" + dr[0].ToString()+
                "\",\"Name\":\"" +  dr[1].ToString() +
                "\",\"Class\":\"" +  dr[2].ToString() +
                //"\",\"Func\":\"" + dr[3].ToString() +
                "\",\"Place\":\"" + dr[4].ToString() +
                "\",\"Volt\":\"" + dr[5].ToString() +
                "\",\"Curre\":\"" + dr[6].ToString() +
                "\",\"Staff\":\"" + dr[7].ToString() +
                "\",\"State\":\"" + dr[8].ToString() );
                //"\",\"PS\":\"" + dr[9].ToString() );  
            
            if(count==1){
                myStringBuilder.Append( "\"}" );
                
            }
            else{
                myStringBuilder.Append( "\"}," );
            }
            count -= 1;        
        }        
        cmd.Cancel(); cmd.Dispose(); dr.Close(); Conn.Close(); Conn.Dispose();
        return myStringBuilder.ToString();

    }



}