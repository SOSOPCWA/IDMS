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
using System.IO;
public partial class Device_test : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {

        if (!IsPostBack)
        {
            ReadGroup();
            ViewState["Group"] = Request["Group"];
            if (Request["Group"] == null) ViewState["Group"] = "0";
        }
    }


    public string Get_Json()
    { //進入點
        string g1 = ViewState["Group"].ToString();
        if (g1 == "0") return Get_All_Json();
        else return Get_Selected_Json2(Int32.Parse(g1));//change here get2!!
    }


    public string Get_All_Json()
    {
        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
        Conn.Open();
        string sql = @"SELECT  d1.群組編號 gid,群組名稱 gname, d2.設備編號 did,設備名稱 dname
                        FROM [dbo].[群組關聯] d1, [dbo].實體設備 d2, [dbo].實體群組 d3
                        WHERE d1.群組編號=d3.群組編號 AND d2.設備編號=d1.設備編號
                        ORDER BY gid";
        SqlCommand cmd = new SqlCommand(sql, Conn);
        DataTable Devtbl = new DataTable();
        SqlDataAdapter dataAdapter = new SqlDataAdapter(cmd);//從Command取得資料存入dataAdapter
        dataAdapter.Fill(Devtbl);//將dataAdapter資料存入dataset
        int count_rows = Devtbl.Rows.Count;
        // SqlDataReader dr = null;
        // dr = cmd.ExecuteReader();


        StringBuilder All = new StringBuilder("{\"class\":\"go.GraphLinksMode\",");
        StringBuilder NodeArr = new StringBuilder("\"nodeDataArray\":[");
        StringBuilder LinkArr = new StringBuilder("\"linkDataArray\":[");
        //StringBuilder TestArr = new StringBuilder("");
        HashSet<int> Group_exist = new HashSet<int>();//已存在之群組
        HashSet<int> Dev_exist = new HashSet<int>();//已存在之群組



        //群組建立
        int loc1 = 50, loc2 = 50;

        //    NodeArr.Append(String.Format("{{\"key\":\"{0}\", \"loc\":\"{2} {3}\",\"text\":\"{1}\"}},",
        //                    "節點", "設備", loc1, loc2));

        //    NodeArr.Append(String.Format("{{\"key\":{1}, \"loc\":\"{2} {3}\",\"text\":\"{0}\"}},", "600室15RT", "17", loc1 + 200, loc2 + 300));
        //     NodeArr.Append(String.Format("{{\"key\":{1}, \"loc\":\"{2} {3}\",\"text\":\"{0}\"}},", "大同15RT氣冷式送風機(550室)", "4220", loc1 + 200, loc2 + 350));
        foreach (DataRow dr in Devtbl.Rows)
        {
            String groupid = dr["gid"].ToString();
            count_rows -= 1;
            if (Group_exist.Contains(Int32.Parse(groupid)))
            {//群組已建立
                loc2 += 50;
                NodeArr.Append(String.Format("{{\"key\":{2}, \"loc\":\"{3} {4}\",\"text\":\"{1}\", \"supers\":[\"{0}\"]}}",
                 groupid, dr["dname"].ToString(), dr["did"].ToString(), loc1, loc2));

            }
            else
            {
                loc1 += 300;
                NodeArr.Append(String.Format("{{\"key\":{0}, \"text\":\"{1}\", \"category\":\"Super\"}},",
                 groupid, dr["gname"].ToString()));
                NodeArr.Append(String.Format("{{\"key\":{2}, \"loc\":\"{3} {4}\",\"text\":\"{1}\", \"supers\":[\"{0}\"]}}",
                 groupid, dr["dname"].ToString(), dr["did"].ToString(), loc1, loc2));
                Group_exist.Add(Int32.Parse(groupid));//新增至已存在群組
                Dev_exist.Add(Int32.Parse(dr["did"].ToString()));
                Dev_exist.Add(Int32.Parse(groupid));

            }   
            if (count_rows != 0) { NodeArr.Append(","); }
        }


        //迴路建立

        sql = @"select d1.上游編號 uid,dt1.群組名稱 uname,d1.下游編號 did,dt2.群組名稱 dname
                from [dbo].空調 d1, 
                (select 群組編號,群組名稱
                from [dbo].實體群組 
                union
                select 設備編號,設備名稱
                from [dbo].實體設備 d2) dt1, 
                (select 群組編號,群組名稱
                from [dbo].實體群組 
                union
                select 設備編號,設備名稱
                from [dbo].實體設備 d2) dt2
                WHERE  (d1.上游編號=dt1.群組編號 ) AND (d1.下游編號=dt2.群組編號 )";
        cmd.CommandText = sql;
        SqlDataAdapter dataAdapter1 = new SqlDataAdapter(cmd);//從Command取得資料存入dataAdapter

        DataTable dt = new DataTable();
        dataAdapter1.Fill(dt);//將dataAdapter資料存入dataset
        count_rows = dt.Rows.Count;

        // LinkArr.Append(String.Format("{{\"from\":\"{0}\", \"to\":\"{1}\" }},",
        //          "節點", "10001"));


        foreach (DataRow dr1 in dt.Rows)
        {
            count_rows -= 1;
            if (!Dev_exist.Contains(Int32.Parse(dr1["did"].ToString())))
            {//下游的設備尚未建立節點
                f1(dr1["did"].ToString());
                loc2 += 50;
                NodeArr.Append(String.Format(",{{\"key\":{0}, \"loc\":\"{2} {3}\",\"text\":\"{1}\"}}",
                    dr1["did"].ToString(), dr1["dname"].ToString(), loc1, loc2));
            }

            if (count_rows == 0)
            {
                LinkArr.Append(String.Format("{{\"from\":{0}, \"to\":{1} }}",
                 dr1["uid"].ToString(), dr1["did"].ToString()));
            }
            else
            {
                LinkArr.Append(String.Format("{{\"from\":{0}, \"to\":{1} }},",
                 dr1["uid"].ToString(), dr1["did"].ToString()));
            }
        }




        cmd.Cancel(); cmd.Dispose(); Conn.Close(); Conn.Dispose();
        return All.Append(NodeArr + "],").Append(LinkArr).ToString();
        //return TestArr.ToString();
    }

    public string Get_Selected_Json(int first)
    {
        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
        Conn.Open();
        String sql = @"select d1.上游編號 uid,dt1.群組名稱 uname,d1.下游編號 did,dt2.群組名稱 dname
                    from [dbo].空調 d1, 
                    (select 群組編號,群組名稱
                    from [dbo].實體群組 
                    union
                    select 設備編號,設備名稱
                    from [dbo].實體設備 d2) dt1, 
                    (select 群組編號,群組名稱
                    from [dbo].實體群組 
                    union
                    select 設備編號,設備名稱
                    from [dbo].實體設備 d2) dt2
                    WHERE  (d1.上游編號=dt1.群組編號 ) AND (d1.下游編號=dt2.群組編號 )"; //讀入所有接冷關係
        SqlCommand cmd = new SqlCommand(sql, Conn);
        SqlDataAdapter dataAdapter1 = new SqlDataAdapter(cmd);//從Command取得資料存入dataAdapter

        DataTable dt = new DataTable();
        dataAdapter1.Fill(dt);//將dataAdapter資料存入dataset
        int count_rows = dt.Rows.Count;

        int loc1 = 50;//x
        int loc2 = 50;//y

        StringBuilder All = new StringBuilder("{\"class\":\"go.GraphLinksMode\",");
        StringBuilder NodeArr = new StringBuilder("\"nodeDataArray\":[");
        StringBuilder LinkArr = new StringBuilder("\"linkDataArray\":[{}");

        List<Tuple<int, string>> Group = new List<Tuple<int, string>>();//待建立底下設備的群組
        Stack<int> UpDown = new Stack<int>(); //待處理迴路設備

        //Group.Add(first); 
        UpDown.Push(first);
        bool first_t = true;
        while (UpDown.Any())
        {
            int tar = UpDown.Pop(); //pop目標出來

            foreach (DataRow dr1 in dt.Rows)
            { //針對每一個迴路 


                if (Int32.Parse(dr1["uid"].ToString()) == tar)
                { //上游相等 ==> 目標的下游
                    if (first_t)
                    {//放入第一項
                        Group.Add(Tuple.Create(Int32.Parse(dr1["uid"].ToString()), dr1["uname"].ToString()));
                        first_t = false;
                    }
                    UpDown.Push(Int32.Parse(dr1["did"].ToString())); //下游放入待處理名單
                    if (Group_or_Device(dr1["did"].ToString()) == "group")
                    { //與上游相接的是群組先放進待處理
                        Group.Add(Tuple.Create(Int32.Parse(dr1["did"].ToString()), dr1["dname"].ToString()));
                    }
                    else
                    { //與上游相接的是設備直接建立
                        NodeArr.Append(String.Format("{{\"key\":{0}, \"loc\":\"{2} {3}\",\"text\":\"{1}\"}},",
                                dr1["did"].ToString(), dr1["dname"].ToString(), loc1, -loc2));
                        loc2 += 50;
                        loc1 += 30;
                    }
                    //與上游的路徑
                    LinkArr.Append(String.Format(",{{\"from\":{0}, \"to\":{1} }}",
                            dr1["uid"].ToString(), dr1["did"].ToString()));
                }
            }
        }
        loc1 += 100;
        //群組名稱建立
        int len_G = Group.Count();
        foreach (var Gid in Group)
        {
            loc2 += 50;
            len_G -= 1;
            NodeArr.Append(String.Format("{{\"key\":{0}, \"text\":\"{1}\", \"category\":\"Super\"}}",
                    Gid.Item1.ToString(), Gid.Item2.ToString()));
            if (len_G != 0)
            {//補逗號                                 
                NodeArr.Append(",");
            }

        }
        //對待處理群組塞入底下的設備
        sql = @"SELECT  d1.群組編號 gid,群組名稱 gname, d2.設備編號 did,設備名稱 dname
                    FROM [dbo].[群組關聯] d1, [dbo].實體設備 d2, [dbo].實體群組 d3
                    WHERE d1.群組編號=d3.群組編號 AND d2.設備編號=d1.設備編號
                    ORDER BY gid ";
        cmd.CommandText = sql;
        DataTable Devtbl = new DataTable();
        SqlDataAdapter dataAdapter = new SqlDataAdapter(cmd);//從Command取得資料存入dataAdapter
        dataAdapter.Fill(Devtbl);//將dataAdapter資料存入dataset        

        bool Group_empty = !Group.Any(); //群組為空 (解決下游為空所以上面沒有建立群組的問題)
        first_t = true; //第一項新增群組
        //群組底下設備放入 Gid()
        Trace.Write("Group_empty?" + Group_empty.ToString());
        foreach (DataRow dr in Devtbl.Rows)
        {

            if (Group_empty)
            {      //對沒有其他供冷路徑的群組須建立群組跟設備          
                if (dr["gid"].ToString() == first.ToString())
                { //將選擇的群組 into JSON
                    if (first_t)
                    {
                        NodeArr.Append(String.Format("{{\"key\":{0}, \"text\":\"{1}\", \"category\":\"Super\"}}",
                                dr["gid"].ToString(), dr["gname"].ToString()));
                        loc2 += 50;
                        first_t = !first_t;
                    }
                    NodeArr.Append(String.Format(",{{\"key\":{2}, \"loc\":\"{3} {4}\",\"text\":\"{1}\", \"supers\":[\"{0}\"]}}",
                            first.ToString(), dr["dname"].ToString(), dr["did"].ToString(), loc1, -loc2));
                    loc2 += 50;
                }

            }
            else
            {//單一群組就建立底下設備
                for (int i = 0; i < Group.Count; i++)
                {

                    if (Group[i].Item1.ToString() == dr["gid"].ToString())
                    { //群組下的設備

                        NodeArr.Append(String.Format(",{{\"key\":{2}, \"loc\":\"{3} {4}\",\"text\":\"{1}\", \"supers\":[\"{0}\"]}}",
                                dr["gid"].ToString(), dr["dname"].ToString(), dr["did"].ToString(), (i * 200) + loc1, -loc2));
                        loc2 += 50;
                    }
                }
            }



        }
        cmd.Cancel(); cmd.Dispose(); Conn.Close(); Conn.Dispose();
        return All.Append(NodeArr + "],").Append(LinkArr).ToString();


    }

    public string Get_Selected_Json2(int first)
    {
        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
        Conn.Open();
        String sql = @"select d1.上游編號 uid,dt1.群組名稱 uname,d1.下游編號 did,dt2.群組名稱 dname
                    from [dbo].空調 d1, 
                    (select 群組編號,群組名稱
                    from [dbo].實體群組 
                    union
                    select 設備編號,設備名稱
                    from [dbo].實體設備 d2) dt1, 
                    (select 群組編號,群組名稱
                    from [dbo].實體群組 
                    union
                    select 設備編號,設備名稱
                    from [dbo].實體設備 d2) dt2
                    WHERE  (d1.上游編號=dt1.群組編號 ) AND (d1.下游編號=dt2.群組編號 )"; //讀入上下游關係(加上名稱)
        SqlCommand cmd = new SqlCommand(sql, Conn);
        SqlDataAdapter dataAdapter1 = new SqlDataAdapter(cmd); //從Command取得資料存入dataAdapter

        DataTable dt = new DataTable();
        dataAdapter1.Fill(dt); //將dataAdapter資料存入dataset
        int count_rows = dt.Rows.Count;

        int loc1 = 0; //x
        int loc2 = 0; //y

        StringBuilder All = new StringBuilder("{\"class\":\"go.GraphLinksMode\",");
        StringBuilder NodeArr = new StringBuilder("\"nodeDataArray\":[");
        StringBuilder LinkArr = new StringBuilder("\"linkDataArray\":[{}");
        HashSet<int> DeviceReplic = new HashSet<int>();
        List<Tuple<int, string, int, int>> Group = new List<Tuple<int, string, int, int>>(); //待建立底下設備的群組
        Stack<int> UpDown = new Stack<int>(); //待處理迴路設備

        UpDown.Push(first);
        bool first_t = true;
        while (UpDown.Any())
        {
            int tar = UpDown.Pop(); //pop目標出來

            foreach (DataRow dr1 in dt.Rows)
            { //針對每一個迴路 
                int DownID = Int32.Parse(dr1["did"].ToString());
                if (Int32.Parse(dr1["uid"].ToString()) == tar)
                { //上游相等 ==> 目標的下游
                    
                    if (first_t)
                    {//放入第一項
                        Group.Add(Tuple.Create(Int32.Parse(dr1["uid"].ToString()), dr1["uname"].ToString(), loc1, loc2));
                        loc1 += 150;
                        loc2 += 350;
                        first_t = false;
                    }
                    if(!DeviceReplic.Contains(DownID)){
                        UpDown.Push(Int32.Parse(dr1["did"].ToString())); //下游放入待處理名單
                        if (Group_or_Device(dr1["did"].ToString()) == "group")
                        { //與上游相接的是群組先放進待處理
                            Group.Add(Tuple.Create(Int32.Parse(dr1["did"].ToString()), dr1["dname"].ToString(), loc1, loc2));
                            loc1 += 150;
                            loc2 += 250;
                        }
                        else
                        { //與上游相接的是設備直接建立
                            
                            
                                NodeArr.Append(String.Format("{{\"key\":{0}, \"loc\":\"{2} {3}\",\"text\":\"{1}\"}},",
                                    dr1["did"].ToString(), dr1["dname"].ToString(), loc1, loc2));
                                DeviceReplic.Add(DownID);
                    
                            
                            loc1 += 150;
                            loc2 += 150;
                        }
                    }
                    
                    
                    //與上游的路徑
                    LinkArr.Append(String.Format(",{{\"from\":{0}, \"to\":{1} }}", dr1["uid"].ToString(), dr1["did"].ToString()));
                }
            }
        }

        //群組名稱建立
        int len_G = Group.Count();
        foreach (var Gid in Group)
        {
            len_G -= 1;
            NodeArr.Append(String.Format("{{\"key\":{0}, \"text\":\"{1}\", \"category\":\"Super\"}}",
                    Gid.Item1.ToString(), Gid.Item2.ToString(), Int32.Parse(Gid.Item3.ToString()), Int32.Parse(Gid.Item4.ToString())));
            if (len_G != 0)
            {//補逗號                                 
                NodeArr.Append(",");
            }
        }
        //對待處理群組塞入底下的設備
        sql = @"SELECT  d1.群組編號 gid,群組名稱 gname, d2.設備編號 did,設備名稱 dname
                    FROM [dbo].[群組關聯] d1, [dbo].實體設備 d2, [dbo].實體群組 d3
                    WHERE d1.群組編號=d3.群組編號 AND d2.設備編號=d1.設備編號
                    ORDER BY gid ";
        cmd.CommandText = sql;
        DataTable Devtbl = new DataTable();
        SqlDataAdapter dataAdapter = new SqlDataAdapter(cmd);//從Command取得資料存入dataAdapter
        dataAdapter.Fill(Devtbl);//將dataAdapter資料存入dataset        

        bool Group_empty = !Group.Any(); //群組為空 (解決下游為空所以上面沒有建立群組的問題)
        first_t = true; //第一項新增群組

        int[] j = new int[Group.Count] ;
        for ( int i = 0; i < Group.Count; i++ ) j[i] = 0;
        foreach (DataRow dr in Devtbl.Rows)  //群組底下設備放入 Gid()
        {
            if (Group_empty)  //對沒有其他供冷路徑的群組須建立群組跟設備 
            {              
                if (dr["gid"].ToString() == first.ToString())
                { //將選擇的群組 into JSON
                    if (first_t)
                    {
                        NodeArr.Append(String.Format("{{\"key\":{0}, \"text\":\"{1}\", \"category\":\"Super\"}}",
                                dr["gid"].ToString(), dr["gname"].ToString()));
                        loc2 += 50;
                        first_t = !first_t;
                    }
                    NodeArr.Append(String.Format(",{{\"key\":{2}, \"loc\":\"{3} {4}\",\"text\":\"{1}\", \"supers\":[\"{0}\"]}}",
                            first.ToString(), dr["dname"].ToString(), dr["did"].ToString(), loc1, loc2));
                    loc2 -= 50;
                }
            }
            else //單一群組就建立底下設備
            {
                for (int i = 0; i < Group.Count; i++)
                {
                    
                    if (dr["gid"].ToString() == Group[i].Item1.ToString()) //設備屬於第i個群組
                    { 
                        NodeArr.Append(String.Format(",{{\"key\":{2}, \"loc\":\"{3} {4}\",\"text\":\"{1}\", \"supers\":[\"{0}\"]}}",
                                dr["gid"].ToString(), dr["dname"].ToString(), dr["did"].ToString(), 
                                Int32.Parse(Group[i].Item3.ToString()), Int32.Parse(Group[i].Item4.ToString()) + j[i]));
                        j[i] += 50;
                    }
                }
            }
        }
        cmd.Cancel(); cmd.Dispose(); Conn.Close(); Conn.Dispose();
        return All.Append(NodeArr + "],").Append(LinkArr).ToString();
    }

    public static void WriteJson(string para){
        try
        {
            StreamWriter sw = new StreamWriter("d:\\ssm_dev\\SSM_DEV2\\IDMS\\Device\\01.txt");
            sw.Write(para);
            sw.Close();            
        }
        catch (Exception e)
        {
            // Let the user know what went wrong.
            //Trace.Write(e.ToString());
            //Trace.Warn(e.ToString());
        }
        
    
    }

    [System.Web.Services.WebMethod]
    public static string GetData(string parameter1, string parameter2)
    {
        WriteJson(parameter1);
        // Perform some server-side logic here and return a response
        return "Response from the server";
    }

    

    protected string Group_or_Device(string sin)
    {
        if (sin.Length >= 5)
        {
            return "group";
        }
        else
        {
            return "device";
        }
    }
    protected void ReadGroup()
    {
        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
        Conn.Open();
        string sql = @"SELECT [群組編號]
                        ,[群組名稱]      
                        ,[群組說明]
                        FROM [dbo].[實體群組]";
        SqlCommand cmd = new SqlCommand(sql, Conn);
        DataTable Devtbl = new DataTable();
        SqlDataAdapter dataAdapter = new SqlDataAdapter(cmd);//從Command取得資料存入dataAdapter
        dataAdapter.Fill(Devtbl);//將dataAdapter資料存入dataset
        //int count_rows = Devtbl.Rows.Count;
        int head = 0;

        if (Devtbl.Rows.Count > 0)
        {
            foreach (DataRow row in Devtbl.Rows)
            {
                GroupList.Items.Insert(head, new ListItem(row[0].ToString() + " -- " + row[1].ToString(),
                                                         row[0].ToString()));
                head += 1;
            }
        }

        cmd.Cancel(); cmd.Dispose(); Conn.Close(); Conn.Dispose();
    }

    protected void Selection_Change(object sender, EventArgs e)
    {
        ViewState["Group"] = GroupList.SelectedItem.Value;

    }

    protected void Linkbtn_click(object sender, EventArgs e){
        string GID = GroupList.SelectedItem.Value;
        Literal Msg = new Literal();
            
            Msg.Text = String.Format("<script>window.open('Group.aspx?GID={0}', '_blank')</script>", GID);
            Page.Controls.Add(Msg);

    }

    protected void f1(string intd)
    {
        Testa.Text = intd;
    }



}