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

public partial class GroupEdit : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {

        Trace.Write("post back");
        //如果有GET參數就放入session，沒有就0
        Session["Group"] = Request.QueryString["GID"]==null?"0":Request.QueryString["GID"];       
        
        
        if (!IsPostBack)
        {
            Trace.Write("no post back");
            //讓選單重新讀取清單列表 
            ReadGroup();
            
            int GID = Int32.Parse(Session["Group"].ToString());
            GroupList.SelectedIndex = GID%10000;
            GroupList_SelectedIndexChanged(this, EventArgs.Empty);

        }
    }
    protected void Page_PreRenderComplete(object sender, EventArgs e)   //放Page_Load無法顯示
    {
        Trace.Write("prerender!!!!");
        ExecList(); 

    }

    protected string Get_list()
    { //產生設備列表
        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
        Conn.Open();
        string sql = @"select COUNT(*) FROM [dbo].[View_設備管理]";
        SqlCommand cmd = new SqlCommand(sql, Conn);
        SqlDataReader dr = null;
        dr = cmd.ExecuteReader();
        int count = 0;
        if (dr.Read())
        {
            count = int.Parse(dr[0].ToString());
        }
        cmd.CommandText = @"select 設備編號,設備名稱,資產編號 FROM [dbo].[View_設備管理]";
        dr.Close();
        dr = cmd.ExecuteReader();

        StringBuilder myStringBuilder = new StringBuilder("[");

        while (dr.Read() & count > 0)
        {

            myStringBuilder.Append("{\"ID\":\"" + dr[0].ToString() +

                "\",\"Name\":\"" + dr[1].ToString() +
                "\",\"prop\":\"" + dr[2].ToString()
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
        myStringBuilder.Replace("	", "SSW");
        cmd.Cancel(); cmd.Dispose(); dr.Close(); Conn.Close(); Conn.Dispose();
        return myStringBuilder.ToString();
    }

    protected void GroupList_SelectedIndexChanged(object sender, EventArgs e)   //顯示資產描述
    {
        Trace.Write ( "In SELECTED!!!!");
        //TEST.Text = GroupList.SelectedValue;
        toplink.Items.Clear();downlink.Items.Clear();        
        String gid = GroupList.SelectedValue.Split(new string[]{"--"}, StringSplitOptions.None)[0];
        String sqlQuery = @"SELECT g.[上游編號] UID,gt.群組名稱 UNAME, g.下游編號 DID, gt2.群組名稱 DNAME
                                FROM [dbo].[空調] g,
                                ( Select gp.群組編號,gp.群組名稱
                                FROM [dbo].[實體群組] gp
                                UNION
                                SELECT d.設備編號,d.設備名稱
                                FROM [dbo].[View_設備管理] d
                                ) gt,
                                ( SELECT gp.群組編號,gp.群組名稱
                                FROM [dbo].[實體群組] gp
                                UNION
                                SELECT d.設備編號,d.設備名稱
                                FROM [dbo].[View_設備管理] d
                                ) gt2
                                WHERE g.上游編號=gt.群組編號 and g.下游編號=gt2.群組編號";
                              
        DataSet ds = RunQuery(sqlQuery + " AND (g.[下游編號]=@ID )" , "ID", gid);
        foreach (DataRow row in ds.Tables[0].Rows){
            toplink.Items.Add(new ListItem(String.Format("<{0}>  {1}", row["UID"].ToString(), row["UNAME"].ToString()),row["UID"].ToString()));                    
        }
        ds = RunQuery(sqlQuery + " AND (g.[上游編號]=@ID )" , "ID", gid);
        foreach (DataRow row in ds.Tables[0].Rows){
            downlink.Items.Add(new ListItem(String.Format("<{0}>  {1}", row["DID"].ToString(), row["DNAME"].ToString()),row["DID"].ToString()));                    
        }

        Session["Group"] = gid;
         ds.Dispose(); ds.Dispose();
    }

    protected void ExecList(){
        Trace.Write("EXEC LIST!");
        String GID = "0";
        try
        {
            String[] glist = GroupList.SelectedValue.Split(new string[]{"--"}, StringSplitOptions.None); //gid, gname, guse, gps, gloc
            add_gname.Text = glist[1];
            add_guse.SelectedValue = glist[2];
            add_gps.Text = glist[3];
            add_gloc.Text = glist[4];
            GID = glist[0];
            //Session["Group"] = glist[0];
            Btn_gedit.Visible = true;
            Btn_gdelete.Visible = true;
            linkmodal.Visible = true;
            groupmodal.Text = "群組編輯";
        }
        catch (System.Exception)
        {
            add_gname.Text = "";
            add_guse.SelectedValue = "";
            add_gps.Text = "";
            add_gloc.Text = "";
            //Session["Group"] = 0;
            Btn_gedit.Visible = false;
            Btn_gdelete.Visible = false;
            linkmodal.Visible = false;
            groupmodal.Text = "新增群組";
            Trace.Write("List create Fail");
        }
        
        SqlDataSource1.SelectCommand = @"SELECT g.[設備編號] ID, d.設備名稱 DName, d.放置地點 Loc
                            FROM [dbo].[群組關聯] g, [dbo].[View_設備管理] d
                            WHERE g.設備編號 = d.設備編號 AND  群組編號 = " + Session["Group"].ToString();
        Trace.Warn(SqlDataSource1.SelectCommand.ToString());
    
        
        
    }


    protected void Row_Command(object sender, GridViewCommandEventArgs e){

        try{
            int index = Convert.ToInt32(e.CommandArgument);
                GridViewRow row = GridView1.Rows[index];
                
            if(e.CommandName == "Edit1"){
                Literal Msg = new Literal();
                
                Msg.Text = String.Format("<script>window.open('DevEdit.aspx?DevNo={0}', '_blank')</script>", row.Cells[0].Text);
                Page.Controls.Add(Msg);
                Session["Group"] = GroupList.SelectedValue.Split(new string[]{"--"}, StringSplitOptions.None)[0];//避免按了編輯就跑掉
            }else{
                Literal Msg = new Literal();                
                BtnDelete_Click(GroupList.SelectedValue.Split(new string[]{"--"}, StringSplitOptions.None)[0], row.Cells[0].Text);
            }
        }catch{
            Literal Msg = new Literal();
            
            Msg.Text = String.Format("<script>alert({0})</script>", sender);
            Page.Controls.Add(Msg);
        }
    }

    protected Button Button_Create(string ID){
        Trace.Write("This is create button");
        Button d = new Button();
        d.Text = "刪除";
        d.Attributes["class"] = "btndelete btn btn-danger";
        

        //d.Name("刪除"); 
        
        d.Click += (BtnDelete_Click);
        /*
        d.Command += (s,e) => {
            Trace.Write ( "In Delete2");
            try{
                Trace.Write ( "In Try");
                throw new NullReferenceException("Student object is null.");
                List<SqlParameter> pars = new List<SqlParameter>();
                
                ExecDbSQL(@"DELETE FROM [dbo].[群組關聯]
                        WHERE 群組編號 = @gid AND 設備編號 = @did", pars);
                TEST.Text = "DWDWDWD!!";
                Literal Msg = new Literal();
                Msg.Text = "<script>alert('新增成功!');</script>";
                Page.Controls.Add(Msg);  
                Trace.Warn ( "DONE");
            }catch(System.Exception)
            {
                Trace.Warn("Fail");
                TEST.Text = "UYUYUYUY!!!";
                Literal Msg = new Literal();
                Msg.Text = "<script>alert('失敗');</script>";
                Page.Controls.Add(Msg);
            }
        };
        */
        d.ID = ID;
        d.OnClientClick += "return confirm('確定刪除此設備嗎?');";
        return d;    
    }
    protected string GetValueWithItemName(string SQL, string key, string value)
    {
        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
        Conn.Open();
        SqlCommand cmd = new SqlCommand(SQL, Conn);
        cmd.Parameters.Add(key, SqlDbType.VarChar).Value = value;
        SqlDataReader dr = null;
        dr = cmd.ExecuteReader();
        string cfg = ""; if (dr.Read()) cfg = HttpUtility.HtmlEncode(dr[0].ToString());

        cmd.Cancel(); cmd.Dispose(); dr.Close(); Conn.Close(); Conn.Dispose();

        return (cfg);
    } 
    protected void ReadGroup() //群組下拉式選單
    {
        SqlCommand sqlQuery = new SqlCommand();
        sqlQuery.CommandText = "SELECT * FROM [dbo].實體群組";
        DataSet ds = RunQuery(sqlQuery);
        GroupList.Items.Clear();
        GroupList.Items.Add(new ListItem("(請選擇)", "0"));
        if (ds.Tables.Count > 0)
        {
            foreach (DataRow row in ds.Tables[0].Rows)
            {
                GroupList.Items.Add(new ListItem(String.Format("{0} ({1}) -- {2}", row["群組編號"], row["群組用途"], row["群組名稱"])
                                                , String.Format("{4}--{0}--{1}--{2}--{3}", row["群組名稱"], row["群組用途"], row["群組說明"], row["擺放地點"], row["群組編號"])));
            }
        }
        sqlQuery.Cancel(); ds.Dispose();
    }
    protected DataSet RunQuery(SqlCommand sqlQuery) //讀取DB資訊
    {
        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
        SqlDataAdapter dbAdapter = new SqlDataAdapter();
        dbAdapter.SelectCommand = sqlQuery;
        sqlQuery.Connection = Conn;
        DataSet QueryDataSet = new DataSet();
        dbAdapter.Fill(QueryDataSet);
        dbAdapter.Dispose(); Conn.Close(); Conn.Dispose();
        return (QueryDataSet);
    }
    protected DataSet RunQuery(String sqlQuerys, string key, string value) //讀取DB資訊
    {
        SqlCommand sqlQuery = new SqlCommand(sqlQuerys);
        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
        SqlDataAdapter dbAdapter = new SqlDataAdapter();
        dbAdapter.SelectCommand = sqlQuery;
        sqlQuery.Parameters.Add(key, value);
        sqlQuery.Connection = Conn;
        DataSet QueryDataSet = new DataSet();
        dbAdapter.Fill(QueryDataSet);
        dbAdapter.Dispose(); Conn.Close(); Conn.Dispose();
        return (QueryDataSet);
    }

    protected void BtnAddGroup_Click(object sender, EventArgs e){
        string gname = add_gname.Text;
        string guse = add_guse.SelectedValue;
        string gps = add_gps.Text;
        string gloc = add_gloc.Text;
        string Group = GetPKNo("群組編號", "實體群組").ToString();
        //TEST.Text = GroupList.SelectedValue;
        try{
            if(gname=="" || gps=="" || gloc=="")throw new Exception("empty fail");
            List<SqlParameter> pars = new List<SqlParameter>();
            pars.Add(new SqlParameter("@group", Group));
            pars.Add(new SqlParameter("@gname", gname));
            pars.Add(new SqlParameter("@guse", guse));
            pars.Add(new SqlParameter("@gps", gps));
            pars.Add(new SqlParameter("@gloc", gloc));

            if(!RightCheck(Session["UserName"].ToString()) ){
                throw new FormatException();
            }

            ExecDbSQL(@"INSERT INTO [dbo].[實體群組]
                        VALUES (@group, @gname, @guse, @gps, @gloc)", pars);
            
            Literal Msg = new Literal();
            Msg.Text = "<script>alert('新增成功!');</script>";
            Page.Controls.Add(Msg);
            
            Session["Group"] = Group;
            ReadGroup();
            GroupList.SelectedIndex = Int32.Parse(Group)%10000;
            Trace.Write(Session["Group"].ToString());
        }
        catch (FormatException se)
        {            
            Trace.Write(se.ToString());        
            Literal Msg = new Literal();
            Msg.Text = String.Format("<script>alert('{0}');</script>", "您不具有權限異動群組!\\n若有使用需求請聯絡作業管控科");
            Page.Controls.Add(Msg);
        }        
        catch (System.Exception se)
        {            
            Trace.Write(se.ToString());
            Literal Msg = new Literal();
            Msg.Text = "<script>alert('新增群組資料失敗!\\n請檢查輸入資料是否完整!');</script>";
            Page.Controls.Add(Msg);
        }
    }

    protected void BtnEditGroup_Click(object sender, EventArgs e){
        string gname = add_gname.Text;
        string guse = add_guse.SelectedValue;
        string gps = add_gps.Text;
        string gloc = add_gloc.Text;
        string gid = GroupList.SelectedValue.Split(new string[]{"--"}, StringSplitOptions.None)[0];
        int index = GroupList.SelectedIndex;
        //TEST.Text = GroupList.SelectedValue;
        try{
            if(!RightCheck(Session["UserName"].ToString()) ){
                throw new FormatException();
            }   
            List<SqlParameter> pars = new List<SqlParameter>();
            pars.Add(new SqlParameter("@gid", gid));
            pars.Add(new SqlParameter("@gname", gname));
            pars.Add(new SqlParameter("@guse", guse));
            pars.Add(new SqlParameter("@gps", gps));
            pars.Add(new SqlParameter("@gloc", gloc));

            ExecDbSQL(@"UPDATE [dbo].[實體群組]
                        SET 群組名稱=@gname, 群組用途=@guse, 群組說明=@gps, 擺放地點=@gloc
                        WHERE 群組編號=@gid", pars);
            
            Literal Msg = new Literal();
            Msg.Text = "<script>alert('修改成功!');</script>";
            Page.Controls.Add(Msg);

            ReadGroup();
            Trace.Write(index.ToString());
            GroupList.SelectedIndex = index;
            Session["Group"] = GroupList.SelectedValue.Split(new string[]{"--"}, StringSplitOptions.None)[0];
        }
        catch (FormatException se)
        {            
            Trace.Write(se.ToString());        
            Literal Msg = new Literal();
            Msg.Text = String.Format("<script>alert('{0}');</script>", "您不具有權限異動群組!\\n若有使用需求請聯絡作業管控科");
            Page.Controls.Add(Msg);
        }
        catch (System.Exception)
        {            
            Literal Msg = new Literal();
            Msg.Text = "<script>alert('修改群組資料失敗!\n請再檢查輸入資料!');</script>";
            Page.Controls.Add(Msg);
        }
    }

    protected void BtnDeleteGroup_Click(object sender, EventArgs e){
        string gid = GroupList.SelectedValue.Split(new string[]{"--"}, StringSplitOptions.None)[0];


        Trace.Write(gid);
        //TEST.Text = GroupList.SelectedValue;
        try{
            if(!RightCheck(Session["UserName"].ToString()) ){
                throw new FormatException();
            }
            List<SqlParameter> pars = new List<SqlParameter>();
            pars.Add(new SqlParameter("@gid", gid));


            ExecDbSQL(@"DELETE FROM [dbo].[實體群組]
                        WHERE 群組編號 = @gid;
                        DELETE FROM [dbo].[群組關聯]
                        WHERE 群組編號 = @gid", pars);
            
            Literal Msg = new Literal();
            Msg.Text = "<script>alert('刪除群組成功!');</script>";
            Page.Controls.Add(Msg);
            ReadGroup();
            Session["Group"] = 0;
            Trace.Write(" delete group done");
            Trace.Write(Session["Group"].ToString());
        }
        catch (FormatException se)
        {            
            Trace.Write(se.ToString());        
            Literal Msg = new Literal();
            Msg.Text = String.Format("<script>alert('{0}');</script>", "您不具有權限異動群組!\\n若有使用需求請聯絡作業管控科");
            Page.Controls.Add(Msg);
        }
        
        catch (System.Exception)
        {            
            Trace.Write("delete group fail");
            Literal Msg = new Literal();
            Msg.Text = "<script>alert('刪除群組資料失敗!\n請再檢查輸入資料!');</script>";
            Page.Controls.Add(Msg);
        }
    }

    protected void BtnAdd_Click(object sender, EventArgs e){ //設備加至群組
        Trace.Write("Write dev into Group");
        
        string DevID = Request.Form["search"]; //送出之設備完整名稱
        string[] devna = DevID.Split(new string[]{" -- "}, StringSplitOptions.None); //(ID), name
        string Group = GroupList.SelectedValue.Split(new string[]{"--"}, StringSplitOptions.None)[0]; //群組編號
       
        //TEST.Text = GroupList.SelectedValue;
        try{
            Trace.Write("go add");
            Trace.Write(RightCheck(Session["UserName"].ToString()).ToString());
            if(!RightCheck(Session["UserName"].ToString()) ){
                throw new FormatException();
            }
            DevID = devna[0].Substring(1,devna[0].Length-2);           
            if(devna.Length!=2)throw new ArgumentException("輸入錯誤");
            List<SqlParameter> pars = new List<SqlParameter>();
            pars.Add(new SqlParameter("@group", Group));
            pars.Add(new SqlParameter("@device", DevID));
            ExecDbSQL(@"INSERT INTO [dbo].[群組關聯]
                        VALUES (@group, @device)", pars);
            Literal Msg = new Literal();
            Msg.Text = "<script>alert('新增至群組成功!');</script>";
            Page.Controls.Add(Msg);
            Session["Group"] = Group;
            
        }
        
        
        catch (FormatException se)
        {            
            Trace.Write(se.ToString());        
            Literal Msg = new Literal();
            Msg.Text = String.Format("<script>alert('{0}');</script>", "您不具有權限異動群組!\\n若有使用需求請聯絡作業管控科");
            Page.Controls.Add(Msg);
        }
        catch (System.Exception se)
        {            
            Trace.Write(se.ToString());        
            Literal Msg = new Literal();
            Msg.Text = "<script>alert('名稱錯誤或資料已存在\\n請再檢查輸入資料!');</script>";
            Page.Controls.Add(Msg);
        }
    }

    protected void BtnDelete_Click(string gid, string did){
        Trace.Write ( "In Delete");
        try{
            Trace.Write ( "In Try");
            if(!RightCheck(Session["UserName"].ToString()) ){
                    throw new FormatException();
                }   
            //throw new NullReferenceException("Student object is null.");
            List<SqlParameter> pars = new List<SqlParameter>();
            pars.Add(new SqlParameter("gid", gid));
            pars.Add(new SqlParameter("did", did));
            Trace.Write ( "In SQL");
            ExecDbSQL(@"DELETE FROM [dbo].[群組關聯]
                    WHERE 群組編號 = @gid AND 設備編號 = @did", pars);
            Trace.Write ( "In END");

            Literal Msg = new Literal();
            Msg.Text = "<script>alert('刪除成功!');</script>";
            Page.Controls.Add(Msg);  
            Trace.Warn ( "DONE");
            Session["Group"] = gid;
        }
        catch (FormatException se)
        {            
            Trace.Write(se.ToString());        
            Literal Msg = new Literal();
            Msg.Text = String.Format("<script>alert('{0}');</script>", "您不具有權限異動群組!\\n若有使用需求請聯絡作業管控科");
            Page.Controls.Add(Msg);
        }
        
        catch(System.Exception e)
        {
            Trace.Warn ( "Fail");
            Trace.Warn (e.ToString());
     
            Literal Msg = new Literal();
            Msg.Text = "<script>alert('刪除設備資料失敗');</script>";
            Page.Controls.Add(Msg);
        }
    }

    protected void BtnDelete_Click(object sender, EventArgs e){
        Trace.Write ( "In Delete");
        try{
            Trace.Write ( "In Try");
            throw new NullReferenceException("Student object is null.");
            List<SqlParameter> pars = new List<SqlParameter>();
            
            ExecDbSQL(@"DELETE FROM [dbo].[群組關聯]
                    WHERE 群組編號 = @gid AND 設備編號 = @did", pars);
            
            Literal Msg = new Literal();
            Msg.Text = "<script>alert('新增成功!');</script>";
            Page.Controls.Add(Msg);  
            Trace.Warn ( "DONE");
            
        }catch(System.Exception)
        {
            Trace.Warn("Fail");

            Literal Msg = new Literal();
            Msg.Text = "<script>alert('刪除設備資料失敗');</script>";
            Page.Controls.Add(Msg);
        }
    }    

    protected int GetPKNo(string PKfield, string PKtbl) //取得主鍵編號
    {
        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
        Conn.Open();
        SqlCommand cmd = new SqlCommand("select ISNULL(MAX(" + PKfield + "), 0) FROM " + PKtbl, Conn);
        SqlDataReader dr = null;
        dr = cmd.ExecuteReader();
        int PkNo = 1; 
        if (dr.Read()) PkNo=int.Parse(dr[0].ToString()) + 1;
        cmd.Cancel(); cmd.Dispose(); dr.Close(); Conn.Close(); Conn.Dispose();
        return (PkNo);
    }


    protected void InsLifeLog(string SQL) //寫入生命履歷
    {
        string LifeNo = GetPKNo("履歷編號", "生命履歷").ToString(); //履歷編號
        string TblName = "實體設備";    //表格名稱
                                    //string PKno = TBid1.Text;   //主鍵編號
                                    // string OldMt = GetValueWithItemName("select [維護人員] from [實體設備] where [設備編號]=@TextDevNo","@TextDevNo",TBid1.Text);
                                    // string OldKeeper = GetValueWithItemName("select [保管人員] from [View_設備管理] where [設備編號]=@TextDevNo", "@TextDevNo", TBid1.Text);     //原保管人
        string UN = Session["UserName"].ToString();   //登入的UserName：異動人員
        string LiftDT = DateTime.Now.ToString("yyyy/MM/dd HH:mm");  //異動日期
        string LifeIP = Request.ServerVariables["REMOTE_ADDR"].ToString();
        List<SqlParameter> pars = new List<SqlParameter>();
        pars.Add(new SqlParameter("@LifeNo", LifeNo));
        pars.Add(new SqlParameter("@TblName", TblName));
        //pars.Add(new SqlParameter("@PKno", PKno));
        pars.Add(new SqlParameter("@SQL", SQL.Replace("'", "''")));
        ////pars.Add(new SqlParameter("@Mt", OldMt));	
        //   pars.Add(new SqlParameter("@OldMt", OldMt));
        //   pars.Add(new SqlParameter("@OldKeeper", OldKeeper));
        pars.Add(new SqlParameter("@UN", UN));
        pars.Add(new SqlParameter("@LiftDT", LiftDT));
        pars.Add(new SqlParameter("@LifeIP", LifeIP));
        ExecDbSQL("Insert INTO [生命履歷] VALUES( @LifeNo , @TblName, @PKno, @SQL, @Mt, @OldMt, @OldKeeper, @UN, @LiftDT, @LifeIP)", pars);
    }
   
    protected void ExecDbSQL(string SQL, List<SqlParameter> pars) //執行資料庫異動
    {
        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
        Conn.Open();
        SqlCommand cmd = new SqlCommand(SQL, Conn);
        cmd.Parameters.AddRange(pars.ToArray());
        cmd.ExecuteNonQuery();
        cmd.Cancel(); cmd.Dispose(); Conn.Close(); Conn.Dispose();
    }

    protected bool RightCheck(string username){ //群組關聯目前只允許管理科成員使用        
        bool TF = false;
        using (SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString))
        {
            SqlCommand cmd =
                new SqlCommand(@"SELECT * FROM [config] WHERE Kind='作業管控科' AND Item = @uname", Conn);
            Conn.Open();
            cmd.Parameters.AddWithValue("@uname", username);
            SqlDataReader dr = cmd.ExecuteReader();
            if(dr.Read())TF = true;        
            dr.Close();
        }
        return TF;
    }





}
