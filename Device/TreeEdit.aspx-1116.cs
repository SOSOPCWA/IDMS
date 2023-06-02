using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.Configuration;
using System.Data;
using System.Data.SqlClient;

public partial class Device_TreeEdit : System.Web.UI.Page
{   
    protected void Page_PreRenderComplete(object sender, EventArgs e)   //��Page_Load�L�k���
    {   
        if (!IsPostBack)
        {
            if (Request["DevNo"] != null) TextDevNo.Text = Request["DevNo"];

            if (TextDevNo.Text == "" | TextDevNo.Text == "0") AddMsg("<script>alert('�z�|�����w�]�ơI �Х������s�W�]�ƫ�A�]�w�]�ưj����T�I');window.close();</script>");
            else
            {
                ReadDevice(TextDevNo.Text);   //Ū������]�Ƹ�T        
                ReadConn(TextDevNo.Text, "���q", "Up", ListUpPower);  //Ū���W�屵�q�j�����
                ReadConn(TextDevNo.Text, "���q", "Dn", ListDnPower); //Ū���U�屵�q�j�����
                ReadConn(TextDevNo.Text, "����", "Up", ListUpNet); //Ū���W�屵���j����� 
                ReadConn(TextDevNo.Text, "����", "Dn", ListDnNet); //Ū���U�屵���j����� 
            }

            ReadHelp();

            Place_Selected(ListSource, SelPlace, SelPointer, SqlDataSourcePointer, SqlDataSourceSource);
        }
    }

    protected void ReadDevice(string DevNo)    //Ū������]��
    {
        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
        Conn.Open();
        SqlCommand cmd = new SqlCommand("SELECT * FROM [View_�]�ƺ޲z] where [�]�ƽs��]=" + DevNo, Conn);
        SqlDataReader dr = null;
        dr = cmd.ExecuteReader();

        if (dr.Read())
        {
            lblDevName.Text = dr["�]�ƦW��"].ToString();
            lblDevice.Text =  dr["�]�ƥγ~"].ToString() + " " + dr["��m�a�I"].ToString() + " " + dr["����"].ToString() + "(U)";
            lblMt.Text = dr["���@�H��"].ToString();
            lblTree.Text = GetPath(dr["�]�ƽs��"].ToString(), "���q", "") + "��<br />" + GetPath(dr["�]�ƽs��"].ToString(), "����", "") + "��";
        }

        cmd.Cancel(); cmd.Dispose(); dr.Close(); Conn.Close(); Conn.Dispose();
    }

    protected void ReadConn(string DevNo, string tbl, string UpDn, ListBox ListName)    //Ū���W�U��]�ưj�����
    {
        SqlCommand sqlQuery = new SqlCommand();
        string UpDn1 = "�W", UpDn2 = "�U";
        if (UpDn == "Dn")
        {
            UpDn1 = "�U"; UpDn2 = "�W";
        }
        sqlQuery.CommandText = "SELECT [�]�ƽs��],[�]�ƦW��] FROM [����]��],[" + tbl + "]"
            + " WHERE [����]��].[�]�ƽs��]=[" + tbl + "].[" + UpDn1 + "��s��]"
            + " and [" + tbl + "].[" + UpDn2 + "��s��] = " + DevNo + " ORDER BY [�]�ƦW��]";
        DataSet ds = RunQuery(sqlQuery);

        ListName.Items.Clear();

        if (ds.Tables.Count > 0)
        {
            foreach (DataRow row in ds.Tables[0].Rows)
            {
                ListName.Items.Add(new ListItem(row[1].ToString(), row[0].ToString()));
            }
        }

        sqlQuery.Cancel(); ds.Dispose();
    }

    protected void ReadHelp() //Ū����컡��
    {
        SqlCommand sqlQuery = new SqlCommand();
        sqlQuery.CommandText = "select [Item],[Memo],[Config] from [Config] where [Kind]='�]�ưj��' order by [Mark]";
        DataSet ds = RunQuery(sqlQuery);

        if (ds.Tables.Count > 0)
        {
            foreach (DataRow row in ds.Tables[0].Rows)
            {
                Label helpObj = (Label)form1.FindControl("help" + row[2].ToString()); //������r
                helpObj.Text = row[1].ToString().Replace("\r\n", "<br />");
            }
        }

        sqlQuery.Cancel(); ds.Dispose();
    }

    //�s�W�j��----------------------------------------------------------------
    protected Boolean BeenConned(string No, ListBox ListName) //�O�_�����ƪ��]�ưj��
    {
        Boolean TF = false;
        for (int i = 0; i < ListName.Items.Count; i++)
        {
            if (ListName.Items[i].Value == No) TF = true;
        }

        return (TF);
    }

    protected void LinkUpPowerAdd_Click(object sender, EventArgs e)
    {
        LinkAdd_Click(ListSource, ListUpPower);
    }
    protected void LinkDnPowerAdd_Click(object sender, EventArgs e)
    {
        LinkAdd_Click(ListSource, ListDnPower);
    }
    protected void LinkUpNetAdd_Click(object sender, EventArgs e)
    {
        LinkAdd_Click(ListSource, ListUpNet);
    }
    protected void LinkDnNetAdd_Click(object sender, EventArgs e)
    {
        LinkAdd_Click(ListSource, ListDnNet);
    }

    protected void LinkAdd_Click(ListBox ListSource, ListBox ListTarget)
    {
        string tbl = "", UpDn = "", UpDnC = "", UpNo = "", DnNo = "",TreeNos="";
        Boolean TF = false;
        
        if (ListSource.SelectedValue == "") AddMsg("<script>alert('�Х�����z�n�s�W���j���I');</script>");
        else if (ListSource.SelectedValue == TextDevNo.Text) AddMsg("<script>alert('�z����s�W�ۤv���]�ưj���I');</script>");
        else if (BeenConned(ListSource.SelectedValue, ListTarget)) AddMsg("<script>alert('�Ӱj���w�s�b�A�z�����A�s�W�I');</script>");
        else
        {
            if (RightCheck())
            {
                TF = true;

                if (ListTarget.ID == "ListUpPower" | ListTarget.ID == "ListDnPower")
                {
                    tbl = "���q";
                    if (ListTarget.ID == "ListUpPower")
                    {
                        UpDn = "Up"; UpDnC = "�W��";
                        UpNo = ListSource.SelectedValue;
                        DnNo = TextDevNo.Text;                        
                    }
                    else
                    {
                        UpDn = "Dn"; UpDnC = "�U��";
                        UpNo = TextDevNo.Text;
                        DnNo = ListSource.SelectedValue;
                    }

                    if (GetValue("select count(*) from [���q] where [�U��s��]=" + UpNo) == "0")  //�@���W�大�j���L�q�O�W��I
                    {
                    //    TF = false;
                        AddMsg("<script>alert('�@���W�大�j��(" + ListSource.SelectedItem.Text + ")�L�q�O�W��I�Y���ݭn�A�зs�W��(�`�q��)���U�C');</script>");
                    }                                        
                }
                else
                {
                    tbl = "����";
                    if (ListTarget.ID == "ListUpNet")
                    {
                        UpDn = "Up"; UpDnC = "�W��";
                        UpNo = ListSource.SelectedValue;
                        DnNo = TextDevNo.Text;

                        if (!NetCheck(UpNo)) TF = false;
                    }
                    else
                    {
                        UpDn = "Dn"; UpDnC = "�U��";
                        UpNo = TextDevNo.Text;
                        DnNo = ListSource.SelectedValue;

                        if (!NetCheck(DnNo)) TF = false;
                    }

                    if (GetValue("select count(*) from [����] where [�U��s��]=" + UpNo) == "0")
                    {
                    //    TF = false;
                        AddMsg("<script>alert('�@���W�大�j��(" + ListSource.SelectedItem.Text + ")�L�����W��I�Y���ݭn�A�зs�W��(�`����)���U�C');</script>");
                    }
                }                
            }            

            if (TF) //�����`�I���W�U���t�����@�`�I�A�s�W���W�U��(������e)
            {
                if (tbl == "���q") TreeNos = FormatChilds(GetTreeNos(tbl, TextDevNo.Text, "Up") + "," + GetTreeNos(tbl, TextDevNo.Text, "Dn"));
                else if (UpDn == "Up") TreeNos = GetTreeNos(tbl, TextDevNo.Text, "Dn");
                else TreeNos = GetTreeNos(tbl, TextDevNo.Text, "Up");

                if (("," + TreeNos + ",").IndexOf("," + ListSource.SelectedValue + ",") >= 0)
                {
                    TF = false;
                    if (tbl == "���q") AddMsg("<script>alert('�����`�I���W�U���t�����@�`�I�A�s�W���W�U��I');</script>");
                    else if (UpDn == "Up") AddMsg("<script>alert('�����`�I���U���t�����@�`�I�A�s�W���W��I');</script>");
                    else AddMsg("<script>alert('�����`�I���W���t�����@�`�I�A�s�W���U��I');</script>");
                }
                else
                {
                    string SQL = "Insert into [" + tbl + "] values(" + UpNo + "," + DnNo + ")";
                    ExecDbSQL(SQL);
                    InsLifeLog("�s�W [" + lblDevName.Text + "] ��" + UpDnC + tbl + "�j�� [" + ListSource.SelectedItem.Text + "]�G" + SQL);
                    AddMsg("<script>alert('�w�����s�W" + UpDnC + tbl + "�j���G" + ListSource.SelectedItem.Text + "');</script>");
                    ReadConn(TextDevNo.Text, tbl, UpDn, ListTarget);  //Ū���]�ưj�����                
                }
            }
            else
            {
                if (tbl == "���q") AddMsg("<script>alert('���q�j���z�S���]�w���ʧ@���v���I�Y�A�����Ҥp�աA�ЦV�䬢�ߡI');</script>");
                if (tbl == "����") AddMsg("<script>alert('�����j���z�S���]�w���ʧ@���v���I�Y�A�κ��ޤp�աA�ЦV�䬢�ߡI');</script>");
            }
        }        
    }

    protected string GetTreeNos(string tbl, string PkNo,string UpDn)   //�̬Y�`�I���o�Ҧ���t�W�U��s��
    {
        string SQL = "select [�U��s��] from [" + tbl + "] where [�W��s��]=" + PkNo;
        if (UpDn == "Up") SQL = "select [�W��s��] from [" + tbl + "] where [�U��s��]=" + PkNo;
        string cfg = "";

        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
        Conn.Open();
        SqlCommand cmd = new SqlCommand(SQL, Conn);
        SqlDataReader dr = null;
        dr = cmd.ExecuteReader();
        while (dr.Read()) cfg = cfg + dr[0].ToString() + "," + GetTreeNos(tbl, dr[0].ToString(), UpDn) + ",";

        cmd.Cancel(); cmd.Dispose(); dr.Close(); Conn.Close(); Conn.Dispose();

        return (FormatChilds(cfg));
    }

    protected string FormatChilds(string PkNos)    //�榡�ƤU��r��
    {
        string cfg = PkNos.Replace(",0,", ",").Replace(",,", ",");
        if (cfg == "" | cfg == ",") cfg = "0";
        if (cfg.Substring(cfg.Length - 1) == ",") cfg = cfg.Substring(0, cfg.Length - 1);
        if (cfg.Substring(0, 1) == ",") cfg = cfg.Substring(1);

        return (cfg);
    } 

    //�R���j��----------------------------------------------------------------
    protected void LinkUpPowerDel_Click(object sender, EventArgs e)
    {
        LinkDel_Click(ListUpPower);
    }
    protected void LinkDnPowerDel_Click(object sender, EventArgs e)
    {
        LinkDel_Click(ListDnPower);
    }
    protected void LinkUpNetDel_Click(object sender, EventArgs e)
    {
        LinkDel_Click(ListUpNet);
    }
    protected void LinkDnNetDel_Click(object sender, EventArgs e)
    {
        LinkDel_Click(ListDnNet);
    }

    protected void LinkDel_Click(ListBox ListTarget)
    {
        string tbl = "", UpDn = "", UpDnC = "", UpNo = "", DnNo = "";
        Boolean TF = false;
        
        if (ListTarget.SelectedValue == "") AddMsg("<script>alert('�Х�����z�n�R�����j���I');</script>");
        else if (RightCheck())
            {
                TF = true;
                switch (ListTarget.ID)
                {
                    case "ListUpPower":
                        {
                            tbl = "���q";
                            UpDn = "Up"; UpDnC = "�W��";
                            UpNo = ListTarget.SelectedValue;
                            DnNo = TextDevNo.Text;

                            if (ListTarget.Items.Count <= 1 & ListDnPower.Items.Count > 0)
                            {
                                TF = false;
                                AddMsg("<script>alert('�W�嬰��j���B�|���U��ɸT��R���W��I');</script>");
                            }
                            break;
                        }
                    case "ListDnPower":
                        {
                            tbl = "���q";
                            UpDn = "Dn"; UpDnC = "�U��";
                            UpNo = TextDevNo.Text;
                            DnNo = ListTarget.SelectedValue;

                            string Count = GetValue("select count(*) from [" + tbl + "] where [�U��s��]=" + DnNo);
                            if ((Count == "0" | Count == "1") & GetValue("select [�U��s��] from [" + tbl + "] where [�W��s��]=" + DnNo) != "")
                            {
                                TF = false;
                                AddMsg("<script>alert('�U�嬰��j���B�U��|���U��ɸT��R���U��I');</script>");
                            }
                            break;
                        }
                    case "ListUpNet":
                        {
                            tbl = "����";
                            UpDn = "Up"; UpDnC = "�W��";
                            UpNo = ListTarget.SelectedValue;
                            DnNo = TextDevNo.Text;
                            if (ListTarget.Items.Count <= 1 & ListDnNet.Items.Count > 0)
                            {
                                TF = false;
                                AddMsg("<script>alert('�W�嬰��j���B�|���U��ɸT��R���W��I');</script>");
                            }

                            if (!NetDelCheck(UpNo)) TF = false;
                            break;
                        }
                    case "ListDnNet":
                        {
                            tbl = "����";
                            UpDn = "Dn"; UpDnC = "�U��";
                            UpNo = TextDevNo.Text;
                            DnNo = ListTarget.SelectedValue;
                            string Count = GetValue("select count(*) from [" + tbl + "] where [�U��s��]=" + DnNo);
                            if ((Count == "0" | Count == "1") & GetValue("select [�U��s��] from [" + tbl + "] where [�W��s��]=" + DnNo) != "")
                            {
                                TF = false;
                                AddMsg("<script>alert('�U�嬰��j���B�U��|���U��ɸT��R���U��I');</script>");
                            }

                            if (!NetDelCheck(DnNo)) TF = false;
                            break;
                        }
                }                
            }
        else AddMsg("<script>alert('�z�S���R���j�����v���A�Y�A�κ��ޤp�ձ����]�ơA�ЦV�䬢�ߡI');</script>");

        if (TF)
        {
            string SQL = "delete from [" + tbl + "] where [�U��s��]=" + DnNo + " and [�W��s��]=" + UpNo;
            ExecDbSQL(SQL);
            InsLifeLog("�R�� [" + lblDevName.Text + "] ��" + UpDnC + tbl + "�j�� [" + ListTarget.SelectedItem.Text + "] �G " + SQL);
            AddMsg("<script>alert('�w�����R��" + UpDnC + tbl + "�j���G" + ListTarget.SelectedItem.Text + "');</script>");
            ReadConn(TextDevNo.Text, tbl, UpDn, ListTarget);  //Ū���]�ưj�����                    
        }
		else
		{
			if (tbl == "���q") AddMsg("<script>alert('���q�j���z�S���]�w���ʧ@���v���I�Y�A�����Ҥp�աA�ЦV�䬢�ߡI');</script>");
			if (tbl == "����") AddMsg("<script>alert('�����j���z�S���]�w���ʧ@���v���I�Y�A�κ��ޤp�աA�ЦV�䬢�ߡI');</script>");
		}		
    }

    //�s��γ]�w�j��----------------------------------------------------------------
    protected void LinkUpPowerEdit_Click(object sender, EventArgs e)
    {
        LinkEdit_Click(ListUpPower, "DevEdit.aspx");
    }
    protected void LinkDnPowerEdit_Click(object sender, EventArgs e)
    {
        LinkEdit_Click(ListDnPower, "DevEdit.aspx");
    }
    protected void LinkUpNetEdit_Click(object sender, EventArgs e)
    {
        LinkEdit_Click(ListUpNet, "DevEdit.aspx");
    }
    protected void LinkDnNetEdit_Click(object sender, EventArgs e)
    {
        LinkEdit_Click(ListDnNet, "DevEdit.aspx");
    }
    protected void LinkSourceEdit_Click(object sender, EventArgs e)
    {
        LinkEdit_Click(ListSource, "DevEdit.aspx");
    }

    protected void LinkUpPowerConn_Click(object sender, EventArgs e)
    {
        LinkEdit_Click(ListUpPower, "TreeEdit.aspx");
    }
    protected void LinkDnPowerConn_Click(object sender, EventArgs e)
    {
        LinkEdit_Click(ListDnPower, "TreeEdit.aspx");
    }
    protected void LinkUpNetConn_Click(object sender, EventArgs e)
    {
        LinkEdit_Click(ListUpNet, "TreeEdit.aspx");
    }
    protected void LinkDnNetConn_Click(object sender, EventArgs e)
    {
        LinkEdit_Click(ListDnNet, "TreeEdit.aspx");
    }
    protected void LinkSourceConn_Click(object sender, EventArgs e)
    {
        LinkEdit_Click(ListSource, "TreeEdit.aspx");
    }

    protected void LinkEdit_Click(ListBox ListName,string code)
    {
        if (ListName.SelectedValue == "") AddMsg("<script>alert('�Х�����z�n�s��γ]�w���j���I');</script>");
        else OpenExecWindow("window.open('" + code + "?DevNo=" + ListName.SelectedValue + "','_self');");
    }

    //�i�}�X�ֳ]�ưj��----------------------------------------------------------------
    protected void LinkSourceUp_Click(object sender, EventArgs e)  //�̩ҿ���j���C�X��W��
    {
        if (ListSource.SelectedValue == "") AddMsg("<script>alert('�Х�����ݿ�M�涵�ءI');</script>");
        else
        {
            SqlDataSourceSource.SelectCommand = "SELECT [�]�ƽs��],[�]�ƦW��] FROM [View_�]�ƺ޲z] "
                + "WHERE [�]�ƽs��] in (select [�W��s��] from [View_�]�ưj��] where [�U��s��]=" + ListSource.SelectedValue + ") ORDER BY [�]�ƦW��]";
            ListSource.DataBind();
        }
    } 
    protected void LinkSourceDn_Click(object sender, EventArgs e)  //�̩ҿ���j���C�X��U��
    {
        if (ListSource.SelectedValue == "") AddMsg("<script>alert('�Х�����ݿ�M�涵�ءI');</script>");
        else
        {
            SqlDataSourceSource.SelectCommand = "SELECT [�]�ƽs��],[�]�ƦW��] FROM [View_�]�ƺ޲z] "
                + "WHERE [�]�ƽs��] in (select [�U��s��] from [View_�]�ưj��] where [�W��s��]=" + ListSource.SelectedValue + ") ORDER BY [�]�ƦW��]";
            ListSource.DataBind();            
        }
    }     

    //�����m�a�I------------------------------------------------------------------------ 
    protected void SelPlace_SelectedIndexChanged(object sender, EventArgs e)
    {
        Place_Selected(ListSource, SelPlace, SelPointer, SqlDataSourcePointer, SqlDataSourceSource);
    }
    protected void Place_Selected(ListBox ListName, DropDownList SelPlace, DropDownList SelPointer, SqlDataSource SqlDS, SqlDataSource SqlDSList)   
    {
        ListName.Items.Clear();
        SqlDS.SelectCommand = "SELECT [�w��s��],[�w��W��] FROM [�w��]�w] WHERE [�w��覡]<>'����' AND [�ϰ�W��]='" + SelPlace.SelectedValue + "' order by [�w��W��]";
        SelPointer.DataBind();
        Pointer_Selected(ListName, SelPointer, SqlDSList);
    }
    //����w��W��------------------------------------------------------------------------ 
    protected void SelPointer_SelectedIndexChanged(object sender, EventArgs e)
    {
        Pointer_Selected(ListSource, SelPointer, SqlDataSourceSource);
    }
    protected void Pointer_Selected(ListBox ListName, DropDownList SelPointer, SqlDataSource SqlDS)
    {
        string SQL = "[�w��s��]=" + SelPointer.SelectedValue;
        if (SelPointer.SelectedValue == "") SQL = "[�w��s��]=" + SelPointer.Items[0].Value;

        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
        Conn.Open();
        //[�w��s��],[�w��覡],[�ϰ�W��],[�w��W��],[���d�e�q],[����X],[����Y],[����X1],[����Y1],[����X2],[����Y2]
        SqlCommand cmd = new SqlCommand("select * from [�w��]�w] where " + SQL, Conn);
        SqlDataReader dr = null;
        dr = cmd.ExecuteReader();
        if (dr.Read())
        {
            //�w��覡�Y�����ϡA�h��ܥX�H���Щw�줧�]�ơA�H�K�藍��]��
            if (dr["�w��覡"].ToString() == "����") SQL = "[�w��覡]<>'���d' and [����X] between " + dr["����X1"].ToString() + " and " + dr["����X2"].ToString() + " and [����Y] between " + dr["����Y1"].ToString() + " and " + dr["����Y2"].ToString();
            SqlDS.SelectCommand = "SELECT [�]�ƽs��],[�]�ƦW��] FROM [View_�]�ƺ޲z] WHERE " + SQL + " ORDER BY [�]�ƦW��]";
            ListName.DataBind();
        }
        cmd.Cancel(); cmd.Dispose(); dr.Close(); Conn.Close(); Conn.Dispose();
    }

    //��ܰj�����|�Ω�m�a�I------------------------------------------------------------------------ 
    protected void ListSource_SelectedIndexChanged(object sender, EventArgs e)
    {
        Conn_Selected(ListSource, lblSource, "�]��");
    }

    protected void ListUpPower_SelectedIndexChanged(object sender, EventArgs e) 
    {
        Conn_Selected(ListUpPower, lblUpPower,"���q");
    }
    protected void ListDnPower_SelectedIndexChanged(object sender, EventArgs e) 
    {
        Conn_Selected(ListDnPower, lblDnPower, "���q");
    }

    protected void ListUpNet_SelectedIndexChanged(object sender, EventArgs e) 
    {
        Conn_Selected(ListUpNet, lblUpNet, "����");
    }
    protected void ListDnNet_SelectedIndexChanged(object sender, EventArgs e)
    {
        Conn_Selected(ListDnNet, lblDnNet, "����");
    }

    protected void Conn_Selected(ListBox ListName,Label lbl,string tbl) //�I��ﶵ�A�C�X��m�a�I�P�j�����|
    {
        lbl.Text = GetValue("select [��m�a�I] from [View_�]�ƺ޲z] where [�]�ƽs��]=" + ListName.SelectedValue);

        if (tbl == "���q" | tbl == "�]��") lbl.Text = lbl.Text + "<br />" + GetPath(ListName.SelectedValue, "���q", "") + "��";
        if (tbl == "����" | tbl == "�]��") lbl.Text = lbl.Text + "<br />" + GetPath(ListName.SelectedValue, "����", "") + "��";        
    }
    protected string GetPath(string DevNo,string tbl, string tail)   //���o�j�����|
    {
        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
        Conn.Open();
        SqlCommand cmd = new SqlCommand("select [�W��s��],[�]�ƦW��] from [" + tbl + "],[����]��] where [" + tbl + "].[�W��s��]=[����]��].[�]�ƽs��] and [" + tbl + "].[�U��s��]=" + DevNo, Conn);
        SqlDataReader dr = null;
        dr = cmd.ExecuteReader();
        string path = "", UpPath = "", DnPath = "";

        int i = 0;
        while (dr.Read())
        {
            i++;
            DnPath = dr["�]�ƦW��"].ToString() + " �� " + tail;
            UpPath = GetPath(dr["�W��s��"].ToString(), tbl, DnPath);
            if (i == 1) path = UpPath + DnPath;
            else path = path + "��<br />" + UpPath + DnPath;

            path = path.Replace(DnPath + DnPath, DnPath);
        }

        cmd.Cancel(); cmd.Dispose(); dr.Close(); Conn.Close(); Conn.Dispose();        
        return (path);
    }

    //�v���޲z------------------------------------------------------------------------ 
    protected Boolean RightCheck() //�O�_���v�ק���
    {
        string UserID = Session["UserID"].ToString().ToLower();
        string UserName = Session["UserName"].ToString();   //�n�J��UserName
        string UnitName = Session["UnitName"].ToString();   //�n�J��UnitName
        string Hw = lblMt.Text; //�]�ƺ��@�H��
        string Dno = TextDevNo.Text; if (Dno == "") Dno = "0";
        string Older = GetValue("select [���@�H��] from [����]��] where [�]�ƽs��]=" + Dno);

        if (UserID != "operator" & (InGroup(UserName, Hw) | InGroup(UserName, Older) | UnitName == "�t�α����" | UnitName == "�q���ާ@��" | InGroup(UserName, "���ޤp��"))) return (true);
        else return (false);
    }
    protected Boolean NetCheck(string Dno) //�O�_���v�קﱵ���j��
    {
        string Mt = "���ޤp��";     //���ޤp���v�d�]�Ʊ����W�U��Ⱥ��ޤp�եi�]�w
        string UserName = Session["UserName"].ToString();   //�n�J��UserName
        string Hw = lblMt.Text;     //�]�ƺ��@�H��
        string Dw = GetValue("select [���@�H��] from [����]��] where [�]�ƽs��]=" + Dno);   //�W�U����@�H��

        if ((Hw != Mt & Dw != Mt) | InGroup(UserName, Mt) | InGroup(UserName, "WII�p��")) return (true);   //�]�ƩΤW�U���v�d���D���ޤp�աA�ά����ަ���
        else return (false);
    }
	protected Boolean NetDelCheck(string Dno) //�O�_���v �R�� �����j��
    {
        string Mt = "���ޤp��";     //���ޤp���v�d�]�Ʊ����W�U��Ⱥ��ޤp�եi�]�w
        string UserName = Session["UserName"].ToString();   //�n�J��UserName
        string Hw = lblMt.Text;     //�]�ƺ��@�H��
        string Dw = GetValue("select [���@�H��] from [����]��] where [�]�ƽs��]=" + Dno);   //�W�U����@�H��
        
        if ((Hw != Mt & Dw != Mt) | InGroup(UserName, Mt) | InGroup(UserName, "WII�p��") | (Hw==UserName)| InGroup(UserName,Hw)) return (true);   //�]�ƩΤW�U���v�d���D���ޤp�աA�ά����ަ��� �A�ά��]�ƭt�d�H(�Ҧb�s��)
        else return (false);
    }
    protected Boolean InGroup(string ChkName, string ChkUnit) //�ˬdChkName�O�_��ChkUnit�����Υ���
    {
        Boolean TF = false;

        if (ChkName == ChkUnit) TF = true;  //�O�_�P�W
        else if (ChkUnit == "") TF = false;	//�ˬd��쥲��
        else //�O�_������UN (�ҧO�P�p�զP�q)
        {
            SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
            Conn.Open();
            SqlCommand cmd = new SqlCommand("select * from Config where (Kind='" + ChkUnit + "') and Item='" + ChkName + "'", Conn);
            SqlDataReader dr = null;
            dr = cmd.ExecuteReader();
            if (dr.Read()) TF = true;
            cmd.Cancel(); cmd.Dispose(); dr.Close(); Conn.Close(); Conn.Dispose();
        }

        return (TF);
    }
    protected void InsLifeLog(string SQL) //�g�J�ͩR�i��
    {
        string LifeNo = GetPKNo("�i���s��", "�ͩR�i��").ToString(); //�i���s��
        string TblName = "�]�ưj��";    //���W��
        string PKno = TextDevNo.Text;   //�D��s��
        string Mt = lblMt.Text;    //���@�H��
        string OldMt = GetValue("select [���@�H��] from [����]��] where [�]�ƽs��]=" + TextDevNo.Text);    //��t�d�H
        string OldKeeper = GetValue("select [�O�ޤH��] from [View_�]�ƺ޲z] where [�]�ƽs��]=" + TextDevNo.Text);    //��O�ޤH
        string UN = Session["UserName"].ToString();   //�n�J��UserName�G���ʤH��
        string LiftDT = DateTime.Now.ToString("yyyy/MM/dd HH:mm");  //���ʤ��
        string LifeIP = Request.ServerVariables["REMOTE_ADDR"].ToString();

        ExecDbSQL("Insert into [�ͩR�i��] values(" + LifeNo + ",'" + TblName + "'," + PKno + ",'" + SQL.Replace("'", "''") + "','" + Mt + "','" + OldMt + "','" + OldKeeper + "','" + UN + "','" + LiftDT + "','" + LifeIP + "')");
    }
    protected int GetPKNo(string PKfield, string PKtbl) //���o�D��s��
    {
        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
        Conn.Open();
        SqlCommand cmd = new SqlCommand("select max(" + PKfield + ") from " + PKtbl, Conn);
        SqlDataReader dr = null;
        dr = cmd.ExecuteReader();
        int PkNo = 1;
        if (dr.Read()) PkNo = int.Parse(dr[0].ToString()) + 1;
        cmd.Cancel(); cmd.Dispose(); dr.Close(); Conn.Close(); Conn.Dispose();

        return (PkNo);
    }
    protected void BtnEdit_Click(object sender, EventArgs e) //�^�s�譶
    {
        OpenExecWindow("window.open('DevEdit.aspx?DevNo=" + TextDevNo.Text + "','_self');");
    }
    protected void BtnLife_Click(object sender, EventArgs e) //�d�ߥͩR�i��
    {
        Session["LifeSQL"] = "select * from [�ͩR�i��] where [���W��]='�]�ưj��' and [�D��s��]=" + TextDevNo.Text + " order by [�i���s��] desc";
        OpenExecWindow("window.open('../Life/LifeLog.aspx?Search=Conn&Tbl=�]�ưj��&PK=" + TextDevNo.Text + "','_self');");
    } 
    //���Ψ��------------------------------------------------------------------------ 
    protected DataSet RunQuery(SqlCommand sqlQuery) //Ū��DB��T
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

    protected void ExecDbSQL(string SQL) //�����Ʈw����
    {
        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
        Conn.Open();
        SqlCommand cmd = new SqlCommand(SQL, Conn);
        //Response.Write(cmd.CommandText);
        //Response.End();
        cmd.ExecuteNonQuery();
        cmd.Cancel(); cmd.Dispose(); Conn.Close(); Conn.Dispose();
    }

    protected string GetValue(string SQL)   //���o��@���
    {
        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
        Conn.Open();
        SqlCommand cmd = new SqlCommand(SQL, Conn);
        SqlDataReader dr = null;
        dr = cmd.ExecuteReader();
        string cfg = ""; if (dr.Read()) cfg = dr[0].ToString();

        cmd.Cancel(); cmd.Dispose(); dr.Close(); Conn.Close(); Conn.Dispose();

        return (cfg);
    }

    protected void OpenExecWindow(string strJavascript)    //�����N�t�}����
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

    protected void AddMsg(string strMsg)
    {
        Literal Msg = new Literal();
        Msg.Text = strMsg;
        Page.Controls.Add(Msg);
    }
}