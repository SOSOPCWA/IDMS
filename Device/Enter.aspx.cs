using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.Configuration;
using System.Data;
using System.Data.SqlClient;

public partial class Device_Enter : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        if (Request["DevNo"] != null & Request["DevNo"] != "") ReadEnter(Request["DevNo"]); 
    }

    protected void Page_PreRenderComplete(object sender, EventArgs e)   //��Page_Load�L�k���
    {
        Literal Msg = new Literal();
        Msg.Text = "<script>alert('1. �аȥ����վ�W�U��ɡA�������������A�������C�L�A�ϯ�Τ@�i�ȴN�n�I\\n\\n"
                                + "2. �h�]�ƥi��i���d��������ﲾ�J��榡�A�ץXExcel�C�L�@������');</script>";
        Page.Controls.Add(Msg);
    }

    protected void ReadEnter(string DevNo)    //Ū������]��
    {
        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
        Conn.Open();
        SqlCommand cmd = new SqlCommand("SELECT * FROM [View_�q�γ]��] WHERE [�]�ƽs��]=" + DevNo, Conn);
        SqlDataReader dr = null;
        dr = cmd.ExecuteReader();

        if (dr.Read())
        {
            lblDevNo.Text = DevNo;

            string Hw = dr["�]�ƺ��@�H��"].ToString(); //if (Hw != "") lblAsker.Text = Hw; else lblAsker.Text = "&nbsp;";
            //string Unit = GetUnit(Hw, "�ҧO"); if (Unit != "") lblAskUnit.Text = Unit; else lblAskUnit.Text = "&nbsp;";          
            
            //lblAskDate.Text = DateTime.Now.ToString("yyyy/MM/dd");
            lblDevName.Text = "�]�ƦW��:"+dr["�]�ƦW��"].ToString(); if (lblDevName.Text == "") lblDevName.Text = "�]�ƦW��:&nbsp;";
            lblIP.Text = "IP:"+dr["IP"].ToString(); if (lblIP.Text == "") lblIP.Text = "IP:&nbsp;";
            //lblEnterDate.Text = DateTime.Now.ToString("yyyy/MM/dd");
            lblPurpose.Text ="�]�ƥγ~:"+ dr["�]�ƥγ~"].ToString(); if (lblPurpose.Text == "") lblPurpose.Text = "�]�ƥγ~:&nbsp;";

            string Dep = GetUnit(Hw, "���");
            if (Dep == "��T����") ChkAP.Checked = true;
            else if (Dep != "") ChkAgent.Checked = true;

            lblHw.Text ="�w��t�d�H:"+ Hw; if (lblHw.Text == "") lblHw.Text = "�w��t�d�H:&nbsp;";
            //lblRepair.Text = dr["���@�t��"].ToString(); if (lblRepair.Text == "") lblRepair.Text = "&nbsp;";
            
            string Sw = "�n��t�d�H:"+dr["�@�~���@�H��"].ToString();
            lblSw.Text = Sw; if (Sw == "") lblSw.Text = "�n��t�d�H:&nbsp;";

            switch (dr["�]�ƫ��A"].ToString())
            {
                case "�t�γ]��":
                    if (dr["�]�ƺ���"].ToString() == "�x�s�]��") ChkST.Checked = true;
                    else ChkSys.Checked = true;
                    break;
                case "�����]��": ChkNet.Checked = true; break;
                case "���ҳ]��": ChkEnv.Checked=true; break;
                default: ChkElse.Checked=true; break;
            }
                        
            if (dr["�w��覡"].ToString() == "���d") ChkRack.Checked = true;
            else ChkXY.Checked = true;
			lblPlaceName.Text ="�a�I:"+ dr["�ϰ�W��"].ToString();
            lblPlace.Text = "���d�s��:"+ dr["�w��W��"].ToString();
            if (dr["����"].ToString() != "0") lblHeight.Text =  dr["����"].ToString() + "U - "  + (int.Parse(dr["����"].ToString()) + int.Parse(dr["�Ŷ��j�p"].ToString())-1).ToString()+ "U";

            if (dr["�ιq�q��"].ToString() != "0") lblVoltage.Text = dr["�ιq�q��"].ToString();
            if (dr["�B�w�q�y"].ToString() != "0") lblCurrent.Text = dr["�B�w�q�y"].ToString();
            string strKVA = (int.Parse(dr["�ιq�q��"].ToString()) * float.Parse(dr["�B�w�q�y"].ToString()) / 1000).ToString();
            if (strKVA != "0") lblKVA.Text = strKVA;

            string strBTU = String.Format("{0:F0}", int.Parse(dr["�ιq�q��"].ToString()) * float.Parse(dr["�B�w�q�y"].ToString()) * 3.41);
            if (strBTU != "0") lblBTU.Text = strBTU;


            SqlConnection Conn1 = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
            Conn1.Open();
            SqlCommand cmd1 = new SqlCommand("SELECT * FROM [����]��],[���q] WHERE [����]��].[�]�ƽs��]=[���q].[�W��s��] and [���q].[�U��s��]=" + DevNo, Conn1);
            SqlDataReader dr1 = null;            

            dr1 = cmd1.ExecuteReader();
			bool isUps=false;
			bool isUpsS=false;
			bool isUpsI=false;
            while (dr1.Read()) 
            {
				if(isUpsS==false){
					lblUps1S.Text = dr1["�]�ƦW��"].ToString();
					isUpsS=true;
				}
				else lblUps2S.Text = dr1["�]�ƦW��"].ToString();
				
				if (dr1["�B�w�q�y"].ToString() != "0"){
					if(isUpsI==false){
						lblUps1I.Text = dr1["�B�w�q�y"].ToString();
						isUpsI=true;
					}
					else lblUps2I.Text = dr1["�B�w�q�y"].ToString();
				}
				
				string upsName="";
                switch (GetUps(dr1["�W��s��"].ToString()))
                {
                    case "2587":
                        ChkUps550.Checked = true;
                        //if (dr1["�B�w�q�y"].ToString() != "0") lblUps375I.Text = dr1["�B�w�q�y"].ToString();
						upsName="550N";
                        break;
                    case "83": 
                        ChkUps550.Checked = true;
                        //if (dr1["�B�w�q�y"].ToString() != "0") lblUps550I.Text = dr1["�B�w�q�y"].ToString();
                        //lblUps550S.Text = dr1["�]�ƦW��"].ToString();
						upsName="550 ";
                        break;
                    case "84":
                        ChkUps550.Checked = true;
                        //if (dr1["�B�w�q�y"].ToString() != "0") lblUps600I.Text = dr1["�B�w�q�y"].ToString();
                        //lblUps600S.Text = dr1["�]�ƦW��"].ToString();
						upsName="600 ";
                        break;
                    case "":                                              
                        break;
                    default:    //���q
                        ChkUpsNo.Checked = true;
                        break;
                }
				
				if(isUps==false){
					lblUps1.Text = upsName;
					isUps=true;
				}
				else lblUps2.Text = upsName;
				
            }
            cmd1.Cancel(); cmd1.Dispose(); dr1.Close(); Conn1.Close(); Conn1.Dispose();
        }

        cmd.Cancel(); cmd.Dispose(); dr.Close(); Conn.Close(); Conn.Dispose();
    }

    protected string GetUnit(string Pname, string which) //Ū���ҧO�γ��
    {
        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
        Conn.Open();
        SqlCommand cmd = new SqlCommand("SELECT [" + which + "] FROM [View_��´�[�c] WHERE [����]='" + Pname + "'", Conn);
        SqlDataReader dr = null;

        string Unit = "";

        dr = cmd.ExecuteReader();
        if (dr.Read())  Unit = dr[0].ToString();

        cmd.Cancel(); cmd.Dispose(); dr.Close(); Conn.Close(); Conn.Dispose();        

        return (Unit);
    }

    protected string GetUpNo(string DevNo) //���o�ҳs���j�����s��
    {
        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
        Conn.Open();
        SqlCommand cmd = new SqlCommand("SELECT [�W��s��] FROM [���q] WHERE [�U��s��]=" + DevNo, Conn);
        SqlDataReader dr = null;

        string UpNo = "";

        dr = cmd.ExecuteReader();
        if (dr.Read()) UpNo=dr[0].ToString();

        cmd.Cancel(); cmd.Dispose(); dr.Close(); Conn.Close(); Conn.Dispose();

        return (UpNo);
    }

    protected string GetUps(string DevNo) //���oUPS�s��
    {
        string UpsNo = "",tmpNo=DevNo;
        while(tmpNo != "-2")  //�`�q���s��
        {
            UpsNo = tmpNo;
            tmpNo = GetUpNo(tmpNo);
        }

        return (UpsNo);
    }
}