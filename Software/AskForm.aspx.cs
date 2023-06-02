using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.Configuration;
using System.Data;
using System.Data.SqlClient;

public partial class SoftWare_AskForm : System.Web.UI.Page
{
    string AskNo = "";

    protected void Page_Load(object sender, EventArgs e)
    {
        if (Request["AskNo"] != null & Request["AskNo"] != "")
        {
            AskNo = Request["AskNo"].ToString();
            ReadForm(AskNo);
        }
    }

    protected void Page_PreRenderComplete(object sender, EventArgs e)   //��Page_Load�L�k���
    {
        Literal Msg = new Literal();
        Msg.Text = "<script>alert('[��ĳ]�G\\n"
                                + "1. �Х��w���C�L�վ�]�w�A�ϯ�Τ@�i�ȴN�n\\n"
                                + "2. �Ш����Ҧ����������]�w\\n"
                                + "3. �w�]���w�ˡA���D�̷s�ӽг�[�ӽШƶ�]��`����`��`����`�r��\\n"
                                + "4. �i�⦡�q���ϥηL�n�n��G�����(NB)\\n"
                                + "   a.[�]�ƺ���]�D���O���q���A�h�פJ�D�ϥ�\\n"
                                + "   b.[�]�ƺ���]�����O���q���A�h�פJ���q\\n"                                
                                + "5. �L�n���v�n��ϥε����ӽг�G�եΪ�\\n"                                
                                + "');</script>";
        Page.Controls.Add(Msg);
    }

    protected void ReadForm(string AskNo)    //Ū�����v���
    {
        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
        Conn.Open();
        SqlCommand cmd = new SqlCommand("SELECT * FROM [View_�n��޲z] WHERE [���v�s��]=" + AskNo, Conn);
        SqlDataReader dr = null;
        dr = cmd.ExecuteReader();

        if (dr.Read())
        {
            //--------------------------�O�޳��T--------------------------------------------------------------------------------------------------------------
            //lblAskNo.Text = AskNo; 
            if (dr["���v���"].ToString() != "") lblAskUnit.Text = dr["���v���"].ToString();
            if (!(dr["���v�����"] is DBNull)) lblKeyDay.Text = DateTime.Parse(dr["���v�����"].ToString()).ToString("yyyy �~ MM �� dd ��");
            lblSwName.Text = "YYYY-NNNN"; ; if (dr["�n��W��"].ToString() != "") lblSwName.Text = dr["�n��W��"].ToString();
            if (dr["�ʶR����"].ToString() != "") lblVer.Text = dr["�ʶR����"].ToString();
            if (dr["���v�D��"].ToString() != "") lblHost.Text = dr["���v�D��"].ToString();
            if (dr["���vIP"].ToString() != "") lblIP.Text = dr["���vIP"].ToString();

            string Ask = "�w��";
            ChkIns.Checked = true;
            //--------------------------�ӽг��T--------------------------------------------------------------------------------------------------------------
            SqlConnection Conn1 = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
            Conn1.Open();
            SqlCommand cmd1 = new SqlCommand("select * from [�ӽЪ��] where [������]='�n��ӽ�' and [�D��s��]=" + AskNo + " order by [�����] desc", Conn1);
            SqlDataReader dr1 = null;
            dr1 = cmd1.ExecuteReader();
            if (dr1.Read())
            {
                lblFormNo.Text = "YYYY-NNNN"; lblFormNo.ForeColor = System.Drawing.Color.LightGray; lblFormNo.Font.Bold = false;
                if (dr1["���s��"].ToString() != "")
                {
                    lblFormNo.Text = dr1["���s��"].ToString();
                    if (lblFormNo.Text.Length > 9) lblFormNo.Text = lblFormNo.Text.Substring(0, 9);
                    lblFormNo.ForeColor = System.Drawing.Color.Black; lblFormNo.Font.Bold = true;
                }

                lblKeyDay.Text = "�@�~�@��@��";
                if (!(dr1["�����"] is DBNull)) lblKeyDay.Text = DateTime.Parse(dr1["�����"].ToString()).ToString("yyyy �~ MM �� dd ��");

                Ask = dr1["�ӽШƶ�"].ToString();
            }
            cmd1.Cancel(); cmd1.Dispose(); dr1.Close(); Conn1.Close(); Conn1.Dispose();

            //���v�覡�G������B�����(NB)�B�}�o�I�B�H�����B�������BServer���BServer��(V1)�BServer��(V4)�BServer��(V��)�BClient���BServer & Client���B�ɯŪ��B�䥦�A������[�n��D��]�A���ο�J
            switch (dr["�ӽб��v"].ToString())
            {
                case "�����(NB)": //�@����@NB(Select �j�q���v)
                    {
                        ChkNB.Checked = true;
                        if (!(dr["�]�ƺ���"] is DBNull))
                        {
                            if (dr["�]�ƺ���"] == "���O���q��")
                            {
                                if (dr["���v�D��"].ToString() != "") lblNBHost.Text = dr["���v�D��"].ToString() + "�@";
                                if (dr["���vIP"].ToString() != "") lblNBIP.Text = dr["���vIP"].ToString() + "�@";
                            }
                            else
                            {
                                if (dr["���v�D��"].ToString() != "") lblSelHost.Text = dr["���v�D��"].ToString() + "�@";
                                if (dr["���vIP"].ToString() != "") lblSelIP.Text = dr["���vIP"].ToString() + "�@";
                            }
                        }
                        break;
                    }
                case "�եΪ�": //60��n��ϥε���(Select)
                    {
                        ChkTest.Checked = true;
                        if (dr["���v�D��"].ToString() != "") lblTestHost.Text = dr["���v�D��"].ToString() + "�@";
                        if (dr["���vIP"].ToString() != "") lblTestIP.Text = dr["���vIP"].ToString() + "�@";
                        break;
                    }
            }

            if (Ask.IndexOf("����") >= 0)
            {
                ChkIns.Checked = false;
                ChkDel.Checked = true;
            }

            if (Ask.IndexOf("����") >= 0)
            {
                ChkIns.Checked = false;
                ChkUpd.Checked = true;
                if (dr["�ϥΥD��"].ToString() != "") lblOldHost.Text = dr["�ϥΥD��"].ToString() + "�@";
                if (dr["�ϥ�IP"].ToString() != "") lblOldIP.Text = dr["�ϥ�IP"].ToString() + "�@";
                if (dr["���v�D��"].ToString() != "") lblNewHost.Text = dr["���v�D��"].ToString() + "�@";
                if (dr["���vIP"].ToString() != "") lblNewIP.Text = dr["���vIP"].ToString() + "�@";
            }
        }

        cmd.Cancel(); cmd.Dispose(); dr.Close(); Conn.Close(); Conn.Dispose();
    }

    protected string GetValue(string DB, string SQL)   //���o��@���
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
}