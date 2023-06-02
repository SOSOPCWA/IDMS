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

    protected void Page_PreRenderComplete(object sender, EventArgs e)   //放Page_Load無法顯示
    {
        Literal Msg = new Literal();
        Msg.Text = "<script>alert('1. 請務必先調整上下邊界，取消頁首頁尾，並雙面列印，使能用一張紙就好！\\n\\n"
                                + "2. 多設備可於進階查詢顯示欄位選移入單格式，匯出Excel列印作為附件');</script>";
        Page.Controls.Add(Msg);
    }

    protected void ReadEnter(string DevNo)    //讀取實體設備
    {
        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
        Conn.Open();
        SqlCommand cmd = new SqlCommand("SELECT * FROM [View_通用設備] WHERE [設備編號]=" + DevNo, Conn);
        SqlDataReader dr = null;
        dr = cmd.ExecuteReader();

        if (dr.Read())
        {
            lblDevNo.Text = DevNo;

            string Hw = dr["設備維護人員"].ToString(); //if (Hw != "") lblAsker.Text = Hw; else lblAsker.Text = "&nbsp;";
            //string Unit = GetUnit(Hw, "課別"); if (Unit != "") lblAskUnit.Text = Unit; else lblAskUnit.Text = "&nbsp;";          
            
            //lblAskDate.Text = DateTime.Now.ToString("yyyy/MM/dd");
            lblDevName.Text = "設備名稱:"+dr["設備名稱"].ToString(); if (lblDevName.Text == "") lblDevName.Text = "設備名稱:&nbsp;";
            lblIP.Text = "IP:"+dr["IP"].ToString(); if (lblIP.Text == "") lblIP.Text = "IP:&nbsp;";
            //lblEnterDate.Text = DateTime.Now.ToString("yyyy/MM/dd");
            lblPurpose.Text ="設備用途:"+ dr["設備用途"].ToString(); if (lblPurpose.Text == "") lblPurpose.Text = "設備用途:&nbsp;";

            string Dep = GetUnit(Hw, "單位");
            if (Dep == "資訊中心") ChkAP.Checked = true;
            else if (Dep != "") ChkAgent.Checked = true;

            lblHw.Text ="硬體負責人:"+ Hw; if (lblHw.Text == "") lblHw.Text = "硬體負責人:&nbsp;";
            //lblRepair.Text = dr["維護廠商"].ToString(); if (lblRepair.Text == "") lblRepair.Text = "&nbsp;";
            
            string Sw = "軟體負責人:"+dr["作業維護人員"].ToString();
            lblSw.Text = Sw; if (Sw == "") lblSw.Text = "軟體負責人:&nbsp;";

            switch (dr["設備型態"].ToString())
            {
                case "系統設備":
                    if (dr["設備種類"].ToString() == "儲存設備") ChkST.Checked = true;
                    else ChkSys.Checked = true;
                    break;
                case "網路設備": ChkNet.Checked = true; break;
                case "環境設備": ChkEnv.Checked=true; break;
                default: ChkElse.Checked=true; break;
            }
                        
            if (dr["定位方式"].ToString() == "機櫃") ChkRack.Checked = true;
            else ChkXY.Checked = true;
			lblPlaceName.Text ="地點:"+ dr["區域名稱"].ToString();
            lblPlace.Text = "機櫃編號:"+ dr["定位名稱"].ToString();
            if (dr["高度"].ToString() != "0") lblHeight.Text =  dr["高度"].ToString() + "U - "  + (int.Parse(dr["高度"].ToString()) + int.Parse(dr["空間大小"].ToString())-1).ToString()+ "U";

            if (dr["用電電壓"].ToString() != "0") lblVoltage.Text = dr["用電電壓"].ToString();
            if (dr["額定電流"].ToString() != "0") lblCurrent.Text = dr["額定電流"].ToString();
            string strKVA = (int.Parse(dr["用電電壓"].ToString()) * float.Parse(dr["額定電流"].ToString()) / 1000).ToString();
            if (strKVA != "0") lblKVA.Text = strKVA;

            string strBTU = String.Format("{0:F0}", int.Parse(dr["用電電壓"].ToString()) * float.Parse(dr["額定電流"].ToString()) * 3.41);
            if (strBTU != "0") lblBTU.Text = strBTU;


            SqlConnection Conn1 = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
            Conn1.Open();
            SqlCommand cmd1 = new SqlCommand("SELECT * FROM [實體設備],[接電] WHERE [實體設備].[設備編號]=[接電].[上游編號] and [接電].[下游編號]=" + DevNo, Conn1);
            SqlDataReader dr1 = null;            

            dr1 = cmd1.ExecuteReader();
			bool isUps=false;
			bool isUpsS=false;
			bool isUpsI=false;
            while (dr1.Read()) 
            {
				if(isUpsS==false){
					lblUps1S.Text = dr1["設備名稱"].ToString();
					isUpsS=true;
				}
				else lblUps2S.Text = dr1["設備名稱"].ToString();
				
				if (dr1["額定電流"].ToString() != "0"){
					if(isUpsI==false){
						lblUps1I.Text = dr1["額定電流"].ToString();
						isUpsI=true;
					}
					else lblUps2I.Text = dr1["額定電流"].ToString();
				}
				
				string upsName="";
                switch (GetUps(dr1["上游編號"].ToString()))
                {
                    case "2587":
                        ChkUps550.Checked = true;
                        //if (dr1["額定電流"].ToString() != "0") lblUps375I.Text = dr1["額定電流"].ToString();
						upsName="550N";
                        break;
                    case "83": 
                        ChkUps550.Checked = true;
                        //if (dr1["額定電流"].ToString() != "0") lblUps550I.Text = dr1["額定電流"].ToString();
                        //lblUps550S.Text = dr1["設備名稱"].ToString();
						upsName="550 ";
                        break;
                    case "84":
                        ChkUps550.Checked = true;
                        //if (dr1["額定電流"].ToString() != "0") lblUps600I.Text = dr1["額定電流"].ToString();
                        //lblUps600S.Text = dr1["設備名稱"].ToString();
						upsName="600 ";
                        break;
                    case "":                                              
                        break;
                    default:    //市電
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

    protected string GetUnit(string Pname, string which) //讀取課別或單位
    {
        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
        Conn.Open();
        SqlCommand cmd = new SqlCommand("SELECT [" + which + "] FROM [View_組織架構] WHERE [成員]='" + Pname + "'", Conn);
        SqlDataReader dr = null;

        string Unit = "";

        dr = cmd.ExecuteReader();
        if (dr.Read())  Unit = dr[0].ToString();

        cmd.Cancel(); cmd.Dispose(); dr.Close(); Conn.Close(); Conn.Dispose();        

        return (Unit);
    }

    protected string GetUpNo(string DevNo) //取得所連接迴路之編號
    {
        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
        Conn.Open();
        SqlCommand cmd = new SqlCommand("SELECT [上游編號] FROM [接電] WHERE [下游編號]=" + DevNo, Conn);
        SqlDataReader dr = null;

        string UpNo = "";

        dr = cmd.ExecuteReader();
        if (dr.Read()) UpNo=dr[0].ToString();

        cmd.Cancel(); cmd.Dispose(); dr.Close(); Conn.Close(); Conn.Dispose();

        return (UpNo);
    }

    protected string GetUps(string DevNo) //取得UPS編號
    {
        string UpsNo = "",tmpNo=DevNo;
        while(tmpNo != "-2")  //總電源編號
        {
            UpsNo = tmpNo;
            tmpNo = GetUpNo(tmpNo);
        }

        return (UpsNo);
    }
}