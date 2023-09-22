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

public partial class Search_Search : System.Web.UI.Page
{
	public override void VerifyRenderingInServerForm(Control control) { }
	
    protected void Page_Load(object sender, EventArgs e)
    {
        if (!IsPostBack)
        {
              
        }
    }

    protected void Page_PreRenderComplete(object sender, EventArgs e)   //放Page_Load無法顯示
    {
		// TO CHECK!!
        //if (ViewState["SQL"] != null) SqlDataSource1.SelectCommand = ViewState["SQL"].ToString();
		//

        if (!IsPostBack)
        {
            LinkUPS_Click(null, null);
            
            if (Request["Search"] != null)  //外部關鍵字查詢
            {
                if (Request["Key"] != null)
                {
                    ChkKey.Checked = true;
                    //TextKey.Text = Request["Key"].ToString();
					TextKey.Text = HttpUtility.HtmlEncode(Request["Key"].ToString());					
                    // TO CHECK!!
					//BtnSearch_Click(null, null);
					//
                }
            }
        }
		
		// TO CHECK!!
		BtnSearch_Click(null, null);
		//
    }

    protected string GetPosSQL(string PosNo)
    {
        string SQL = "select [定位編號] from [定位設定]";

        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
        Conn.Open();
        SqlCommand cmd = new SqlCommand("select * from [定位設定] where [定位編號]=" + PosNo, Conn);
        SqlDataReader dr = null;
        dr = cmd.ExecuteReader();
        if (dr.Read()) SQL = SQL + " where [坐標X]>=" + dr["坐標X1"].ToString() + " and [坐標X]<=" + dr["坐標X2"].ToString()
            + " and [坐標Y]>=" + dr["坐標Y1"].ToString() + " and [坐標Y]<=" + dr["坐標Y2"].ToString();

        cmd.Cancel(); cmd.Dispose(); dr.Close(); Conn.Close(); Conn.Dispose();
        return (SQL);
    }
	
    protected string GetPosSQL(string PosNo, SqlCommand cmdSQL)
    {
        string SQL = "select [定位編號] from [定位設定]";

        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
        Conn.Open();
        //SqlCommand cmd = new SqlCommand("select * from [定位設定] where [定位編號]=" + PosNo, Conn);
		SqlCommand cmd = new SqlCommand("select * from [定位設定] where [定位編號]=@PosNo", Conn);
		cmd.Parameters.AddWithValue("@PosNo", PosNo);
		
        SqlDataReader dr = null;
        dr = cmd.ExecuteReader();
        //if (dr.Read()) SQL = SQL + " where [坐標X]>=" + dr["坐標X1"].ToString() + " and [坐標X]<=" + dr["坐標X2"].ToString()
        //    + " and [坐標Y]>=" + dr["坐標Y1"].ToString() + " and [坐標Y]<=" + dr["坐標Y2"].ToString();
		if (dr.Read()) SQL = SQL + " where [坐標X]>=@PosX1 and [坐標X]<=@PosX2"
            + " and [坐標Y]>=@PosY1 and [坐標Y]<=@PosY2";
	
		cmdSQL.Parameters.AddWithValue("@PosX1", dr["坐標X1"].ToString());
		cmdSQL.Parameters.AddWithValue("@PosX2", dr["坐標X2"].ToString());
		cmdSQL.Parameters.AddWithValue("@PosY1", dr["坐標Y1"].ToString());
		cmdSQL.Parameters.AddWithValue("@PosY2", dr["坐標Y2"].ToString());
	
        cmd.Cancel(); cmd.Dispose(); dr.Close(); Conn.Close(); Conn.Dispose();
        return (SQL);
    }	

    protected void BtnSearch_Click(object sender, EventArgs e)
    {
        if (SelNode.SelectedValue=="" & (ChkPower.Checked | ChkNet.Checked)) AddMsg("<script>alert('請先選取節點設備再查詢接電或接網迴路！');</script>");
        else
        {
            SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
            Conn.Open();
            SqlCommand cmd = new SqlCommand("", Conn);
            SqlDataReader dr = null;

            string SQL = "SELECT * FROM [View_通用設備] WHERE 1=1";
            //---------------------------------------------------------分頁
            if (ChkPage.Checked) GridView1.AllowPaging = true;
            else GridView1.AllowPaging = false;
            //---------------------------------------------------------設備
            // if (ChkType.Checked) SQL = SQL + " and [設備型態]='" + SelType.SelectedValue + "'";
			if (ChkType.Checked) {
				SQL = SQL + " and [設備型態]=@SelType";
				cmd.Parameters.AddWithValue("@SelType", SelType.SelectedValue);
			}
					
            // if (ChkKind.Checked) SQL = SQL + " and [設備種類]='" + SelKind.SelectedValue + "'";
			if (ChkKind.Checked) {
				SQL = SQL + " and [設備種類]=@SelKind";
				cmd.Parameters.AddWithValue("@SelKind", SelKind.SelectedValue);
			}
			
			// if (ChkDevStatus.Checked) SQL = SQL + " and [設備狀態]='" + SelDevStatus.SelectedValue + "'";
			if (ChkDevStatus.Checked) {
				SQL = SQL + " and [設備狀態]=@SelDevStatus";
				cmd.Parameters.AddWithValue("@SelDevStatus", SelDevStatus.SelectedValue);
			}
            
			// if (ChkPlace.Checked) SQL = SQL + " and [定位編號] in (select [定位編號] from [定位設定] where [區域名稱]='" + SelPlace.SelectedValue + "')";
			if (ChkPlace.Checked) {
				SQL = SQL + " and [定位編號] in (select [定位編號] from [定位設定] where [區域名稱]=@SelPlace)";
				cmd.Parameters.AddWithValue("@SelPlace", SelPlace.SelectedValue);
			}
			
            // if (ChkPointer.Checked) SQL = SQL + " and ([定位編號]=" + SelPointer.SelectedValue + "or [定位編號] in (" + GetPosSQL(SelPointer.SelectedValue) + "))";
			if (ChkPointer.Checked) {
				SQL = SQL + " and ([定位編號]=@SelPointer or [定位編號] in (" + GetPosSQL(SelPointer.SelectedValue, cmd) + "))";
				cmd.Parameters.AddWithValue("@SelPointer", SelPointer.SelectedValue);
			}
			
            if (ChkControl.Checked) SQL = SQL + " and [區域名稱] in (select [Item] from [Config] where [Kind]='門禁管制區')";
            //---------------------------------------------------------主機型態

            if (ChkHostType.Checked)
            {
                switch (SelHostType.SelectedValue)
                {
                    case "主機設備": SQL = SQL + " and ([設備型態]='系統設備' or [設備型態]='網路設備')"; break;
                    case "迴路設備": SQL = SQL + " and ([設備型態]='系統設備' or [設備型態]='網路設備' or [設備型態]='環境設備' or [設備型態]='週邊設備')"; break;
                    case "作業主機": SQL = SQL + " and [作業編號]>0"; break;
                    case "空機": SQL = SQL + " and ([設備型態]='系統設備' or [設備型態]='網路設備') and [作業編號] is null"; break;
                    case "系統主機": SQL = SQL + " and [作業編號]>0 and [設備型態]='系統設備'"; break;
                    case "網路主機": SQL = SQL + " and [作業編號]>0 and [設備型態]='網路設備'"; break;
                    case "中心設備": SQL = SQL + " and [設備維護人員] in (select [成員] from [View_組織架構] where [單位]='資訊中心')"; break;
                    case "代管設備": SQL = SQL + " and [設備維護人員] not in (select [成員] from [View_組織架構] where [單位]='資訊中心')"; break;
                    case "中心主機": SQL = SQL + " and [作業維護人員] in (select [成員] from [View_組織架構] where [單位]='資訊中心')"; break;
                    case "代管主機": SQL = SQL + " and [作業維護人員] not in (select [成員] from [View_組織架構] where [單位]='資訊中心')"; break;
                    case "個人主機": SQL = SQL + " and [系統名稱]='個人作業系統'"; break;
                    case "伺服主機": SQL = SQL + " and [系統名稱]<>'個人作業系統'"; break;
                    case "實體主機": SQL = SQL + " and [作業編號]>0 and [設備種類]<>'虛擬主機'"; break;
                    case "虛擬主機": SQL = SQL + " and [作業編號]>0 and [設備種類]='虛擬主機'"; break;
                    case "功能主機": SQL = SQL + " and [作業編號] in (select [作業編號] from [系統功能])"; break;
                    case "非功能主機": SQL = SQL + " and [作業編號] not in (select [作業編號] from [系統功能])"; break;
                    case "單一主機":  //非虛擬主機之同一設備編號之作業主機數等於1，且排除無主機
                        SQL = SQL + " and Exists(SELECT [設備編號],COUNT([設備編號]) FROM [作業主機] where [View_通用設備].[設備編號]=[設備編號]"
                            + " and [設備編號]>0 and [設備編號] not in (select [設備編號] from [實體設備] where [設備種類]='虛擬主機')"
                            + " group by [設備編號] having COUNT([設備編號])=1)";
                        break;
                    case "多重主機": //非虛擬主機之同一設備編號之作業主機數大於1，且排除無主機
                        SQL = SQL + " and Exists(SELECT [設備編號],COUNT([設備編號]) FROM [作業主機] where [View_通用設備].[設備編號]=[設備編號]"
                            + " and [設備編號]>0 and [設備編號] not in (select [設備編號] from [實體設備] where [設備種類]='虛擬主機')"
                            + " group by [設備編號] having COUNT([設備編號])>1)";
                        break;
                    default: break;
                }
            }
            //---------------------------------------------------------設備迴路
            if (ChkPower.Checked | ChkNet.Checked) 
            {
                string PowerNos = "", NetNos = "", DevNos = "";
                if (ChkPower.Checked)   //接電
                {
                    //PowerNos = GetValue("select dbo.RgetPowerChilds('" + SelNode.SelectedValue + "','N','" + SelNode.SelectedValue + "')");
					PowerNos = GetValue("select dbo.RgetPowerChilds(@SelNode,'N',@SelNode)", "@SelNode", SelNode.SelectedValue);
					
                    //if (ChkNet.Checked) NetNos = GetValue("select dbo.RgetNetChilds('" + PowerNos + "','N','" + PowerNos + "')");   //亦接網
					if (ChkNet.Checked) NetNos = GetValue("select dbo.RgetNetChilds(@PowerNos,'N',@PowerNos)", "@PowerNos", PowerNos);   //亦接網
                }
                //else NetNos = GetValue("select dbo.RgetNetChilds('" + SelNode.SelectedValue + "','N','" + SelNode.SelectedValue + "')");  //僅接網
				else NetNos = GetValue("select dbo.RgetNetChilds(@SelNode,'N',@SelNode)", "@SelNode", SelNode.SelectedValue);  //僅接網

                DevNos = FormatChilds(PowerNos + "," + NetNos);   //所有迴路下游
                if (ChkPower.Checked & SelPower.SelectedValue != "x" | ChkNet.Checked & SelNet.SelectedValue != "x") SQL = SQL + " and [設備編號] in (" + DevNos + ")";

                if (ChkPower.Checked)   //接電
                {
                    string UPS1="",UPS2="",UPS3="";
                    
                    switch (SelPower.SelectedValue)
                    {
                        case "x": SQL = SQL + " and [設備編號] not in (select [下游編號] from [接電])"; break;
                        case "v": break;
                        case "1": SQL = SQL + " and (select count(*) as Power  from [接電] where [下游編號]=[設備編號] group by [下游編號])=1"; break;
                        case "2": SQL = SQL + " and (select count(*) as Power  from [接電] where [下游編號]=[設備編號] group by [下游編號])=2"; break;
                        case "*": SQL = SQL + " and (select count(*) as Power  from [接電] where [下游編號]=[設備編號] group by [下游編號])>2"; break;
                        default:
                            {
                                UPS1 = "(" + GetValue("select dbo.RgetPowerChilds('2587','N','2587')") + ")"; //UPS375 -> 550N
                                UPS2 = "(" + GetValue("select dbo.RgetPowerChilds('83','N','83')") + ")"; //UPS550
                                UPS3 = "(" + GetValue("select dbo.RgetPowerChilds('84','N','84')") + ")"; //UPS600						
                                //EXCEPT
                                
                                switch (SelPower.SelectedValue)
                                {
                                    case "a": SQL = SQL + " and ([設備編號] in " + UPS1 + " and [設備編號] not in " + UPS2 + " and [設備編號] not in " + UPS3
                                        + " or [設備編號] not in " + UPS1 + " and [設備編號] in " + UPS2 + " and [設備編號] not in " + UPS3
                                        + " or [設備編號] not in " + UPS1 + " and [設備編號] not in " + UPS2 + " and [設備編號] in " + UPS3 + ")"; break;
                                    case "b": SQL = SQL + " and ([設備編號] in " + UPS1 + " and [設備編號] in " + UPS2 + " and [設備編號] not in " + UPS3
                                        + " or [設備編號] in " + UPS1 + " and [設備編號] not in " + UPS2 + " and [設備編號] in " + UPS3
                                        + " or [設備編號] not in " + UPS1 + " and [設備編號] in " + UPS2 + " and [設備編號] in " + UPS3 + ")"; break;
                                    case "c": SQL = SQL + " and ([設備編號] in " + UPS1 + " and [設備編號] in " + UPS2 + " and [設備編號] in " + UPS3 + ")"; break;
                                    case "?": SQL = SQL + " and (select count(*) as Power  from [接電] where [下游編號]=[設備編號] group by [下游編號])>1"
                                        + " and ([設備編號] in " + UPS1 + " and [設備編號] not in " + UPS2 + " and [設備編號] not in " + UPS3
                                        + " or [設備編號] not in " + UPS1 + " and [設備編號] in " + UPS2 + " and [設備編號] not in " + UPS3
                                        + " or [設備編號] not in " + UPS1 + " and [設備編號] not in " + UPS2 + " and [設備編號] in " + UPS3 + ")"; break;
                                }
                                break; 
                            }                        
                    }
                }

                if (ChkNet.Checked)   //接網
                {
                    switch (SelNet.SelectedValue)
                    {
                        case "x": SQL = SQL + " and [設備編號] not in (select [下游編號] from [接網])"; break;
                        case "v": break;
                        case "1": SQL = SQL + " and (select count(*) as Net  from [接網] where [下游編號]=[設備編號] group by [下游編號])=1"; break;
                        case "2": SQL = SQL + " and (select count(*) as Net  from [接網] where [下游編號]=[設備編號] group by [下游編號])=2"; break;
                        case "*": SQL = SQL + " and (select count(*) as Net  from [接網] where [下游編號]=[設備編號] group by [下游編號])>2"; break;
                    }
                }
            }

            //---------------------------------------------------------資產、財產、關機
            //if (ChkAssets.Checked) SQL = SQL + " and ([資產編號]='" + SelAssets.SelectedValue + "' or [系統編號] in (select [資源編號] from [系統資源] where [資產編號]='" + SelAssets.SelectedValue + "')"
            //    + "or [作業編號] in (select [作業編號] from [系統功能] where [資源編號] in (select [資源編號] from [系統資源] where [資產編號]='" + SelAssets.SelectedValue + "')))";
			if (ChkAssets.Checked) {
				SQL = SQL + " and ([資產編號]=@SelAssets or [系統編號] in (select [資源編號] from [系統資源] where [資產編號]=@SelAssets)"
					+ " or [作業編號] in (select [作業編號] from [系統功能] where [資源編號] in (select [資源編號] from [系統資源] where [資產編號]=@SelAssets)))";
				cmd.Parameters.AddWithValue("@SelAssets", SelAssets.SelectedValue);
			}
			
            if (ChkProp.Checked)
            {
                if (SelProp.SelectedValue == "有對應") SQL = SQL + " and [財產編號A]+' '+[財產編號B] in (select [財產編號A]+' '+[財產編號B] from [財產主檔])";
                else if (SelProp.SelectedValue == "不對應") SQL = SQL + "and not ([財產編號A]='' or [財產編號B]='' or [財產編號A]='00000000000') and [財產編號A]+' '+[財產編號B] not in (select [財產編號A]+' '+[財產編號B] from [財產主檔])";
                else if (SelProp.SelectedValue == "無財編") SQL = SQL + " and ([財產編號A]='' or [財產編號B]='' or [財產編號A]='00000000000')";
            }

            if (ChkOver.Checked) SQL = SQL + " and (DATEDIFF(dd,[購買日期],GETDATE())>365*[使用年限]/12 and [使用年限]<>0)";

            //if (ChkPowerOff.Checked) SQL = SQL + " and [關機順序] like '" + SelPowerOff.SelectedValue + "%'";
			if (ChkPowerOff.Checked) {
				SQL = SQL + " and [關機順序] like @SelPowerOff + '%'";
				cmd.Parameters.AddWithValue("@SelPowerOff", SelPowerOff.SelectedValue);				
			}
			
            //---------------------------------------------------------人員
            //string UnitSQL = "select distinct [成員] from [View_組織架構] where [課別]='" + SelUnit.SelectedValue + "'";			
			string UnitSQL = "select distinct [成員] from [View_組織架構] where [課別]=''";
			if (ChkUnit.Checked) {				
				UnitSQL = "select distinct [成員] from [View_組織架構] where [課別]=@SelUnit";
				cmd.Parameters.AddWithValue("@SelUnit", SelUnit.SelectedValue);
			}
					
			//
			string strSelManKindSelectedValue = "";
			switch (SelManKind.SelectedValue) {
				case "負責人員":
					strSelManKindSelectedValue = "負責人員"; break;
				case "設備維護人員":
					strSelManKindSelectedValue = "設備維護人員"; break;
                case "作業維護人員":
					strSelManKindSelectedValue = "作業維護人員"; break;
				case "保管人員":
					strSelManKindSelectedValue = "保管人員"; break;
                case "設備保管人員":
					strSelManKindSelectedValue = "設備保管人員"; break;
                case "財產保管人員":
					strSelManKindSelectedValue = "財產保管人員"; break;
                case "設備建立人員":
					strSelManKindSelectedValue = "設備建立人員"; break;
                case "設備修改人員":
					strSelManKindSelectedValue = "設備修改人員"; break;
                case "作業建立人員":
					strSelManKindSelectedValue = "作業建立人員"; break;
                case "作業修改人員":
					strSelManKindSelectedValue = "作業修改人員"; break;
				case "使用人員":
					strSelManKindSelectedValue = "使用人員"; break;
			}
				
            switch (SelManKind.SelectedValue)
            {
                case "負責人員":
                    {
                        if (ChkUnit.Checked) SQL = SQL + " and ([設備維護人員] in (" + UnitSQL + ")"
                            + " or [作業維護人員] in (" + UnitSQL + ")"
                            + " or [設備保管人員] in (" + UnitSQL + ")"
                            + " or [財產保管人員] in (" + UnitSQL + ")"
                            + ")";
                        //if (ChkGroup.Checked) SQL = SQL + " and ([設備維護人員]='" + SelGroup.SelectedValue + "' or [作業維護人員]='" + SelGroup.SelectedValue + "'"
                        //    + " or [設備保管人員]='" + SelGroup.SelectedValue + "' or [財產保管人員]='" + SelGroup.SelectedValue + "')";
						if (ChkGroup.Checked) {
							SQL = SQL + " and ([設備維護人員]=@SelGroup or [作業維護人員]=@SelGroup"
								+ " or [設備保管人員]=@SelGroup or [財產保管人員]=@SelGroup)";
							cmd.Parameters.AddWithValue("@SelGroup", SelGroup.SelectedValue);
						}
                        
						//if (ChkMan.Checked) SQL = SQL + " and ([設備維護人員]='" + SelMan.SelectedValue + "' or [設備維護人員] in (select Kind from Config where Item='" + SelMan.SelectedValue + "')"
                        //    + " or [作業維護人員]='" + SelMan.SelectedValue + "' or [作業維護人員] in (select Kind from Config where Item='" + SelMan.SelectedValue + "')"
                        //    + " or [設備保管人員]='" + SelMan.SelectedValue + "'" + " or [財產保管人員]='" + SelMan.SelectedValue + "'"
                        //    + ")";
						if (ChkMan.Checked) {
							SQL = SQL + " and ([設備維護人員]=@SelMan or [設備維護人員] in (select Kind from Config where Item=@SelMan)"
								+ " or [作業維護人員]=@SelMan or [作業維護人員] in (select Kind from Config where Item=@SelMan)"
								+ " or [設備保管人員]=@SelMan or [財產保管人員]=@SelMan"
								+ ")";									
							cmd.Parameters.AddWithValue("@SelMan", SelMan.SelectedValue);
						}							
                        break;
                    }
                case "設備維護人員":
                case "作業維護人員":
                    {
                        //if (ChkUnit.Checked) SQL = SQL + " and [" + SelManKind.SelectedValue + "] in (" + UnitSQL + ")";
						if (ChkUnit.Checked) SQL = SQL + " and [" + strSelManKindSelectedValue + "] in (" + UnitSQL + ")";
						
                        //if (ChkGroup.Checked) SQL = SQL + " and [" + SelManKind.SelectedValue + "]='" + SelGroup.SelectedValue + "'";
						if (ChkGroup.Checked) {
							SQL = SQL + " and [" + strSelManKindSelectedValue + "]=@SelGroup";
							cmd.Parameters.AddWithValue("@SelGroup", SelGroup.SelectedValue);
						}
						
                        //if (ChkMan.Checked) SQL = SQL + " and ([" + SelManKind.SelectedValue + "]='" + SelMan.SelectedValue + "' or [" + SelManKind.SelectedValue + "] in (select Kind from Config where Item='" + SelMan.SelectedValue + "'))";
						if (ChkMan.Checked) {
							SQL = SQL + " and ([" + strSelManKindSelectedValue + "]=@SelMan  or [" + strSelManKindSelectedValue + "] in (select Kind from Config where Item=@SelMan))";
							cmd.Parameters.AddWithValue("@SelMan", SelMan.SelectedValue);
						}
                        break;
                    }
                case "保管人員":
                case "設備保管人員":
                case "財產保管人員":
                case "設備建立人員":
                case "設備修改人員":
                case "作業建立人員":
                case "作業修改人員":
                    {
                        //if (ChkUnit.Checked) SQL = SQL + " and [" + SelManKind.SelectedValue + "] in (" + UnitSQL + ")";
						if (ChkUnit.Checked) SQL = SQL + " and [" + strSelManKindSelectedValue + "] in (" + UnitSQL + ")";
						
                        //if (ChkGroup.Checked) SQL = SQL + " and [" + SelManKind.SelectedValue + "]='" + SelGroup.SelectedValue + "'";
						if (ChkGroup.Checked) {
							SQL = SQL + " and [" + strSelManKindSelectedValue + "]=@SelGroup";
							cmd.Parameters.AddWithValue("@SelGroup", SelGroup.SelectedValue);
						}
						
                        //if (ChkMan.Checked) SQL = SQL + " and [" + SelManKind.SelectedValue + "]='" + SelMan.SelectedValue + "'";
						if (ChkMan.Checked) {
							SQL = SQL + " and [" + strSelManKindSelectedValue + "]=@SelMan";
							cmd.Parameters.AddWithValue("@SelMan", SelMan.SelectedValue);
						}
                        break;
                    }
                case "使用人員":
                    {
                        //if (ChkUnit.Checked) SQL = SQL + " and [使用課別]='" + SelUnit.SelectedValue + "'";
						if (ChkUnit.Checked) SQL = SQL + " and [使用課別]=@SelUnit";
						
                        //if (ChkGroup.Checked) SQL = SQL + " and [使用人員]='" + SelGroup.SelectedValue + "'";
						if (ChkGroup.Checked) {
							SQL = SQL + " and [使用人員]=@SelGroup";
							cmd.Parameters.AddWithValue("@SelGroup", SelGroup.SelectedValue);
						}
						
                        //if (ChkMan.Checked) SQL = SQL + " and ([使用人員]='" + SelMan.SelectedValue + "' or [使用人員] in (select [Kind] from [Config] where [Item]='" + SelMan.SelectedValue + "'))";
						if (ChkMan.Checked) {
							SQL = SQL + " and ([使用人員]=@SelMan or [使用人員] in (select [Kind] from [Config] where [Item]=@SelMan))";
							cmd.Parameters.AddWithValue("@SelMan", SelMan.SelectedValue);
						}
                        break;
                    }
            }
            //---------------------------------------------------------作業        
            //if (ChkSysNo.Checked) SQL = SQL + " and ([系統編號]=" + SelSysNo.SelectedValue + " or [作業編號] in (select [作業編號] from [系統功能] where [資源編號]=" + SelSysNo.SelectedValue + "))";
			if (ChkSysNo.Checked) {
				SQL = SQL + " and ([系統編號]=@SelSysNo or [作業編號] in (select [作業編號] from [系統功能] where [資源編號]=@SelSysNo))";
				cmd.Parameters.AddWithValue("@SelSysNo", SelSysNo.SelectedValue);
			}
			
            //if (ChkOsKind.Checked) SQL = SQL + " and [作業平台] in (Select Item from Config where Kind='作業平台' and Config='" + SelOsKind.SelectedValue + "')";
			if (ChkOsKind.Checked) {
				SQL = SQL + " and [作業平台] in (Select Item from Config where Kind='作業平台' and Config=@SelOsKind)";
				cmd.Parameters.AddWithValue("@SelOsKind", SelOsKind.SelectedValue);
			}			
			
            //if (ChkOS.Checked) SQL = SQL + " and [作業平台]='" + SelOS.SelectedValue + "'";
			if (ChkOS.Checked) {
				SQL = SQL + " and [作業平台]=@SelOS";
				cmd.Parameters.AddWithValue("@SelOS", SelOS.SelectedValue);
			}			
			
            //if (ChkSaveIO.Checked) SQL = SQL + " and [安內外]='" + SelSaveIO.SelectedValue + "'";
			if (ChkSaveIO.Checked) {
				SQL = SQL + " and [安內外]=@SelSaveIO";
				cmd.Parameters.AddWithValue("@SelSaveIO", SelSaveIO.SelectedValue);
			}
						
            //if (ChkOnCall.Checked) SQL = SQL + " and [緊急程度]='" + SelOnCall.SelectedValue + "'";
			if (ChkOnCall.Checked) {
				SQL = SQL + " and [緊急程度]=@SelOnCall";
				cmd.Parameters.AddWithValue("@SelOnCall", SelOnCall.SelectedValue);
			}
						
            //if (ChkApStatus.Checked) SQL = SQL + " and [作業狀態]='" + SelApStatus.SelectedValue + "'";
			if (ChkApStatus.Checked) {
				SQL = SQL + " and [作業狀態]=@SelApStatus";
				cmd.Parameters.AddWithValue("@SelApStatus", SelApStatus.SelectedValue);
			}
			
            //---------------------------------------------------------日期        
            if (ChkDate.Checked)
            {
                string strDate = DateTime.Now.AddDays(-int.Parse(SelDate.SelectedValue)).ToString("yyyy/MM/dd");
                if (SelDateKind.SelectedValue == "") SQL = SQL + " and ([設備建立日期]>='" + strDate + "' or [設備修改日期]>='" + strDate + "' or [作業建立日期]>='" + strDate
                               + "' or [作業修改日期]>='" + strDate + "' or [作業上線日期]>='" + strDate + "' or [購買日期]>='" + strDate + "')";
                //else SQL = SQL + " and [" + SelDateKind.SelectedValue + "]>='" + strDate + "'";
				else {
					string strSelDateKindSelectedValue = "";
					switch (SelDateKind.SelectedValue) {
						case "設備建立日期": 
							strSelDateKindSelectedValue = "設備建立日期"; break;
						case "設備修改日期": 							
							strSelDateKindSelectedValue = "設備修改日期"; break;
						case "作業建立日期": 
							strSelDateKindSelectedValue = "作業建立日期"; break;
						case "作業修改日期": 
							strSelDateKindSelectedValue = "作業修改日期"; break;
						case "作業上線日期": 
							strSelDateKindSelectedValue = "作業上線日期"; break;
						case "購買日期": 
							strSelDateKindSelectedValue = "購買日期"; break;
					}
					SQL = SQL + " and [" + strSelDateKindSelectedValue + "]>='" + strDate + "'";
				}
            }
            //---------------------------------------------------------SQL、關鍵字
            //if (ChkSQL.Checked & TextSQL.Text != "") SQL = SQL + " and (" + TextSQL.Text + ")";
			//FIXME!!

            if (ChkKey.Checked)
            {
                string[] KeyA = TextKey.Text.Trim().Split(',');

                for (int i = 0; i < KeyA.GetLength(0); i++)
                {				
                    // SQL = SQL + " and ([設備型態] like '%" + KeyA[i] + "%'"
                       // + " or [設備種類] like '%" + KeyA[i] + "%'"
                       // + " or [設備名稱] like '%" + KeyA[i] + "%'"
                       // + " or [設備用途] like '%" + KeyA[i] + "%'"
                       // + " or [財產編號A] like '%" + KeyA[i] + "%'"
                       // + " or [財產編號B] like '%" + KeyA[i] + "%'"
                       // + " or [財產名稱] like '%" + KeyA[i] + "%'"
                       // + " or [財產別名] like '%" + KeyA[i] + "%'"
                       // + " or [廠牌] like '%" + KeyA[i] + "%'"
                       // + " or [型式] like '%" + KeyA[i] + "%'"
                       // + " or [數量單位] like '%" + KeyA[i] + "%'"
                       // + " or [保管人員] like '%" + KeyA[i] + "%'"
                       // + " or [區域名稱]+[定位名稱] like '%" + KeyA[i] + "%'"
                       // + " or [財產保管人員] like '%" + KeyA[i] + "%'"
                       // + " or [取得來源] like '%" + KeyA[i] + "%'"
                       // + " or [規格] like '%" + KeyA[i] + "%'"
                       // + " or [硬體序號] like '%" + KeyA[i] + "%'"
                       // + " or [iLoIP] like '%" + KeyA[i] + "%'"
                       // + " or [維護廠商] like '%" + KeyA[i] + "%'"
                       // + " or [維護廠商] in (select [Item] from [Config] where [Kind]='維護廠商' and [Config] like '%" + KeyA[i] + "%')"
                       // + " or [設備維護人員] like '%" + KeyA[i] + "%'"
                       // + " or [設備維護人員] in (select [Kind] from [Config] where [Item] like '%" + KeyA[i] + "%')"
                       // + " or [設備狀態] like '%" + KeyA[i] + "%'"
                       // + " or [關機順序] like '%" + KeyA[i] + "%'"
                       // + " or [設備建立人員] like '%" + KeyA[i] + "%'"
                       // + " or [設備修改人員] like '%" + KeyA[i] + "%'"
                       // + " or [設備備註說明] like '%" + KeyA[i] + "%'"
                       // + " or [主機名稱] like '%" + KeyA[i] + "%'"
                       // + " or [系統編號] in (select [資源編號] from [View_系統資源] where [系統全名] like '%" + KeyA[i] + "%')"
                       // + " or [主要功能] like '%" + KeyA[i] + "%'"
                       // + " or [作業平台] like '%" + KeyA[i] + "%'"
                       // + " or [核心版本] like '%" + KeyA[i] + "%'"
                       // + " or [IP] like '%" + KeyA[i] + "%'"
                       // + " or [監控IP] like '%" + KeyA[i] + "%'"
                       // + " or [緊急程度] like '%" + KeyA[i] + "%'"
                       // + " or [作業狀態] like '%" + KeyA[i] + "%'"
                       // + " or [作業維護人員] like '%" + KeyA[i] + "%'"
                       // + " or [作業維護人員] in (select [Kind] from [Config] where [Item] like '%" + KeyA[i] + "%')"
                       // + " or [作業建立人員] like '%" + KeyA[i] + "%'"
                       // + " or [作業修改人員] like '%" + KeyA[i] + "%'"
                       // + " or [作業備註說明] like '%" + KeyA[i] + "%'"
                       // + " or [資產編號] like '%" + KeyA[i] + "%'"
                       // + ")";
					   
					   string par_key_name = "@KeyA" + Convert.ToString(i);
					   SQL = SQL + " and ([設備型態] like '%' + " + par_key_name + " + '%'"
						   + " or [設備種類] like '%' + " + par_key_name + " + '%'"
						   + " or [設備名稱] like '%' + " + par_key_name + " + '%'"
						   + " or [設備用途] like '%' + " + par_key_name + " + '%'"
						   + " or [財產編號A] like '%' + " + par_key_name + " + '%'"
						   + " or [財產編號B] like '%' + " + par_key_name + " + '%'"
						   + " or [財產名稱] like '%' + " + par_key_name + " + '%'"
						   + " or [財產別名] like '%' + " + par_key_name + " + '%'"
						   + " or [廠牌] like '%' + " + par_key_name + " + '%'"
						   + " or [型式] like '%' + " + par_key_name + " + '%'"
						   + " or [數量單位] like '%' + " + par_key_name + " + '%'"
						   + " or [保管人員] like '%' + " + par_key_name + " + '%'"
						   + " or [區域名稱]+[定位名稱] like '%' + " + par_key_name + " + '%'"
						   + " or [財產保管人員] like '%' + " + par_key_name + " + '%'"
						   + " or [取得來源] like '%' + " + par_key_name + " + '%'"
						   + " or [規格] like '%' + " + par_key_name + " + '%'"
						   + " or [硬體序號] like '%' + " + par_key_name + " + '%'"
						   + " or [iLoIP] like '%' + " + par_key_name + " + '%'"
						   + " or [維護廠商] like '%' + " + par_key_name + " + '%'"
						   + " or [維護廠商] in (select [Item] from [Config] where [Kind]='維護廠商' and [Config] like '%' + " + par_key_name + " + '%')"
						   + " or [設備維護人員] like '%' + " + par_key_name + " + '%'"
						   + " or [設備維護人員] in (select [Kind] from [Config] where [Item] like '%' + " + par_key_name + " + '%')"
						   + " or [設備狀態] like '%' + " + par_key_name + " + '%'"
						   + " or [關機順序] like '%' + " + par_key_name + " + '%'"
						   + " or [設備建立人員] like '%' + " + par_key_name + " + '%'"
						   + " or [設備修改人員] like '%' + " + par_key_name + " + '%'"
						   + " or [設備備註說明] like '%' + " + par_key_name + " + '%'"
						   + " or [主機名稱] like '%' + " + par_key_name + " + '%'"
						   + " or [系統編號] in (select [資源編號] from [View_系統資源] where [系統全名] like '%' + " + par_key_name + " + '%')"
						   + " or [主要功能] like '%' + " + par_key_name + " + '%'"
						   + " or [作業平台] like '%' + " + par_key_name + " + '%'"
						   + " or [核心版本] like '%' + " + par_key_name + " + '%'"
						   + " or [IP] like '%' + " + par_key_name + " + '%'"
						   + " or [監控IP] like '%' + " + par_key_name + " + '%'"
						   + " or [緊急程度] like '%' + " + par_key_name + " + '%'"
						   + " or [作業狀態] like '%' + " + par_key_name + " + '%'"
						   + " or [作業維護人員] like '%' + " + par_key_name + " + '%'"
						   + " or [作業維護人員] in (select [Kind] from [Config] where [Item] like '%' + " + par_key_name + " + '%')"
						   + " or [作業建立人員] like '%' + " + par_key_name + " + '%'"
						   + " or [作業修改人員] like '%' + " + par_key_name + " + '%'"
						   + " or [作業備註說明] like '%' + " + par_key_name + " + '%'"
						   + " or [資產編號] like '%' + " + par_key_name + " + '%'"
						   + ")";
					   cmd.Parameters.AddWithValue(par_key_name, KeyA[i]);					   
                }
            }
            //---------------------------------------------------------整合SQL之再加工
            int pos = SQL.IndexOf("FROM");
            //---------------------------------------------------------總筆數(統計起始)   
            cmd.CommandText = "SELECT count(*) " + SQL.Substring(pos);
            dr = cmd.ExecuteReader();
            TextCount.Text = "0";
            //try { if (dr.Read()) TextCount.Text = dr[0].ToString(); }
			try { if (dr.Read()) TextCount.Text = HttpUtility.HtmlEncode(dr[0].ToString()); }            
			catch { }
            dr.Close();
            //---------------------------------------------------------總價      
            //cmd.CommandText = "SELECT sum([價值]) from [View_設備管理] where [設備編號] in (select [設備編號] " + SQL.Substring(pos) + ")";
            //dr = cmd.ExecuteReader();
            //try { if (dr.Read()) TextTotal.Text = String.Format("{0:C0}", dr[0]); }
			//try { if (dr.Read()) TextTotal.Text = HttpUtility.HtmlEncode(String.Format("{0:C0}", dr[0])); }
            //catch { }
            //dr.Close();
            //---------------------------------------------------------總額定電流
            //cmd.CommandText = "SELECT sum([額定電流]) " + SQL.Substring(pos) + " and ([設備種類]<>'電源' and [設備種類]<>'PDC' and [設備種類]<>'配電盤' and [設備種類]<>'迴路')";
            //dr = cmd.ExecuteReader();
            //try { if (dr.Read()) TextSumCurrent.Text = String.Format("{0:F}", dr[0]); }
			//try { if (dr.Read()) TextSumCurrent.Text = HttpUtility.HtmlEncode(String.Format("{0:F}", dr[0])); }
            //catch { }

            cmd.Cancel(); cmd.Dispose(); //dr.Close();
			Conn.Close(); Conn.Dispose();
            //---------------------------------------------------------查詢顯示(統計結束)
            string ViewCols = SelShow.SelectedValue; //所有欲顯示欄位名稱
            SQL = "Select " + ViewCols + SQL.Substring(pos);
            //Test.Text = SQL;
            
            SqlDataSource1.SelectCommand = SQL;
						
			foreach (SqlParameter spar in cmd.Parameters) {
				string par_name = spar.ParameterName.Substring(spar.ParameterName.IndexOf("@") + 1);			
				if (SqlDataSource1.SelectParameters[par_name] != null) {
					SqlDataSource1.SelectParameters[par_name].DefaultValue = spar.Value.ToString();
				} else {
					Parameter par = new Parameter(par_name);
					par.DefaultValue = spar.Value.ToString();
					SqlDataSource1.SelectParameters.Add(par);
				}
			}
		
            GridView1.DataBind();
            ViewState["SQL"] = SQL;
            
            //Response.Write(SQL);    
        }
    }

    protected string FormatChilds(string PkNos)    //格式化下游字串
    {
        string cfg = PkNos.Replace(",0,", ",").Replace(",,", ",");
        if (cfg == "" | cfg == ",") cfg = "0";
        if (cfg.Substring(cfg.Length - 1) == ",") cfg = cfg.Substring(0, cfg.Length - 1);
        if (cfg.Substring(0, 1) == ",") cfg = cfg.Substring(1);

        return (cfg);
    }

    protected string GetValue(string SQL)   //取得單一資料
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
	
    protected string GetValue(string SQL, string par_name, object par_value)   //取得單一資料
    {
        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
        Conn.Open();
        SqlCommand cmd = new SqlCommand(SQL, Conn);
		//
		cmd.Parameters.AddWithValue(par_name, par_value);
		
        SqlDataReader dr = null;
        dr = cmd.ExecuteReader();
        string cfg = ""; if (dr.Read()) cfg = dr[0].ToString();
        cmd.Cancel(); cmd.Dispose(); dr.Close(); Conn.Close(); Conn.Dispose();

        return (cfg);
    }	

    protected DataSet RunQuery(SqlCommand sqlQuery) //讀取節點資訊
    {
        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
        SqlDataAdapter dbAdapter = new SqlDataAdapter();
        dbAdapter.SelectCommand = sqlQuery;
        sqlQuery.Connection = Conn;
        DataSet QueryDataSet = new DataSet();

        try
        {
            dbAdapter.Fill(QueryDataSet);
        }
        catch
        {
            Response.Write(sqlQuery.CommandText);
            Response.End();
        }
        dbAdapter.Dispose(); Conn.Close(); Conn.Dispose();
        return (QueryDataSet);   
    }

    protected void GridView1_RowCommand(object sender, CommandEventArgs e)
    {
        if (e.CommandName == "設備" | e.CommandName == "作業")
        {
            int Idx = int.Parse(e.CommandArgument.ToString());
            string DevNo = GridView1.DataKeys[Idx].Value.ToString(); if (DevNo == "") DevNo = "0";
            string ApNo = GridView1.Rows[Idx].Cells[3].Text; //if (ApNo == ""| ApNo == "&nbsp;") ApNo = "0"; 
            int n; if (!int.TryParse(ApNo, out n)) ApNo = "0";

            switch (e.CommandName)
            {
                case "設備":
                    {
                        //if (DevNo != "0") OpenExecWindow("window.open('../Device/DevEdit.aspx?DevNo=" + DevNo + "','_blank');");
						if (DevNo != "0") OpenExecWindow("window.open('../Device/DevEdit.aspx?DevNo=" + HttpUtility.UrlEncode(DevNo) + "','_blank');");
                        break;
                    }
                case "作業":
                    {						
						//if (ApNo != "0") OpenExecWindow("window.open('../AP/ApEdit.aspx?ApNo=" + ApNo + "','_blank');");
                        if (ApNo != "0") OpenExecWindow("window.open('../AP/ApEdit.aspx?ApNo=" + HttpUtility.UrlEncode(ApNo) + "','_blank');");
                        else AddMsg("<script>alert('顯示欄位第二欄是 [作業編號] 且有值方能開啟作業編輯介面！');</script>");
                        break;
                    }
            }
        }
    }

    protected void OpenExecWindow(string strJavascript)    //選取後就另開視窗
    {
        if (ScriptManager.GetCurrent(Page) == null)
        {
            Page.ClientScript.RegisterStartupScript(Page.GetType(), "LifeLog", strJavascript, true);
        }
        else
        {
            ScriptManager.RegisterStartupScript(Page, Page.GetType(), "LifeLog", strJavascript, true);
        }
    }

    protected void SelSQL_SelectedIndexChanged(object sender, EventArgs e)
    {
        //TextSQL.Text = SelSQL.SelectedValue;
		TextSQL.Text = HttpUtility.HtmlEncode(SelSQL.SelectedValue);		
    }     

    protected void BtnHtml_Click(object sender, EventArgs e)
    {
        if (GridView1.Rows.Count > 0)
        {
            int pos = ViewState["SQL"].ToString().IndexOf("FROM");
            string ViewCols = SelShow.SelectedValue; //所有欲顯示欄位名稱

            GridView gvExport = new GridView();
            gvExport.DataSource = SqlDataSource1;
            SqlDataSource1.SelectCommand = "Select " + ViewCols + ViewState["SQL"].ToString().Substring(pos);
            gvExport.DataBind();

            string strExportFilename = "IDMS";

            Response.Clear();
            Response.ClearContent();
            Response.Write("<meta http-equiv=Content-Type content=text/html;charset=utf-8>");            
			Response.AddHeader("content-disposition", "attachment;filename=" + strExportFilename + ".html");
            Response.Cache.SetCacheability(HttpCacheability.NoCache);
            Response.ContentType = "application/vnd.xls";
			
            Response.Charset = "utf-8";

            System.IO.StringWriter stringWrite = new System.IO.StringWriter();
            System.Web.UI.HtmlTextWriter htmlWrite = new HtmlTextWriter(stringWrite);
            gvExport.RenderControl(htmlWrite);
            Response.Write(stringWrite.ToString().Replace("<div>", "").Replace("</div>", ""));
            Response.End();
        }
    }
	
    protected void BtnExcel_Click(object sender, EventArgs e)
    {
        if (GridView1.Rows.Count > 0)
        {		
            int pos = ViewState["SQL"].ToString().IndexOf("FROM");
            string ViewCols = SelShow.SelectedValue; //所有欲顯示欄位名稱

            GridView gvExport = new GridView();
            gvExport.DataSource = SqlDataSource1;
            SqlDataSource1.SelectCommand = "Select " + ViewCols + ViewState["SQL"].ToString().Substring(pos);
			gvExport.AllowPaging = false;
			gvExport.DataBind();

            Response.Clear();
            Response.ClearContent();
			Response.AddHeader("content-disposition", "attachment;filename=IDMS.csv");
            Response.Cache.SetCacheability(HttpCacheability.NoCache);
			Response.ContentType = "application/text";			            
			Response.Charset = "BIG5";
			Response.ContentEncoding = System.Text.Encoding.GetEncoding(950);
						
			StringBuilder sb = new StringBuilder();			
			foreach (var c in gvExport.HeaderRow.Cells)
            {
				sb.Append('"' + (((System.Web.UI.WebControls.DataControlFieldCell)(c)).ContainingField).HeaderText + '"' + ",");				
            }

			sb.Append("\r\n");
			for (int i = 0; i < gvExport.Rows.Count; i++)
			{
				for (int k = 0; k < gvExport.Rows[i].Cells.Count; k++)
				{
					sb.Append('"' + gvExport.Rows[i].Cells[k].Text.Replace("&nbsp;", "") + '"' + ',');
				}

				sb.Append("\r\n");
			}

			Response.Output.Write(sb.ToString());
			Response.Flush();
			Response.End();
        }
    }

    protected void AddMsg(string strMsg)
    {
        Literal Msg = new Literal();
        Msg.Text = strMsg;
        Page.Controls.Add(Msg);
    }

    protected void BtnMap_Click(object sender, EventArgs e) //設備分佈圖
    {
        Session["DevSQL"] = ViewState["SQL"];
        OpenExecWindow("window.open('../Lib/map.aspx','_blank','location=no;menubar=no;resizable=yes;scrollbars=no;status=no;toolbar=no;fullscreen=yes');");
    }

    protected void LinkHostType_Click(object sender, EventArgs e)  //帶入
    {
        OpenExecWindow("window.open('../Help/主機型態.txt','_blank');");
    }

    protected void LinkHostsSplit_Click(object sender, EventArgs e)  //帶入
    {
        //TextSQL.Text = "[主機名稱] in ('" + TextSQL.Text.Replace(",", "','") + "')";
		TextSQL.Text = "[主機名稱] in ('" + HttpUtility.HtmlEncode(TextSQL.Text).Replace(",", "','") + "')";		
    }

    protected void LinkUPS_Click(object sender, EventArgs e)  //列出UPS電源，以便查詢下游接電、接網迴路
    {
        SqlDataSourceNode.SelectCommand = "SELECT [設備編號],[設備名稱] FROM [View_設備管理] where [設備種類] in ('配電盤','PDC','電源','連接埠') or [定位名稱]='連接埠' ORDER BY [設備種類] desc,[設備名稱]";
        SelNode.DataBind();
    }
    protected void LinkPointer_Click(object sender, EventArgs e)  //依放置地點列出設備清單，以便查詢下游接電、接網迴路
    {
        //SqlDataSourceNode.SelectCommand = "SELECT [設備編號],[設備名稱] FROM [View_設備管理] where [定位編號]=" + SelPointer.SelectedValue + " ORDER BY [高度] desc";
		SqlDataSourceNode.SelectCommand = "SELECT [設備編號],[設備名稱] FROM [View_設備管理] where [定位編號]=" + "@SelPointer" + " ORDER BY [高度] desc";
				
		string par_name = "SelPointer";			
		if (SqlDataSourceNode.SelectParameters[par_name] != null) {
			SqlDataSourceNode.SelectParameters[par_name].DefaultValue = SelPointer.SelectedValue;
		} else {
			Parameter par = new Parameter(par_name);
			par.DefaultValue = SelPointer.SelectedValue;
			SqlDataSourceNode.SelectParameters.Add(par);
		}
				
        SelNode.DataBind();
    }
}