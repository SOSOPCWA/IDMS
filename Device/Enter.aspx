<%@ Page Language="C#" AutoEventWireup="true" CodeFile="Enter.aspx.cs" Inherits="Device_Enter" MaintainScrollPositionOnPostback="true" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
    <style type="text/css">
        table
        {
            width: 640px;
            margin: auto;
        }
        
        td
        {
            font-family: 標楷體;
            font-size: 16px;
            line-height: 18px;
        }
    </style>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <title>氣象數值資訊組硬體設備移入機房申請單</title>
</head>
<body>
    <form id="form1" runat="server">
    <br />
    <div align="center" style="margin: auto; width: 640px; font-size: 16px; font-family: 標楷體;font-weight: bold;">
        氣象數值資訊組<br />硬體設備移入機房申請單
    </div>

    <br />

    <div style="margin: auto; width: 640px; font-size: 16px; font-family: 標楷體;">
        <span style="float: left;">申請日期：       &nbsp&nbsp&nbsp 年 &nbsp 月  &nbsp日</span> 
        <span style="float: right;">表單編號：SOS02-<asp:Label ID="lblFormNo" runat="server" Font-Size="16px" Text="yyyy-nnnn" ForeColor="#CCCCCC"></asp:Label></span>
    </div>
    <table border="1" cellspacing="0" cellpadding="0">
        <tr style="padding: 3px">
            <td width="16%">申請單位</td>
            <td width="16%" align="center"><asp:Label ID="lblAskUnit" runat="server" Font-Size="16px" Text="　" Font-Bold="True"></asp:Label></td>
            <td width="16%">申請人</td>
            <td width="14%" align="center"><asp:Label ID="lblAsker" runat="server" Font-Size="16px" Text="　" Font-Bold="True"></asp:Label></td>
            <td width="20%" align="center">上架/異動日期</td>
            <td width="18%" align="center"><asp:Label ID="lblEnterDate" runat="server" Font-Size="16px" Text="　年　月　日" Font-Bold="True"></asp:Label></td>
        </td>
        </tr>
		
		<tr style="padding: 3px">
            <td>設備種類</td>
            <td colspan="5">
                <asp:CheckBox ID="ChkSys" runat="server" />系統設備(主機)
                <asp:CheckBox ID="ChkNet" runat="server" />網路設備
                <asp:CheckBox ID="ChkST" runat="server" />儲存設備
                <asp:CheckBox ID="ChkEnv" runat="server" />環境設備
                <asp:CheckBox ID="ChkElse" runat="server" />其他
            </td>
        </tr>
		
		<tr style="padding: 3px">
            <td rowspan=4 >設備資訊</td>
            <td colspan="5"><asp:Label ID="SysName" runat="server" Font-Size="16px" Text="系統名稱:　" ></asp:Label></td>
        </tr>
		
		 <tr style="padding: 3px">
            <td colspan="5"><asp:Label ID="lblPurpose" runat="server" Font-Size="16px" Text="設備用途:　" ></asp:Label></td>
        </tr>
		
		<tr style="padding: 3px">
            <td colspan="3"><asp:Label ID="lblDevName" runat="server" Font-Size="16px" Text="設備名稱:　" ></asp:Label></td>
			<td colspan="2"><asp:Label ID="lblIP" runat="server" Font-Size="16px" Text="IP: 　" ></asp:Label></td>
            
		</tr>
		
        <tr style="padding: 3px">
            <td colspan="3"><asp:Label ID="lblHw" runat="server" Font-Size="16px" Text="硬體負責人:　" ></asp:Label></td>
            <td colspan="2"><asp:Label ID="lblSw" runat="server" Font-Size="16px" Text="軟體負責人:　" ></asp:Label></td>
        
        </tr>

       

        <tr style="padding: 3px">
            <td>移入理由</td>
            <td colspan="5">
                <asp:CheckBox ID="ChkAP" runat="server" Font-Bold="False" />本中心業務即時作業及監控 &nbsp;&nbsp;
                <asp:CheckBox ID="ChkReplace" runat="server" />預備替換相關機器&nbsp;&nbsp;
				<asp:CheckBox ID="Chkbku" runat="server" />異地備援
				<br />
                <asp:CheckBox ID="ChkAgent" runat="server" />代管機器 &nbsp;&nbsp;
                <asp:CheckBox ID="ChkTest" runat="server" />測試機器 &nbsp;&nbsp;
                <asp:CheckBox ID="ChkOther" runat="server" />其他
            </td>
        </tr>

 

        <tr style="padding: 3px">
            <td rowspan=3 >放置位置</td>
			<td colspan="5"><asp:Label ID="lblPlaceName" runat="server" Font-Size="16px" Text="地點:&nbsp&nbsp&nbsp&nbsp機房　" ></asp:Label></td>
		</tr>	
		
		
		<tr>
            <td colspan="5">
                <asp:CheckBox ID="ChkRack" runat="server" />機架式
                
                <asp:Label ID="lblPlace"  runat="server" Font-Size="16px" Text="　機櫃編號:　　　&nbsp&nbsp&nbsp&nbsp&nbsp　　" ></asp:Label>，
                <asp:Label ID="lblHeight" runat="server" Font-Size="16px" Text="高度:&nbsp;&nbsp;U 　- &nbsp;&nbsp;U 　&nbsp;" ></asp:Label>
            </td>
        </tr>
		
		<tr>
			<td colspan="5">
				<asp:CheckBox ID="ChkXY" runat="server" />非機架式&nbsp;  座標(&nbsp;&nbsp;,&nbsp;&nbsp;) 或其他說明:
			</td>
		</tr>

        <tr style="padding: 3px">
            <td>用電需求</td>
            <td colspan="5">
                電壓(V):&nbsp;<asp:Label ID="lblVoltage" runat="server" Font-Size="16px" Text="　　" Font-Bold="True" Font-Underline="True"></asp:Label>伏特，
                使用電流(I):&nbsp;<asp:Label ID="lblCurrent" runat="server" Font-Size="16px" Text="　　" Font-Bold="True" Font-Underline="True"></asp:Label>安培，
                用電量/負載:&nbsp;<asp:Label ID="lblKVA" runat="server" Font-Size="16px" Text="　　" Font-Bold="True" Font-Underline="True"></asp:Label>&nbsp;KVA
            </td>
        </tr>

        <tr style="padding: 3px">
            <td rowspan=2>電源資訊</td>
            <td colspan="5">
                
                <asp:CheckBox ID="ChkUpsNo" runat="server" />使用市電(未接發電機) &nbsp;
				<asp:CheckBox ID="ChkUpsNoCon" runat="server" />使用市電(接發電機) &nbsp;&nbsp;
				配電盤:
            </td>
        </tr>
		
		<tr style="padding: 3px">
            <td colspan="5">
				<asp:CheckBox ID="ChkUps550" runat="server" />使用不斷電系統(UPS)
                <br />
                
                UPS:<asp:Label ID="lblUps1" runat="server" Font-Size="16px" Text="　　&nbsp;"  Font-Underline="True"></asp:Label>
                ，插座編號：<asp:Label ID="lblUps1S" runat="server" Font-Size="16px" Text="　　　　　　&nbsp;"  Font-Underline="True"></asp:Label>，
				設計容量:<asp:Label ID="lblUps1I" runat="server" Font-Size="16px" Text="　　"  Font-Underline="True"></asp:Label>安培<br />
                
                
                UPS:<asp:Label ID="lblUps2" runat="server" Font-Size="16px" Text="　　&nbsp;" Font-Underline="True"></asp:Label>
                ，插座編號：<asp:Label ID="lblUps2S" runat="server" Font-Size="16px" Text="　　　　　　&nbsp;" Font-Underline="True"></asp:Label>，
				設計容量:<asp:Label ID="lblUps2I" runat="server" Font-Size="16px" Text="　　"  Font-Underline="True"></asp:Label>安培
				<br />
             </td>
            
        </tr>
		
		<tr style="padding: 3px">
            <td>熱量排放</td>
            <td colspan="5">
                預估設備排放熱量:&nbsp;<asp:Label ID="lblBTU" runat="server" Font-Size="16px" Text="　　　" Font-Bold="True" Font-Underline="True"></asp:Label>&nbsp;BTU/H
            </td>
        </tr>

        
    </table>

    <br />

    <div style="width: 640px; font-family: 標楷體; font-size: 16px; margin: auto;">機房管理單位填寫</div>
    <table border="1" cellspacing="0" cellpadding="0">
        <tr style="padding: 3px">
            <td width="16%">管理單位:<br/>作業管理科</td>
            <td width="84%" align="center">審核日期:______年___月___日</td>
        </tr>

        <tr style="padding: 3px">
            <td>UPS使用量安全評估</td>
            <td>
                目前UPS使用量<br />
                UPS ：<u>　　　</u>  使用量:<u>　　　</u>KVA/(<u>　　　</u>%) <br/>
                UPS ：<u>　　　</u>  使用量:<u>　　　</u>KVA/(<u>　　　</u>%)
                
            </td>
        </tr>

        <tr style="padding: 3px">
            <td>溫度評估</td>
            <td>________機房 ________區 &nbsp;&nbsp;&nbsp; 目前溫度:________℃ <u></u></td>
        </tr>
        <tr style="padding: 3px">
            <td>移入機房綜合評估/意見</td>
            <td>
                加此設備後UPS使用量安全評估：
                <input type="checkbox" />是
                <input type="checkbox" />否 為安全使用範圍內
                <br />
                溫度及空調需求評估:
                <input type="checkbox" />是
                <input type="checkbox" />否 為安全使用範圍內
                <br />
                <input type="checkbox" />同意
                <input type="checkbox" />不同意
                <input type="checkbox" />其他/建議 _______________________
            </td>
        </tr>
    </table>

    <br />

    <table border="1" cellspacing="0" cellpadding="0">
        <tr style="padding: 3px">
            <td align="center">
                申請單位 <br />
                <span style="font-size:12px">
                    <input type="checkbox" />承辦人 <input type="checkbox" />資安人員 <input type="checkbox" />承辦主管 <br /> 
                    <input type="checkbox" />單位主管(申請與管理同單位者免勾選)
                </span>
            </td>
            <td align="center">
                機房管理單位 <br />
                    <span style="font-size:12px">
                    <input type="checkbox" />承辦人 <input type="checkbox" />資安人員 <input type="checkbox" />承辦主管
                </span>
            </td>
            <td align="center">
                核示 <br />
                <span style="font-size:12px">
                    <input type="checkbox" />單位資安長 <input type="checkbox" />單位主管 <br /> 
                </span>
            </td>
        </tr>

        <tr style="padding: 3px">
            <td height="120">&nbsp;</td>
            <td height="120">&nbsp;</td>
            <td height="120">&nbsp;</td>
        </tr>
    </table>

    <div style="width: 640px; font-family: 標楷體; font-size: 16px; margin: auto;">
        附註:本表單簽核完成後由機房管理權責單位留存，保存期限1年。
    </div>
    
    <div style="width: 640px; font-family: 標楷體; font-size: 12px; margin: auto;color:#CCCCCC">
        表單資訊：CWB-ISMS-MIC-SOS-D002-硬體設備移入機房申請單_1080906_v1.2 &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; 機密等級：限閱
    </div>



    <p style='page-break-after: always'></p>
    <p>&nbsp;</p>
    <table>
        <tr>
            <td>
                <table border="0" cellspacing="0" cellpadding="0">
                    
                    <tr>
                        <td style="font-size: 24px; line-height: 32px;">填寫說明</td>
                    </tr>
                    <tr>
                        <td style="font-size: 16px; line-height: 20px;">
                            <ol>
                                <li>設備名稱、設備用途、硬體負責人必填。</li>
                                <li>異地備援：移入之設備其用途為異地備援作業使用，請勾選。</li>
                                <li>放置位置：設備放置機櫃內，請填機櫃編號、第幾U-第幾U。 其他則填位置座標(X.Y)</li>
                                <li>用電需求：請檢視設備規格標貼(W、A、V)，或詢問廠商。
								<br/>公式參考：V=IR(單相)，V=IR×√3 (三相)，W=VI，KW=W/1000=kva，<br/>請儘量評估實電流，
								可不考慮功率因數PF（與電力耗損計算相關）。</li>
                                <li>UPS：請依各單位機房之不斷電系統設備配置填寫。
								<br/>例如氣象數值資訊組之UPS有550UPS、550UPSN、600UPS。</li>
                             
                                <li>熱量排放：請洽設備廠商，或自行查閱電源供應器說明。
								<br/>BTU記算式：BTU = 伏特數 x 安培數 x 3.41。</li>
                                <li>UPS使用量安全評估：根據廠商建議，使用率一般設限於80%以內，其餘 20% 為預備所有設備瞬間啟動時之電力供應。</li>
                            </ol>
                        </td>
                    </tr>
					<tr>
                        <td align="center" style="font-size: 24px; line-height: 64px;">
                            IDMS編號：<asp:Label ID="lblDevNo" runat="server" Font-Size="24px" Text="　" Font-Bold="True"></asp:Label>
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
        <tr><td style="font-size: 16px; line-height: 500px;">&nbsp;</td></tr>
        <tr>
            <td>
                <div style="width: 640px; font-family: 標楷體; font-size: 12px; margin: auto;color:#CCCCCC">
                    表單資訊：CWB-ISMS-MIC-SOS-D002-硬體設備移入機房申請單_20171221_v1.1 &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; 機密等級：限閱
                </div>
            </td>
        </tr>
    </table>
    </form>
</body>
</html>
