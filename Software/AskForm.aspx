<%@ Page Language="C#" AutoEventWireup="true" CodeFile="AskForm.aspx.cs" Inherits="SoftWare_AskForm" MaintainScrollPositionOnPostback="true" Debug="true" %>

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
            font-size: 14px;
            line-height: 18px;
        }
    </style>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <title>軟體使用申請單</title>
</head>
<body>
    <form id="form1" runat="server">
    <div align="center" style="margin: auto; width: 640px; font-size: 24px; font-family: 標楷體;font-weight: bold;">
        軟體使用申請單
    </div>
    <br />
    <div style="margin: auto; width: 640px; font-size: 14px; font-family: 標楷體;">
        <span style="float: right;">表單號碼：SMM02-<asp:Label ID="lblFormNo" runat="server" ForeColor="Gray" Font-Size="14px" Text="YYYY-NNNN"></asp:Label></span>
    </div>
    <table border="1" cellspacing="0" cellpadding="0">
        <tr style="padding: 3px">
            <td colspan="2" align="center" width="100px">單位名稱</td>
            <td colspan="2" width="250px">
                <asp:Label ID="lblAskUnit" runat="server" Font-Size="14px" Text="　" Font-Bold="True"></asp:Label>
            </td>

            <td align="center" width="80px">填表日期</td>
            <td>
                <asp:Label ID="lblKeyDay" runat="server" Font-Size="14px" Text="　　年　　月　　日" Font-Bold="true" />
            </td>
        </tr>
        
        <tr style="padding: 3px">
            <td colspan="2" align="center" width="100px">軟體名稱</td>
            <td colspan="2" width="250px">
                <asp:Label ID="lblSwName" runat="server" Font-Size="14px" Text="　" Font-Bold="True"></asp:Label>
            </td>

            <td align="center">版　本</td>
            <td>
                <asp:Label ID="lblVer" runat="server" Font-Size="14px" Text="　" Font-Bold="True"></asp:Label>
            </td>
        </tr>

        <tr style="padding: 3px">
            <td colspan="6">
                <asp:CheckBox ID="ChkIns" runat="server" Text="安裝" Font-Bold="true" /> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                <asp:CheckBox ID="ChkDel" runat="server" Text="移除" Font-Bold="true" />
            </td>
        </tr>

        <tr style="padding: 3px">
            <td colspan="2" align="center" width="100px">機器名稱</td>
            <td colspan="4">
                <asp:Label ID="lblHost" runat="server" Font-Size="14px" Text="　" Font-Bold="True"></asp:Label>
            </td>
        </tr>

        <tr style="padding: 3px">
            <td colspan="2" align="center" width="100px">IP</td>
            <td colspan="4">
                <asp:Label ID="lblIP" runat="server" Font-Size="14px" Text="　" Font-Bold="True"></asp:Label>
            </td>
        </tr>

        <tr style="padding: 3px">
            <td colspan="6">
                <table>
                    <tr style="font-weight:bold;text-align:left">
                        <td><asp:CheckBox ID="ChkNB" runat="server" /></td>
                        <td colspan="2">可攜式電腦使用微軟軟體</td>
                    </tr>
                    <tr>
                        <td style="vertical-align:top"></td>
                        <td style="vertical-align:top">※</td>
                        <td>依據微軟「Select 大量授權」，應用程式軟體第二份複本可安裝於主使用者的可攜式電腦(筆記型電腦)上使用。</td>
                    </tr>
                </table>
            </td>
        </tr>

        <tr style="padding: 3px">
            <td align="center" rowspan="2" width="20px">主使用</td>
            <td align="center" width="80px">機器名稱</td>
            <td width="250px">
                <asp:Label ID="lblSelHost" runat="server" Font-Size="14px" Text="　" Font-Bold="True"></asp:Label>
            </td>
            <td align="center" rowspan="2" width="20px">筆電</td>
            <td align="center" width="80px">機器名稱</td>
            <td width="250px">
                <asp:Label ID="lblNBHost" runat="server" Font-Size="14px" Text="　" Font-Bold="True"></asp:Label>
            </td>
        </tr>

        <tr style="padding: 3px">
            <td align="center" width="80px">IP</td>
            <td>
                <asp:Label ID="lblSelIP" runat="server" Font-Size="14px" Text="　" Font-Bold="True"></asp:Label>
            </td>
            <td align="center" width="80px">IP</td>
            <td>
                <asp:Label ID="lblNBIP" runat="server" Font-Size="14px" Text="　" Font-Bold="True"></asp:Label>
            </td>
        </tr>

        <tr style="padding: 3px">
            <td colspan="6">
                <table>
                    <tr style="font-weight:bold;text-align:left">
                        <td><asp:CheckBox ID="ChkTest" runat="server" /></td>
                        <td colspan="2">微軟授權軟體使用評估申請單</td>
                    </tr>
                    <tr>
                        <td style="vertical-align:top"></td>
                        <td style="vertical-align:top">※</td>
                        <td>依據Select合約，微軟提供60日軟體使用評估。評估期滿，使用者自行移除；未移除者應自負使用非授權軟體之責。</td>
                    </tr>
                </table>
            </td>
        </tr>

        <tr style="padding: 3px">
            <td colspan="2" align="center" width="80px">機器名稱</td>
            <td>
                <asp:Label ID="lblTestHost" runat="server" Font-Size="14px" Text="　" Font-Bold="True"></asp:Label>
            </td>
            <td rowspan="2" colspan="3" style="vertical-align:middle">
                使用日期：
                <asp:Label ID="lblUseDay1" runat="server" Font-Size="14px" Text="　年　月　日"></asp:Label>
                至
                <asp:Label ID="lblUseDay2" runat="server" Font-Size="14px" Text="　年　月　日"></asp:Label>
            </td>
        </tr>

        <tr style="padding: 3px">
            <td colspan="2" align="center" width="80px">IP</td>
            <td>
                <asp:Label ID="lblTestIP" runat="server" Font-Size="14px" Text="　" Font-Bold="True"></asp:Label>
            </td>
        </tr>

        <tr style="padding: 3px">
            <td colspan="6">
                <asp:CheckBox ID="ChkUpd" runat="server" Text="異動" Font-Bold="true" />
                &nbsp;&nbsp;&nbsp;&nbsp;※僅限在申請人所列管的機器之間異動
            </td>
        </tr>

        <tr style="padding: 3px">
            <td align="center" rowspan="2" width="20px">原安裝</td>
            <td align="center" width="80px">機器名稱</td>
            <td width="250px">
                <asp:Label ID="lblOldHost" runat="server" Font-Size="14px" Text="　" Font-Bold="True"></asp:Label>
            </td>
            <td align="center" rowspan="2" width="20px">新安裝</td>
            <td align="center" width="80px">機器名稱</td>
            <td width="250px">
                <asp:Label ID="lblNewHost" runat="server" Font-Size="14px" Text="　" Font-Bold="True"></asp:Label>
            </td>
        </tr>
        <tr style="padding: 3px">
            <td align="center" width="80px">IP</td>
            <td>
                <asp:Label ID="lblOldIP" runat="server" Font-Size="14px" Text="　" Font-Bold="True"></asp:Label>
            </td>
            <td align="center" width="80px">IP</td>
            <td>
                <asp:Label ID="lblNewIP" runat="server" Font-Size="14px" Text="　" Font-Bold="True"></asp:Label>
            </td>
        </tr>

        <tr>
            <td align="center" width="20px">申請人注意事項</td>
            <td colspan="5"> 
                <ol>
                    <li>依據『軟體管理手冊』辦理。</li>
                    <li>涉及軟體非法拷貝、使用、盜賣、循私營利或有違反智慧財產權之法律規定者須自行負責賠償及負所有法律責任。</li>
                    <li>使用軟體需填寫『軟體使用申請單』，經查合於授權規定及經核准後，始得借用。</li>
                    <li>申請核可後，憑此單向資管課借出軟體之媒體，僅能安裝於中心內部之電腦上。</li>
                    <li>借用期間，軟體之使用人應負保管之責任，如有使用或保管不當，造成毀損或遺失，應負賠償責任，但正常使用之毀損不在此限。</li>
                    <li>軟體分配使用後，如因人員離職、更換安裝機器或其它異動事項等，需即時填寫『軟體使用申請單』，辦理軟體異動登記。</li>
                    <li>軟體管理單位保有軟體使用調度權，申請者不得拒絕。</li>
                </ol>
            </td>
        </tr>
    </table>

    <br />
    <table border="1" cellspacing="0" cellpadding="0">
        <tr style="padding: 3px">
            <td width="25%" align="center"><b>申請科/課/室</b></td>
            <td width="25%" align="center"><b>承辦科/課/室</b></td>
            <td width="25%" align="center"><b>審核</b></td>
            <td width="25%" align="center"><b>核示</b></td>
        </tr>
        <tr style="padding: 3px">
            <td align="center">
                <asp:CheckBox ID="ChkAsker" runat="server" Text="申請人" />
                <asp:CheckBox ID="ChkAskMgr" runat="server" Text="申請主管" />
            </td>
            <td align="center"><asp:CheckBox ID="ChkJober" runat="server" Text="承辦人" /></td>
            <td align="center"><asp:CheckBox ID="ChkISer" runat="server" Text="資安人員" /></td>
            <td align="center"><asp:CheckBox ID="ChkJobMgr" runat="server" Text="承辦主管" /></td>
        <tr style="padding: 3px">
            <td>&nbsp;<br /><br /><br /><br /><br /></td>
            <td>&nbsp;<br /><br /><br /><br /><br /></td>
            <td>&nbsp;<br /><br /><br /><br /><br /></td>
            <td>&nbsp;<br /><br /><br /><br /><br /></td>
        </tr>
    </table>

    <div style="margin: auto; width: 640px; font-family: 標楷體; font-size: 14px;">
        　※本表單之正本請送<b>各權責單位軟體管理員</b>存查，並向<b>其</b>領用軟體‧
    </div>
    <br /><br /><br />

    <div style="margin: auto; width: 640px; font-family: 標楷體; font-size: 12px;">
        <font color="gray">表單資訊：CWB-ISMS-CWB-SMM-D002-軟體使用申請單_20170518_v1.1 &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; 機密等級：一般</font>
    </div>

    </form>
</body>
</html>
