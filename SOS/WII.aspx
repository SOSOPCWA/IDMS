<%@ Page Title="局屬網路" Language="C#" MasterPageFile="~/IDMS/MasterPage.master" AutoEventWireup="true"
    MaintainScrollPositionOnPostback="true" CodeFile="WII.aspx.cs" Inherits="SOS_WII" %>

<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="Server">
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1" runat="Server">    
    <table width="100%" background="~/IDMS/images/bkgd.gif" bgcolor="#DDEEFF">
        <tr>
            <td>
                <asp:ImageMap runat="server" ID="ImgWII" ImageUrl="~/IDMS/images/taiwan.gif">
                </asp:ImageMap>
            </td>
            <td valign="top">
                <B><FONT color="blue" size="5"><asp:Label runat="server" ID="lblHost" Font-Bold="true" Font-Size="X-Large" ForeColor="Green"></asp:Label></FONT></B> <br>
                <font color="#990000" size="4">===================== 報修資訊 ========================</font>
                <table>
                    <tr>
                        <td>
                            <font color="#660000">安內主線</font>
                        </td>
                        <td>
                            <font color="green">：<asp:Label runat="server" ID="lblPno"></asp:Label></font>
                        </td>
                    </tr>
                    <tr>
                        <td>
                            <font color="#660000">安內備線</font>
                        </td>
                        <td>
                            <font color="green">：<asp:Label runat="server" ID="lblBno"></asp:Label></font>
                        </td>
                    </tr>
                    <tr>
                        <td>
                            <font color="#660000">安內主IP</font>
                        </td>
                        <td>
                            <font color="#CC0066">：<asp:Label runat="server" ID="lblIP"></asp:Label></font>
                        </td>
                    </tr>
                    <tr>
                        <td>&nbsp;</td>
                        <td>&nbsp;</td>
                    </tr>
                    <tr>
                        <td>
                            <font color="#660000">安外線路</font>
                        </td>
                        <td>
                            <font color="green">：<asp:Label runat="server" ID="lblOno"></asp:Label></font>
                        </td>
                    </tr>
                    <tr>
                        <td>&nbsp;</td>
                        <td>&nbsp;</td>
                    </tr>
                    <tr>
                        <td>
                            <font color="#660000">聯絡資訊</font>
                        </td>
                        <td>
                            <font color="#CC0066">：<asp:Label runat="server" ID="lblTel"></asp:Label></font>
                        </td>
                    </tr>   
                    <tr>
                        <td>
                            <font color="#660000">測站維護</font>
                        </td>
                        <td>
                            ：<asp:HyperLink ID="lnHost" runat="server" Target="_blank" Text="請按此"></asp:HyperLink>
                        </td>
                    </tr>                 
                    
                    <tr>
                        <td>&nbsp;</td>
                        <td>&nbsp;</td>
                    </tr>       
                    <tr>
                        <td>
                            <font color="#660000">報修電話</font>
                        </td>
                        <td>
                            <font color="#CC0066">：0800-080-123、2344-3118(HiLink機房)</font> <br />
                            <font color="#CC0066">：7738-8066(GSN-VPN)、2344-3007(機房)</font> <br />
                            <font color="#CC0066">：WiFi電路通知當地CHT</font>
                        </td>
                    </tr>

                    <tr>
                        <td>&nbsp;</td>
                        <td>&nbsp;</td>
                    </tr> 
                    <tr>
                        <td>
                            <font color="#660000">關鍵字</font>
                        </td>
                        <td>
                            <font color="#660000">：</font>
                            <asp:TextBox runat="server" ID="txtKey"></asp:TextBox> &nbsp;&nbsp;
                            <asp:Button runat="server" ID="BtnKey" Text="查詢" OnClick="BtnKey_Click" ToolTip="僅顯示第一筆" /> &nbsp;&nbsp;
                            <asp:Label runat="server" ID="lblKey" ForeColor="green"></asp:Label>
                        </td>
                    </tr> 
                </table>
                
                <br/><br/>
                可以從線路編號為 ??AT112???判斷是否為中心負責
                <br/><br/> 
            </td>
        </tr>
    </table>       
</asp:Content>
