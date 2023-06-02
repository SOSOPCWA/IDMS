<%@ Page Title="設備管理" Language="C#" MasterPageFile="~/IDMS/MasterPage.master" AutoEventWireup="true" CodeFile="Device.aspx.cs" Inherits="Device_Device" MaintainScrollPositionOnPostback="true" %>

<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="Server">
    <style type="text/css">
        .style4
        {
            width: 100%;
        }
        .style5
        {
            text-align: center;
        }
        .style6
        {
            font-size: Small;
            color: #006600;
        }
    </style>
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1" runat="Server">
    <table border="1" class="style4">
        <tr>
            <td width="240" valign="top">
                <asp:DropDownList ID="SelNode1" runat="server" AutoPostBack="True" ForeColor="#006600">
                    <asp:ListItem Selected="True">設備種類</asp:ListItem>
                    <asp:ListItem>區域名稱</asp:ListItem> 
                    <asp:ListItem>定位名稱</asp:ListItem>
                    <asp:ListItem>放置地點</asp:ListItem> 
                    <asp:ListItem>維護課別</asp:ListItem>
                    <asp:ListItem>維護人員</asp:ListItem>
                    <asp:ListItem>保管課別</asp:ListItem>
                    <asp:ListItem>保管人員</asp:ListItem>
                    <asp:ListItem>維護廠商</asp:ListItem>
                    <asp:ListItem>廠牌</asp:ListItem>
                    <asp:ListItem>廠牌型式</asp:ListItem>
                    <asp:ListItem>設備狀態</asp:ListItem>
                    <asp:ListItem>用電電壓</asp:ListItem>
                    <asp:ListItem>取得來源</asp:ListItem>
                    <asp:ListItem>關機順序</asp:ListItem>
                    <asp:ListItem>建立人員</asp:ListItem>
                    <asp:ListItem>修改人員</asp:ListItem>
                </asp:DropDownList>
                <asp:DropDownList ID="SelNode2" runat="server" AutoPostBack="True" ForeColor="#006600">
                    <asp:ListItem Selected="True"></asp:ListItem>
                    <asp:ListItem>設備種類</asp:ListItem>
                    <asp:ListItem>區域名稱</asp:ListItem> 
                    <asp:ListItem>定位名稱</asp:ListItem>
                    <asp:ListItem>放置地點</asp:ListItem>                   
                    <asp:ListItem>維護課別</asp:ListItem>
                    <asp:ListItem>維護人員</asp:ListItem>
                    <asp:ListItem>保管課別</asp:ListItem>
                    <asp:ListItem>保管人員</asp:ListItem>
                    <asp:ListItem>維護廠商</asp:ListItem>
                    <asp:ListItem>廠牌</asp:ListItem>
                    <asp:ListItem>廠牌型式</asp:ListItem>
                    <asp:ListItem>設備狀態</asp:ListItem>
                    <asp:ListItem>用電電壓</asp:ListItem>
                    <asp:ListItem>取得來源</asp:ListItem>
                    <asp:ListItem>關機順序</asp:ListItem>
                    <asp:ListItem>建立人員</asp:ListItem>
                    <asp:ListItem>修改人員</asp:ListItem>
                </asp:DropDownList>
                <hr />
                <asp:TreeView ID="TreeView1" runat="server" ImageSet="XPFileExplorer" NodeIndent="15" NodeWrap="True"
                    OnTreeNodePopulate="TreeView1_TreeNodePopulate" ShowLines="True" OnSelectedNodeChanged="TreeView1_SelectedNodeChanged">
                    <HoverNodeStyle Font-Underline="True" ForeColor="#6666AA" />
                    <Nodes>
                        <asp:TreeNode PopulateOnDemand="True" SelectAction="SelectExpand" Text="新節點" Value="新節點">
                        </asp:TreeNode>
                    </Nodes>
                    <NodeStyle Font-Names="Tahoma" Font-Size="10pt" ForeColor="Black" HorizontalPadding="2px"
                        NodeSpacing="0px" VerticalPadding="2px" />
                    <ParentNodeStyle Font-Bold="False" />
                    <SelectedNodeStyle BackColor="#B5B5B5" Font-Underline="False" HorizontalPadding="0px"
                        VerticalPadding="0px" ForeColor="Red" />
                </asp:TreeView>
            </td>
            <td valign="top">
                <div class="style5">
                    <br /> 
                    <asp:Button ID="BtnNew" runat="server" OnClick="BtnNew_Click" Text="　新　增　" />&nbsp;&nbsp;                                                                             
                    顯示欄位：<asp:DropDownList ID="SelShow" runat="server" ForeColor="#009933" 
                        Visible="true" DataSourceID="SqlDataSourceShow" DataTextField="Item" DataValueField="Memo">
                    </asp:DropDownList>                    
                    <asp:SqlDataSource ID="SqlDataSourceShow" runat="server" ConnectionString="<%$ ConnectionStrings:IDMSConnectionString %>"
                        SelectCommand="SELECT [Item],[Memo] from Config where Kind='設備查詢' and (Config='' or Config='系統設備') order by Mark">
                    </asp:SqlDataSource>&nbsp;&nbsp; 
                    <asp:CheckBox ID="ChkPage" Checked="true" runat="server" Text="分頁" />&nbsp;&nbsp;
                    關鍵字：<asp:TextBox ID="TextKey" runat="server"></asp:TextBox>&nbsp;&nbsp;
                    <asp:Button ID="BtnSearch" runat="server" Text=" 查　詢 " OnClick="BtnSearch_Click" /> &nbsp;&nbsp;
                    <asp:Button ID="BtnDevNo" runat="server" Text="設備編號" OnClick="BtnPKNo_Click" />   &nbsp;&nbsp; 
                    <span class="style6" onclick="alert('1.關鍵字可以逗點分隔，條件包含左方樹狀項目 ； \n2.查詢範圍均限於實體設備 ； \n3.顯示欄位的設定路徑：[系統設定] -> [顯示欄位]，\n4.欄位名稱在[系統設定] -> [一般設定] -> [顯示欄位] 可查到');">
                        <u style="cursor:pointer">說明</u>
                    </span> 
                </div>
                <br />
                <asp:GridView ID="GridView1" runat="server" AllowPaging="True" AllowSorting="True"
                    AutoGenerateColumns="True" BackColor="#CCFF66" BorderColor="#999999" BorderStyle="None"
                    BorderWidth="1px" CellPadding="3" DataKeyNames="設備編號" DataSourceID="SqlDataSource1"
                    GridLines="Vertical" Font-Size="Small" OnSelectedIndexChanging="GridView1_SelectedIndexChanging"
                    Style="text-align: left">
                    <Columns>
                        <asp:CommandField ButtonType="Button" SelectText="編輯" ShowSelectButton="True" />
                    </Columns>
                    <AlternatingRowStyle BackColor="#DCDCDC" />
                    <FooterStyle BackColor="#CCCCCC" ForeColor="Black" />
                    <HeaderStyle BackColor="#000084" Font-Bold="True" ForeColor="White" />
                    <PagerSettings Mode="NumericFirstLast" />
                    <PagerStyle BackColor="#999999" ForeColor="Black" HorizontalAlign="Center" />
                    <RowStyle BackColor="#EEEEEE" ForeColor="Black" />
                    <SelectedRowStyle BackColor="#008A8C" Font-Bold="True" ForeColor="White" />
                    <SortedAscendingCellStyle BackColor="#F1F1F1" />
                    <SortedAscendingHeaderStyle BackColor="#0000A9" />
                    <SortedDescendingCellStyle BackColor="#CAC9C9" />
                    <SortedDescendingHeaderStyle BackColor="#000065" />
                </asp:GridView>
                <asp:SqlDataSource ID="SqlDataSource1" runat="server" ConnectionString="<%$ ConnectionStrings:IDMSConnectionString %>"
                    SelectCommand="" ProviderName="<%$ ConnectionStrings:IDMSConnectionString.ProviderName %>">
                </asp:SqlDataSource> 
                <br />
                <div class="style5">
                    總價：<asp:TextBox ID="TextTotal" 
                        runat="server" BackColor="#CCCCCC" ReadOnly="True" Width="120px"  
                        CssClass="style5" Font-Bold="True">0</asp:TextBox>&nbsp;&nbsp;&nbsp;&nbsp;
                    總額定電流：<asp:TextBox ID="TextSumCurrent" 
                        runat="server" BackColor="#CCCCCC" ReadOnly="True" Width="60px"  
                        CssClass="style5" Font-Bold="True">0</asp:TextBox>&nbsp;&nbsp;&nbsp;&nbsp;
                    總筆數：<asp:TextBox ID="TextCount" runat="server" BackColor="#CCCCCC" ReadOnly="True" Width="60px"  
                        CssClass="style5" Font-Bold="True"></asp:TextBox> &nbsp;&nbsp;&nbsp;&nbsp;
                    <asp:Button ID="BtnMap" runat="server" OnClick="BtnMap_Click" Text="設備分佈圖" />&nbsp;&nbsp;   
                </div>             
            </td>
        </tr>
    </table>
</asp:Content>
