<%@ Page Title="作業主機" Language="C#" MasterPageFile="~/IDMS/MasterPage.master" AutoEventWireup="true" CodeFile="AP.aspx.cs" Inherits="AP_AP" MaintainScrollPositionOnPostback="true" %>

<asp:Content ID="Content1" ContentPlaceHolderID="head" Runat="Server">
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
<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server">

    <table border="1" class="style4">
        <tr>
            <td width="240" valign="top">
                <asp:DropDownList ID="SelNode1" runat="server" AutoPostBack="True" ForeColor="#006600">
                    <asp:ListItem Selected="True">系統全名</asp:ListItem>
                    <asp:ListItem>作業平台</asp:ListItem>
                    <asp:ListItem>安內外</asp:ListItem>
                    <asp:ListItem>緊急程度</asp:ListItem>
                    <asp:ListItem>作業狀態</asp:ListItem>
                    <asp:ListItem>維護課別</asp:ListItem>
                    <asp:ListItem>維護人員</asp:ListItem>
                    <asp:ListItem>建立人員</asp:ListItem>
                    <asp:ListItem>修改人員</asp:ListItem>
                </asp:DropDownList>
                <asp:DropDownList ID="SelNode2" runat="server" AutoPostBack="True" ForeColor="#006600">
                    <asp:ListItem Selected="True"></asp:ListItem>
                    <asp:ListItem>系統全名</asp:ListItem>
                    <asp:ListItem>作業平台</asp:ListItem>
                    <asp:ListItem>安內外</asp:ListItem>
                    <asp:ListItem>緊急程度</asp:ListItem>
                    <asp:ListItem>作業狀態</asp:ListItem>
                    <asp:ListItem>維護課別</asp:ListItem>
                    <asp:ListItem>維護人員</asp:ListItem>
                    <asp:ListItem>建立人員</asp:ListItem>
                    <asp:ListItem>修改人員</asp:ListItem>
                </asp:DropDownList>
                
                <hr />

                <asp:TreeView ID="TreeView1" runat="server" ImageSet="XPFileExplorer" NodeWrap="True"
                    NodeIndent="15" ontreenodepopulate="TreeView1_TreeNodePopulate" 
                    ShowLines="True" onselectednodechanged="TreeView1_SelectedNodeChanged">
                    <HoverNodeStyle Font-Underline="True" ForeColor="#6666AA" />
                    <Nodes>
                        <asp:TreeNode PopulateOnDemand="True" SelectAction="SelectExpand" Text="新節點" 
                            Value="新節點"></asp:TreeNode>
                    </Nodes>
                    <NodeStyle Font-Names="Tahoma" Font-Size="10pt" ForeColor="Black" 
                        HorizontalPadding="2px" NodeSpacing="0px" VerticalPadding="2px" />
                    <ParentNodeStyle Font-Bold="False" />
                    <SelectedNodeStyle BackColor="#B5B5B5" Font-Underline="False" 
                        HorizontalPadding="0px" VerticalPadding="0px" ForeColor="Red" />
                </asp:TreeView>
            </td>

            <td valign="top" class="style5"> 
                <div class="style5">
                    <br />
                    <asp:Button ID="BtnNew" runat="server" onclick="BtnNew_Click" Text="　新　增　" />　&nbsp;&nbsp                     
                    <asp:CheckBox ID="ChkPage" Checked="true" runat="server" Text="分頁" />&nbsp;&nbsp;
                    關鍵字：<asp:TextBox ID="TextKey" runat="server"></asp:TextBox>&nbsp;&nbsp;                
                    <asp:Button ID="BtnSearch" runat="server" Text=" 查　詢 " OnClick="BtnSearch_Click" />&nbsp;&nbsp; 
                    <asp:Button ID="BtnApNo" runat="server" Text="作業編號" OnClick="BtnPKNo_Click" />&nbsp;&nbsp;  
                    <span class="style6" onclick="alert('1.關鍵字可以逗點分隔，條件包含左方樹狀項目 ； \n2.範圍均限於作業主機 ；');">
                        <u style="cursor:pointer">說明</u>
                    </span>                      
                </div>
                <br />          
                <asp:GridView ID="GridView1" runat="server" AllowPaging="True" 
                    AllowSorting="True" AutoGenerateColumns="False" BackColor="#CCFF66" 
                    BorderColor="#999999" BorderStyle="None" BorderWidth="1px" CellPadding="3" 
                    DataKeyNames="作業編號" DataSourceID="SqlDataSource1" GridLines="Vertical" 
                    Font-Size="Small" 
                    onselectedindexchanging="GridView1_SelectedIndexChanging" 
                    style="text-align: left">
                    <AlternatingRowStyle BackColor="#DCDCDC" />
                    <Columns>
                        <asp:CommandField ButtonType="Button" ShowSelectButton="True" HeaderText="執行" SelectText="選取" />
                        <asp:BoundField DataField="作業編號" HeaderText="No." SortExpression="作業編號" ReadOnly="True" Visible="True" />
                        <asp:BoundField DataField="主機名稱" HeaderText="主機名稱" SortExpression="主機名稱" />
                        <asp:BoundField DataField="系統名稱" HeaderText="系統名稱" SortExpression="系統名稱" />
                        <asp:BoundField DataField="主要功能" HeaderText="主要功能" SortExpression="主要功能" />
                        <asp:BoundField DataField="系統功能" HeaderText="系統功能" SortExpression="系統功能" />
                        <asp:BoundField DataField="作業平台" HeaderText="作業平台" SortExpression="作業平台" />
                        <asp:BoundField DataField="IP" HeaderText="IP" SortExpression="IP" />
                        <asp:BoundField DataField="安內外" HeaderText="安內外" SortExpression="安內外" />
                        <asp:BoundField DataField="緊急程度" HeaderText="緊急程度" SortExpression="緊急程度" />
                        <asp:BoundField DataField="作業狀態" HeaderText="作業狀態" SortExpression="作業狀態" />
                        <asp:BoundField DataField="維護人員" HeaderText="維護人員" SortExpression="維護人員" />
                        <asp:BoundField DataField="備註說明" HeaderText="備註說明" SortExpression="備註說明" />
                        <asp:BoundField DataField="建立日期" DataFormatString="{0:yyyy/MM/dd}" 
                            HeaderText="建立日期" SortExpression="建立日期" />
                        <asp:BoundField DataField="修改日期" DataFormatString="{0:yyyy/MM/dd}" 
                            HeaderText="修改日期" SortExpression="修改日期" />
                    </Columns>
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
                <asp:SqlDataSource ID="SqlDataSource1" runat="server" 
                    ConnectionString="<%$ ConnectionStrings:IDMSConnectionString %>"                     
                    SelectCommand="" 
                    ProviderName="<%$ ConnectionStrings:IDMSConnectionString.ProviderName %>">                    
                </asp:SqlDataSource> 
                <br />
                <div class="style5">
                    總筆數：<asp:TextBox ID="TextCount" runat="server" BackColor="#CCCCCC" ReadOnly="True" Width="60px" CssClass="style5" Font-Bold="True"></asp:TextBox>    
                </div>                    
            </td>
        </tr>
    </table>
</asp:Content>


