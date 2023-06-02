<%@ Page Title="秘總財產" Language="C#" MasterPageFile="~/IDMS/MasterPage.master" AutoEventWireup="true" CodeFile="Property.aspx.cs" Inherits="Property_Property" MaintainScrollPositionOnPostback="true" Debug="true" %>

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
            font-size: X-Small;
            color: #006600;
        }
    </style>
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server">

    <table border="1" class="style4">
        <tr>
            <td width="240" valign="top">
                <asp:DropDownList ID="SelNode1" runat="server" AutoPostBack="True" ForeColor="#006600">                    
                    <asp:ListItem>財產大類</asp:ListItem>
                    <asp:ListItem>財產別名</asp:ListItem>
                    <asp:ListItem>廠牌</asp:ListItem>
                    <asp:ListItem>廠牌型式</asp:ListItem>
                    <asp:ListItem>計量單位</asp:ListItem>
                    <asp:ListItem>使用年限(月)</asp:ListItem>
                    <asp:ListItem Selected="True">保管課別</asp:ListItem>
                    <asp:ListItem>保管人員</asp:ListItem>
                </asp:DropDownList>
                <asp:DropDownList ID="SelNode2" runat="server" AutoPostBack="True" ForeColor="#006600">
                    <asp:ListItem></asp:ListItem>
                    <asp:ListItem>財產大類</asp:ListItem>
                    <asp:ListItem>財產別名</asp:ListItem>
                    <asp:ListItem>廠牌</asp:ListItem>
                    <asp:ListItem>廠牌型式</asp:ListItem>
                    <asp:ListItem>計量單位</asp:ListItem>
                    <asp:ListItem>使用年限(月)</asp:ListItem>
                    <asp:ListItem>保管課別</asp:ListItem>
                    <asp:ListItem Selected="True">保管人員</asp:ListItem>
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

            <td valign="top"> 
                <div class="style5">
                    <br />
                    日期：<asp:DropDownList ID="SelBuy" runat="server" ForeColor="#006600">
                        <asp:ListItem  Value="" Selected="True">(全部)</asp:ListItem>
                        <asp:ListItem Value="30">最近一月</asp:ListItem>
                        <asp:ListItem Value="90">最近三月</asp:ListItem>
                        <asp:ListItem Value="365">最近一年</asp:ListItem>
                        <asp:ListItem Value="1095">最近三年</asp:ListItem>
                        <asp:ListItem Value="3650">最近十年</asp:ListItem>
                    </asp:DropDownList>&nbsp;&nbsp; 
                    
                    登錄IDMS：<asp:DropDownList ID="SelIDMS" runat="server" ForeColor="#009933" Visible="true">
                        <asp:ListItem Value=""></asp:ListItem>
                        <asp:ListItem Value="已登">已登</asp:ListItem>
                        <asp:ListItem Value="未登">未登</asp:ListItem>
                    </asp:DropDownList>&nbsp;&nbsp;                  
                    
                    <asp:CheckBox ID="ChkOffice" Checked="true" runat="server" />
                    <asp:Label ID="Label2" runat="server" Font-Size="Small" Text="含辦公"></asp:Label>&nbsp;&nbsp;

                    <asp:CheckBox ID="ChkMark" runat="server" />
                    <asp:Label ID="Label4" runat="server" Font-Size="Small" Text="含報廢"></asp:Label>&nbsp;&nbsp;  

                    <asp:CheckBox ID="ChkPage" Checked="true" runat="server" Text="分頁" />&nbsp;&nbsp;
                    關鍵字：<asp:TextBox ID="TextKey" runat="server"></asp:TextBox>&nbsp;&nbsp;                
                    <asp:Button ID="BtnSearch" runat="server" Text=" 查　詢 " OnClick="BtnSearch_Click" />&nbsp;&nbsp; 
                    <asp:Button ID="BtnBuy" runat="server" Text="購買日期" OnClick="BtnBuy_Click" />                    
                </div>

                <span class="style6">＊ 關鍵字查詢：可以空白分隔多個關鍵字 ； 購買日期：yyyy/mm/dd，月日可略。</span> 
                <br />
                <span class="style6">＊ 點選設備按鈕，會以財產編號為Key去查實體設備；若無資料，編輯介面會帶入財產資料以便做新增。</span> 
                          
                <asp:GridView ID="GridView1" runat="server" AllowPaging="True" 
                    AllowSorting="True" CellPadding="4" DataSourceID="SqlDataSource1" 
                    GridLines="None" Font-Size="Small" OnSelectedIndexChanging="GridView1_SelectedIndexChanging"
                    Style="text-align: left" AutoGenerateColumns="False" ForeColor="#333333">
                    <Columns>
                        <asp:CommandField ButtonType="Button" ShowSelectButton="True" HeaderText="設備" SelectText="設備" />
                        <asp:BoundField DataField="財產編號A" HeaderText="財產編號A" SortExpression="財產編號A" />
                        <asp:BoundField DataField="財產編號B" HeaderText="財產編號B" SortExpression="財產編號B" />
                        <asp:BoundField DataField="財產大類" HeaderText="財產大類" SortExpression="財產大類" />
                        <asp:BoundField DataField="財產別名" HeaderText="財產別名" SortExpression="財產別名" />
                        <asp:BoundField DataField="廠牌" HeaderText="廠牌" SortExpression="廠牌" />
                        <asp:BoundField DataField="型式" HeaderText="型式" SortExpression="型式" />
                        <asp:BoundField DataField="數量" HeaderText="數量" SortExpression="數量" />
                        <asp:BoundField DataField="計量單位" HeaderText="計量單位" SortExpression="計量單位" />
                        <asp:BoundField DataField="價值" HeaderText="價值" SortExpression="價值" />
                        <asp:BoundField DataField="購買日期" DataFormatString="{0:yyyy/MM/dd}" HeaderText="購買日期" SortExpression="購買日期" />
                        <asp:BoundField DataField="使用年限" HeaderText="使用年限(月)" SortExpression="使用年限" />
                        <asp:BoundField DataField="保管人員" HeaderText="保管人員" SortExpression="保管人員" />
                        <asp:BoundField DataField="無效註記" HeaderText="無效註記" SortExpression="無效註記" />
                        <asp:BoundField DataField="財產附屬" HeaderText="財產附屬" SortExpression="財產附屬" />
                    </Columns>
                    <AlternatingRowStyle BackColor="White" />
                    <FooterStyle BackColor="#990000" ForeColor="White" Font-Bold="True" />
                    <HeaderStyle BackColor="#990000" Font-Bold="True" ForeColor="White" />
                    <PagerStyle BackColor="#FFCC66" ForeColor="#333333" HorizontalAlign="Center" />
                    <RowStyle BackColor="#FFFBD6" ForeColor="#333333" />
                    <SelectedRowStyle BackColor="#FFCC66" Font-Bold="True" ForeColor="Navy" />
                    <SortedAscendingCellStyle BackColor="#FDF5AC" />
                    <SortedAscendingHeaderStyle BackColor="#4D0000" />
                    <SortedDescendingCellStyle BackColor="#FCF6C0" />
                    <SortedDescendingHeaderStyle BackColor="#820000" />
                </asp:GridView>
                <asp:SqlDataSource ID="SqlDataSource1" runat="server" 
                    ConnectionString="<%$ ConnectionStrings:IDMSConnectionString %>"                     
                    SelectCommand="" 
                    ProviderName="<%$ ConnectionStrings:IDMSConnectionString.ProviderName %>">                    
                </asp:SqlDataSource>    
                <br />
                <div class="style5">
                    總價：<asp:TextBox ID="TextTotal" runat="server" BackColor="#CCCCCC" ReadOnly="True" Width="150px"  
                        CssClass="style5" Font-Bold="True">0</asp:TextBox>&nbsp;&nbsp;&nbsp;&nbsp;
                    總筆數：<asp:TextBox ID="TextCount" runat="server" BackColor="#CCCCCC" ReadOnly="True" Width="60px"  
                        CssClass="style5" Font-Bold="True"></asp:TextBox> &nbsp;&nbsp; &nbsp;&nbsp;
                    <asp:Button ID="BtnExcel" runat="server" Text="匯出Excel檔 (html格式)" OnClick="BtnExcel_Click" />   
                </div>             
            </td>
        </tr>
    </table>    
</asp:Content>


