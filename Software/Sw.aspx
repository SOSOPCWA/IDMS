<%@ Page Title="軟體保管" Language="C#" MasterPageFile="~/IDMS/MasterPage.master" AutoEventWireup="true"
    CodeFile="Sw.aspx.cs" Inherits="SoftWare_Sw" MaintainScrollPositionOnPostback="true" %>

<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="Server">
    <style type="text/css">
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
    <table border="1">
        <tr>
            <td width="240" valign="top">
                <asp:DropDownList ID="SelNode1" runat="server" AutoPostBack="True" ForeColor="#006600">
                    <asp:ListItem Selected="True">表單年度</asp:ListItem>
                    <asp:ListItem>表單編號</asp:ListItem>
                    <asp:ListItem>軟體單位</asp:ListItem>
                    <asp:ListItem>軟體人員</asp:ListItem>
                    <asp:ListItem>軟體名稱</asp:ListItem>
                    <asp:ListItem>購買版本</asp:ListItem>
                    <asp:ListItem>授權方式</asp:ListItem>
                    <asp:ListItem>軟體來源</asp:ListItem>
                    <asp:ListItem>軟體功能</asp:ListItem>                    
                    <asp:ListItem>適用機型</asp:ListItem>
                    <asp:ListItem>適用廠牌</asp:ListItem>
                    <asp:ListItem>期限說明</asp:ListItem>
                    <asp:ListItem>提供單位</asp:ListItem>
                    <asp:ListItem>軟體狀態</asp:ListItem>
                    <asp:ListItem>減損原因</asp:ListItem>
                    <asp:ListItem>減損方式</asp:ListItem>                    
                    <asp:ListItem>建立人員</asp:ListItem>
                    <asp:ListItem>修改人員</asp:ListItem>
                </asp:DropDownList>
                <asp:DropDownList ID="SelNode2" runat="server" AutoPostBack="True" ForeColor="#006600">
                    <asp:ListItem></asp:ListItem>
                    <asp:ListItem>表單年度</asp:ListItem>
                    <asp:ListItem Selected="True">表單編號</asp:ListItem>
                    <asp:ListItem>軟體單位</asp:ListItem>
                    <asp:ListItem>軟體人員</asp:ListItem>
                    <asp:ListItem>軟體名稱</asp:ListItem>
                    <asp:ListItem>購買版本</asp:ListItem>
                    <asp:ListItem>授權方式</asp:ListItem>
                    <asp:ListItem>軟體來源</asp:ListItem>
                    <asp:ListItem>軟體功能</asp:ListItem>
                    <asp:ListItem>適用機型</asp:ListItem>
                    <asp:ListItem>適用廠牌</asp:ListItem>
                    <asp:ListItem>期限說明</asp:ListItem>
                    <asp:ListItem>提供單位</asp:ListItem>
                    <asp:ListItem>軟體狀態</asp:ListItem>
                    <asp:ListItem>減損原因</asp:ListItem>
                    <asp:ListItem>減損方式</asp:ListItem>
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
            <td valign="top" class="style5">
                <div class="style5">
                    <br />
                    <asp:Button ID="BtnNew" runat="server" OnClick="BtnNew_Click" Text="　新　增　" />&nbsp;&nbsp;                   
                    顯示欄位：<asp:DropDownList ID="SelShow" runat="server" ForeColor="#009933" Visible="true" DataSourceID="SqlDataSourceShow" DataTextField="Item" DataValueField="Memo">
                    </asp:DropDownList>                    
                    <asp:SqlDataSource ID="SqlDataSourceShow" runat="server" ConnectionString="<%$ ConnectionStrings:IDMSConnectionString %>"
                        SelectCommand="SELECT [Item],[Memo] from Config where [Kind]='軟體查詢' and [Config]='軟體保管' order by [Mark]">
                    </asp:SqlDataSource>&nbsp;&nbsp; 
                    <asp:CheckBox ID="ChkPage"  Checked="true" runat="server" />
                    <asp:Label ID="Label6" runat="server" Font-Size="Small" Text="分頁"></asp:Label>&nbsp;&nbsp;  
                    關鍵字：<asp:TextBox ID="TextKey" runat="server"></asp:TextBox>&nbsp;&nbsp;     
                    <asp:DropDownList ID="SelKey" runat="server" ForeColor="#006600">
                        <asp:ListItem Selected="True" Value="">(全部)</asp:ListItem>                        
                        <asp:ListItem>軟體編號</asp:ListItem>
                        <asp:ListItem>表單編號</asp:ListItem>                        
                        <asp:ListItem>軟體單位</asp:ListItem>
                        <asp:ListItem>軟體人員</asp:ListItem>
                        <asp:ListItem>軟體名稱</asp:ListItem>
                        <asp:ListItem>購買版本</asp:ListItem>
                        <asp:ListItem>可用版本</asp:ListItem>
                        <asp:ListItem>授權方式</asp:ListItem>
                        <asp:ListItem>授權說明</asp:ListItem>
                        <asp:ListItem>軟體來源</asp:ListItem>
                        <asp:ListItem>軟體功能</asp:ListItem>
                        <asp:ListItem>適用廠牌</asp:ListItem>
                        <asp:ListItem>適用型式</asp:ListItem>
                        <asp:ListItem>購買序號</asp:ListItem>
                        <asp:ListItem>降級序號</asp:ListItem>
                        <asp:ListItem>期限說明</asp:ListItem>
                        <asp:ListItem>價格說明</asp:ListItem>
                        <asp:ListItem>提供單位</asp:ListItem>
                        <asp:ListItem>存放媒體</asp:ListItem>
                        <asp:ListItem>圖書編號</asp:ListItem>
                        <asp:ListItem>軟體附件</asp:ListItem>
                        <asp:ListItem>軟體狀態</asp:ListItem>
                        <asp:ListItem>減損原因</asp:ListItem>
                        <asp:ListItem>減損方式</asp:ListItem>                     
                        <asp:ListItem>建立人員</asp:ListItem>
                        <asp:ListItem>修改人員</asp:ListItem>
                        <asp:ListItem>備註說明</asp:ListItem>
                    </asp:DropDownList>&nbsp;&nbsp;                
                    <asp:Button ID="BtnSearch" runat="server" Text=" 查　詢 " OnClick="BtnSearch_Click" />&nbsp;&nbsp;
                    <span class="style6" onclick="alert('1.範圍均限於軟體主檔 ； \n2.關鍵字可以逗點分隔，條件包含左方樹狀項目 ； \n3.顯示欄位的設定路徑：[系統設定] -> [顯示欄位]，\n4.欄位名稱在[系統設定] -> [一般設定] -> [顯示欄位] 可查到');">
                        <u style="cursor:pointer">說明</u>
                    </span>                     
                </div>
                <br />
                <asp:GridView ID="GridView1" runat="server" AllowPaging="True" AllowSorting="True"
                    AutoGenerateColumns="True" BackColor="#CCFF66" BorderColor="#999999" BorderStyle="None"
                    BorderWidth="1px" CellPadding="3" DataKeyNames="軟體編號" DataSourceID="SqlDataSource1"
                    GridLines="Vertical" Font-Size="Small" OnSelectedIndexChanging="GridView1_SelectedIndexChanging"
                    Style="text-align: left" PageSize="100">
                    <AlternatingRowStyle BackColor="#DCDCDC" />
                    <Columns>
                        <asp:CommandField ButtonType="Button" ShowSelectButton="True" HeaderText="保管單" SelectText="編輯" />                        
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
                <asp:SqlDataSource ID="SqlDataSource1" runat="server" ConnectionString="<%$ ConnectionStrings:IDMSConnectionString %>"
                    SelectCommand="" ProviderName="<%$ ConnectionStrings:IDMSConnectionString.ProviderName %>">
                </asp:SqlDataSource> 
                <br />
                <div class="style5">
                    全部總價：<asp:TextBox ID="TextTotal" runat="server" BackColor="#CCCCCC" ReadOnly="True"
                        Width="120px" CssClass="style5" Font-Bold="True">0</asp:TextBox>&nbsp;&nbsp;&nbsp;&nbsp;
                    總筆數：<asp:TextBox ID="TextCount" runat="server" BackColor="#CCCCCC" ReadOnly="True" Width="60px"  
                        CssClass="style5" Font-Bold="True"></asp:TextBox> &nbsp;&nbsp;&nbsp;&nbsp;                    
                    <asp:Button ID="BtnExcel" runat="server" Text="匯出Excel檔 (html格式)" OnClick="BtnExcel_Click" />&nbsp;&nbsp; 
                </div>                    
            </td>
        </tr>
    </table>
</asp:Content>
