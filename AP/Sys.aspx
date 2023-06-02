<%@ Page Title="系統查詢" Language="C#" MasterPageFile="~/IDMS/MasterPage.master" AutoEventWireup="true" MaintainScrollPositionOnPostback="true" CodeFile="Sys.aspx.cs" Inherits="Ap_Sys" %>

<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="Server">
    <style type="text/css">
        .style5
        {
            font-size: x-small;
            text-align: center;
            color: Green;
        }
        .style6
        {
            width: 100%;
            border-style: solid;
            border-width: 1px;
        }
        .style7
        {
            text-align: center;
            font-weight: bold;
        }
    </style>
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1" runat="Server">
    <table cellpadding="0" cellspacing="0" class="style6" border="1">
        <tr>
            <td>
                <asp:CheckBox ID="ChkSysNo" runat="server" Font-Size="Small" Text="系統" />
                <asp:DropDownList ID="SelSysNo" runat="server" ForeColor="#009933" DataSourceID="SqlDataSourceSysNo"
                    DataTextField="系統選項" DataValueField="系統編號">
                </asp:DropDownList>
                <asp:SqlDataSource ID="SqlDataSourceSysNo" runat="server" ConnectionString="<%$ ConnectionStrings:IDMSConnectionString %>"
                    SelectCommand="SELECT [資源名稱] as [系統選項],[資源編號] as [系統編號] FROM [View_系統資源] where [資源種類] in ('系統','分類') order by [系統全名]">                    
                </asp:SqlDataSource>
                
                <asp:CheckBox ID="ChkKind" runat="server" Font-Size="Small" Text="種類" />
                <asp:DropDownList ID="SelKind" runat="server" ForeColor="#009933" Visible="true"
                    DataSourceID="SqlDataSourceKind" DataTextField="Item" DataValueField="Item">
                </asp:DropDownList>
                <asp:SqlDataSource ID="SqlDataSourceKind" runat="server" ConnectionString="<%$ ConnectionStrings:IDMSConnectionString %>"
                    SelectCommand="select [Item] from [Config] where [Kind]='資源種類' order by [Mark]">
                </asp:SqlDataSource>
                
                <asp:CheckBox ID="ChkUnit" runat="server" Font-Size="Small" Text="課別" />
                <asp:DropDownList ID="SelUnit" runat="server" ForeColor="#009933" Visible="true"
                    AutoPostBack="True" DataSourceID="SqlDataSourceUnit" DataTextField="成員" DataValueField="成員">
                </asp:DropDownList>
                <asp:SqlDataSource ID="SqlDataSourceUnit" runat="server" ConnectionString="<%$ ConnectionStrings:IDMSConnectionString %>"
                    SelectCommand="Select [成員] from [View_組織架構] where [性質]='課別' order by [代號]"></asp:SqlDataSource>
                
                <asp:CheckBox ID="ChkGroup" runat="server" Font-Size="Small" Text="群組" />
                <asp:DropDownList ID="SelGroup" runat="server" ForeColor="#009933" Visible="true" DataSourceID="SqlDataSourceGroup"
                    DataTextField="成員" DataValueField="成員">
                </asp:DropDownList>
                <asp:SqlDataSource ID="SqlDataSourceGroup" runat="server" ConnectionString="<%$ ConnectionStrings:IDMSConnectionString %>"
                    SelectCommand="Select [成員] from [View_組織架構] where [課別]=@課別 and [性質]='群組' order by [代號]">
                    <SelectParameters>
                        <asp:ControlParameter ControlID="SelUnit" Name="課別" PropertyName="SelectedValue" Type="String" />
                    </SelectParameters>
                </asp:SqlDataSource>
                
                <asp:CheckBox ID="ChkMan" runat="server" Font-Size="Small" Text="員工" />
                <asp:DropDownList ID="SelMan" runat="server" ForeColor="#009933" Visible="true" DataSourceID="SqlDataSourceMan"
                    DataTextField="Item" DataValueField="Item">
                </asp:DropDownList>
                <asp:SqlDataSource ID="SqlDataSourceMan" runat="server" ConnectionString="<%$ ConnectionStrings:IDMSConnectionString %>"
                    SelectCommand="Select Item from Config where Kind=@課別 order by Mark">
                    <SelectParameters>
                        <asp:ControlParameter ControlID="SelUnit" Name="課別" PropertyName="SelectedValue" Type="String" />
                    </SelectParameters>
                </asp:SqlDataSource>&nbsp;&nbsp;

                <asp:CheckBox ID="ChkAssets" runat="server" Font-Size="Small" Text="資產" />
                <asp:DropDownList ID="SelAssets" runat="server" ForeColor="#009933" Visible="true"
                    DataSourceID="SqlDataSourceAssets" DataTextField="資產名稱" DataValueField="資產編號">
                </asp:DropDownList>
                <asp:SqlDataSource ID="SqlDataSourceAssets" runat="server" ConnectionString="<%$ ConnectionStrings:IDMSConnectionString %>"
                    SelectCommand="select [Item] as [資產編號],[Item]+ ' ' +[Config] as [資產名稱] from [Config] where [Kind]='常用資產' or [Kind]='資產清冊' order by [Kind],[Item]">
                </asp:SqlDataSource>
            </td>
        </tr>
        <tr>
            <td>
                
                <asp:CheckBox ID="ChkSQL" runat="server" Font-Size="Small" Text="SQL" />
                <asp:TextBox ID="TextSQL" Width="450px" runat="server"></asp:TextBox>               
            
                <asp:CheckBox ID="ChkKey" runat="server" Font-Size="Small" Text="關鍵字" />
                <asp:TextBox ID="TextKey" runat="server" Width="120px"></asp:TextBox>&nbsp;<span class="style5">(逗點)</span>&nbsp;
                <asp:CheckBox ID="ChkPage" Checked="true" runat="server" Font-Size="Small" Text="分頁" />
            </td>
        </tr>
    </table>

    <div align="right"><asp:Button ID="BtnSearch" runat="server" Text="　查　詢　" OnClick="BtnSearch_Click" /></div>

    <p>
        <asp:GridView ID="GridView1" runat="server" AllowPaging="True" AllowSorting="True"
            BackColor="#CCCCCC" BorderColor="#999999" BorderStyle="Solid" BorderWidth="3px"
            CellPadding="4" CellSpacing="2" DataSourceID="SqlDataSource1" DataKeyNames="資源編號"
            ForeColor="Black" Font-Size="Small" OnRowCommand="GridView1_RowCommand">
            <Columns>
                <asp:ButtonField ButtonType="Button" CommandName="編號" Text="編號" HeaderText="編輯" /> 
            </Columns>
            <FooterStyle BackColor="#CCCCCC" />
            <HeaderStyle BackColor="Black" Font-Bold="True" ForeColor="White" />
            <PagerStyle BackColor="#CCCCCC" ForeColor="Black" HorizontalAlign="Center" />
            <RowStyle BackColor="White" />
            <SelectedRowStyle BackColor="#000099" Font-Bold="True" ForeColor="White" />
            <SortedAscendingCellStyle BackColor="#F1F1F1" />
            <SortedAscendingHeaderStyle BackColor="#808080" />
            <SortedDescendingCellStyle BackColor="#CAC9C9" />
            <SortedDescendingHeaderStyle BackColor="#383838" />
        </asp:GridView>
        <asp:SqlDataSource ID="SqlDataSource1" runat="server" ConnectionString="<%$ ConnectionStrings:IDMSConnectionString %>"
            SelectCommand=""></asp:SqlDataSource>
        <br />
    </p>
    <div class="style7">
        總筆數：<asp:TextBox ID="TextCount" runat="server" BackColor="#CCCCCC" ReadOnly="True" Width="60px" CssClass="style7" Font-Bold="True"></asp:TextBox>&nbsp;&nbsp; 
        <asp:Button ID="BtnExcel" runat="server" Text="匯出Excel檔 (html格式)" OnClick="BtnExcel_Click" />
    </div>
</asp:Content>
