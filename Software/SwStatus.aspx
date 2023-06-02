<%@ Page Title="申請狀態" Language="C#" MasterPageFile="~/IDMS/MasterPage.master" AutoEventWireup="true" MaintainScrollPositionOnPostback="true" CodeFile="SwStatus.aspx.cs" Inherits="Software_SwStatus" %>

<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="Server">
    <style type="text/css">
        .style4
        {
            font-size: x-small;
            color: Blue;
            cursor: pointer;
        }
        
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
    <p style="font-size: small">
        軟體名稱：<asp:TextBox ID="TextKey" runat="server"></asp:TextBox>&nbsp;&nbsp;&nbsp;&nbsp;                   
        顯示欄位：<asp:DropDownList ID="SelShow" runat="server" ForeColor="#009933" Visible="true" DataSourceID="SqlDataSourceShow" DataTextField="Item" DataValueField="Memo">
        </asp:DropDownList>                    
        <asp:SqlDataSource ID="SqlDataSourceShow" runat="server" ConnectionString="<%$ ConnectionStrings:IDMSConnectionString %>"
            SelectCommand="SELECT [Item],[Memo] from Config where [Kind]='軟體查詢' and [Config]='申請狀態' order by [Mark]">
        </asp:SqlDataSource>&nbsp;&nbsp;&nbsp;&nbsp;
        <asp:CheckBox ID="ChkPage" Checked="true" Text="分頁" runat="server" />&nbsp;&nbsp;
        <asp:Button ID="Button1" runat="server" Text="　查　詢　" OnClick="Button1_Click" /> <br /><br />
    </p> 

    <p>
        <asp:GridView ID="GridView1" runat="server" AllowPaging="True" AllowSorting="True"
            BackColor="#CCCCCC" BorderColor="#999999" BorderStyle="Solid"
            BorderWidth="3px" CellPadding="4" CellSpacing="2" DataSourceID="SqlDataSource1"
            DataKeyNames="軟體編號" ForeColor="Black" Font-Size="Small" OnSelectedIndexChanging="GridView1_SelectedIndexChanging" PageSize="10">
            <Columns>
                <asp:CommandField ButtonType="Button" SelectText="編輯" HeaderText="保管單" ShowSelectButton="True" />
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
        總筆數：<asp:TextBox ID="TextCount" runat="server" BackColor="#CCCCCC" ReadOnly="True" Width="60px" CssClass="style7" Font-Bold="True"></asp:TextBox>&nbsp;&nbsp;&nbsp;&nbsp;
        <asp:Button ID="BtnExcel" runat="server" Text="匯出Excel檔 (html格式)" OnClick="BtnExcel_Click" />
    </div>
</asp:Content>
