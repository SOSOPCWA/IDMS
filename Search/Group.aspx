<%@ Page Title="群組查詢" Language="C#" MasterPageFile="~/IDMS/MasterPage.master" AutoEventWireup="true" MaintainScrollPositionOnPostback="true" CodeFile="Group.aspx.cs" Inherits="Search_Group" %>

<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="Server">
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1" runat="Server">   
    <br /> 
    <div align="center">
        <asp:DropDownList ID="SelGroup" runat="server" ForeColor="#009933" DataSourceID="SqlDataSource2" DataTextField="Item" DataValueField="Item" AppendDataBoundItems="true">
            <asp:ListItem Value="">(全部)</asp:ListItem>
        </asp:DropDownList>
        <asp:SqlDataSource ID="SqlDataSource2" runat="server" ConnectionString="<%$ ConnectionStrings:IDMSConnectionString %>" 
            SelectCommand="SELECT [Item] FROM [Config] WHERE [Kind] = '維護群組' ORDER BY [mark]">
        </asp:SqlDataSource> &nbsp;&nbsp;
        <asp:Button ID="Button1" runat="server" Text="　查　詢　" OnClick="Button1_Click" /> &nbsp;&nbsp;
        <asp:Button ID="BtnExcel" runat="server" Text="匯出Excel檔 (html格式)" OnClick="BtnExcel_Click" />
        
        <br /><br />
        <table><tr><td style="font-size: small" align="left">
            <asp:GridView ID="GridView1" runat="server" AllowPaging="false" AllowSorting="True"
                BackColor="#CCCCCC" BorderColor="#999999" BorderStyle="Solid"
                BorderWidth="3px" CellPadding="4" CellSpacing="2" DataSourceID="SqlDataSource1"
                ForeColor="Black" Font-Size="Small">
                <FooterStyle BackColor="#CCCCCC" />
                <HeaderStyle BackColor="Black" Font-Bold="True" ForeColor="White" HorizontalAlign="Center" />
                <PagerStyle BackColor="#CCCCCC" ForeColor="Black" HorizontalAlign="Center" />
                <RowStyle BackColor="White" />
                <SelectedRowStyle BackColor="#000099" Font-Bold="True" ForeColor="White" />
                <SortedAscendingCellStyle BackColor="#F1F1F1" />
                <SortedAscendingHeaderStyle BackColor="#808080" />
                <SortedDescendingCellStyle BackColor="#CAC9C9" />
                <SortedDescendingHeaderStyle BackColor="#383838" />
            </asp:GridView>
            <asp:SqlDataSource ID="SqlDataSource1" runat="server" ConnectionString="<%$ ConnectionStrings:IDMSConnectionString %>" SelectCommand=""></asp:SqlDataSource>
        </td></tr></table>
    </div>
</asp:Content>
