<%@ Page Title="報廢查詢" Language="C#" MasterPageFile="~/IDMS/MasterPage.master" AutoEventWireup="true" MaintainScrollPositionOnPostback="true" CodeFile="Discard.aspx.cs" Inherits="Search_Discard" %>

<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="Server">
    <style type="text/css">
        .style4
        {
            font-size: small;
            color: Blue;
            cursor: pointer;
        }
        
        .style5
        {
            text-align: center;
            color: Green;
            font-size: x-small;
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
        }
    </style>
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1" runat="Server">   
    <br /> 
    <div class="style7">
        報廢總價：<asp:TextBox ID="TextTotal" runat="server" BackColor="#CCCCCC" ReadOnly="True"
            Width="120px" CssClass="style7" Font-Bold="True"></asp:TextBox>&nbsp;&nbsp;

        秘總：<asp:DropDownList ID="SelDiscard" runat="server" ForeColor="#009933" Visible="true">
                <asp:ListItem Value="" Text="(全部)" Selected="True"> </asp:ListItem>
                <asp:ListItem Value="No" Text="未報廢"> </asp:ListItem>
                <asp:ListItem Value="Yes" Text="已報廢"> </asp:ListItem>
              </asp:DropDownList>&nbsp;&nbsp;
        
        顯示欄位：<asp:DropDownList ID="SelShow" runat="server" ForeColor="#009933" Visible="true" DataSourceID="SqlDataSourceShow" DataTextField="Item" DataValueField="Memo">
                    </asp:DropDownList>
                    &nbsp;
                    <asp:SqlDataSource ID="SqlDataSourceShow" runat="server" ConnectionString="<%$ ConnectionStrings:IDMSConnectionString %>"
                        SelectCommand="SELECT [Item],[Memo] from Config where Kind='進階查詢' and (Config='' or Config='報廢查詢') order by Mark">
                    </asp:SqlDataSource>&nbsp;&nbsp;
        <asp:CheckBox ID="ChkPage" Checked="true" runat="server" />分頁&nbsp;&nbsp;
        關鍵字：<asp:TextBox ID="TextKey" runat="server"></asp:TextBox>&nbsp;&nbsp;

        <asp:Button ID="Button1" runat="server" Text="　查　詢　" OnClick="Button1_Click" />&nbsp;&nbsp;
        <asp:Button ID="BtnExcel" runat="server" Text="匯出Excel檔 (html格式)" OnClick="BtnExcel_Click" />&nbsp;&nbsp;
        <span class="style5">(無用資料建議刪除)</span>
    </div>    
    
    <p style="font-size: small">
        <asp:GridView ID="GridView1" runat="server" AllowPaging="True" AllowSorting="True"
            BackColor="#CCCCCC" BorderColor="#999999" BorderStyle="Solid"
            BorderWidth="3px" CellPadding="4" CellSpacing="2" DataSourceID="SqlDataSource1"
            DataKeyNames="設備編號" ForeColor="Black" Font-Size="Small" OnSelectedIndexChanging="GridView1_SelectedIndexChanging">
            <Columns>
                <asp:CommandField ButtonType="Button" SelectText="編輯" ShowSelectButton="True" />
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
        <asp:SqlDataSource ID="SqlDataSource1" runat="server" ConnectionString="<%$ ConnectionStrings:IDMSConnectionString %>" SelectCommand=""></asp:SqlDataSource>        
    </p>
</asp:Content>
