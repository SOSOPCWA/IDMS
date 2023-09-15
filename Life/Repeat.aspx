<%@ Page Title="重複資料" Language="C#" MasterPageFile="~/IDMS/MasterPage.master" AutoEventWireup="true"  MaintainScrollPositionOnPostback="true"
    CodeFile="Repeat.aspx.cs" Inherits="Life_Repeat" %>

<asp:Content ID="Content1" ContentPlaceHolderID="head" Runat="Server">
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server">    
    <p>
        &nbsp;<asp:Label ID="Label2" runat="server" Font-Size="Small" Text="維護人員："></asp:Label>
        <asp:DropDownList ID="SelUnit" runat="server" ForeColor="#009933" 
            DataSourceID="SqlDataSource1" DataTextField="Item" DataValueField="Item" 
            AutoPostBack="True" onselectedindexchanged="SelUnit_SelectedIndexChanged">
        </asp:DropDownList>
        <asp:SqlDataSource ID="SqlDataSource1" runat="server" 
            ConnectionString="<%$ ConnectionStrings:IDMSConnectionString %>" 
            SelectCommand="SELECT [Item] FROM [Config] WHERE (([Kind] = @Kind)) ORDER BY [mark]">
            <SelectParameters>
                <asp:Parameter DefaultValue="數值資訊組" Name="Kind" Type="String" />
            </SelectParameters>
        </asp:SqlDataSource>
        <asp:DropDownList ID="SelMt" runat="server" DataSourceID="SqlDataSource2" ForeColor="#009933" 
            DataTextField="Item" DataValueField="Item">
        </asp:DropDownList>
        <asp:SqlDataSource ID="SqlDataSource2" runat="server" 
            ConnectionString="<%$ ConnectionStrings:IDMSConnectionString %>" 
            SelectCommand="SELECT [Item] FROM [Config] WHERE ([Kind] = @Kind) ORDER BY [mark]">
            <SelectParameters>
                <asp:ControlParameter ControlID="SelUnit" Name="Kind" 
                    PropertyName="SelectedValue" Type="String" />
            </SelectParameters>
        </asp:SqlDataSource>

        <asp:DropDownList ID="SelRepeat" runat="server" ForeColor="#009933">
            <asp:ListItem Selected="True">全部檢查</asp:ListItem>
            <asp:ListItem>設備名稱</asp:ListItem>                        
            <asp:ListItem>財產編號</asp:ListItem>
            <asp:ListItem>重複安裝</asp:ListItem> 
            <asp:ListItem>主機名稱</asp:ListItem>
            <asp:ListItem>IP位址</asp:ListItem>
        </asp:DropDownList>

        <asp:Button ID="Button1" runat="server" Text=" 查 詢 " onclick="Button1_Click" />&nbsp;&nbsp;&nbsp;&nbsp;
        
        <asp:Label ID="Label1" runat="server" Font-Size="Medium" ForeColor="#006600" 
        Text="檢查是否有重複的欄位設定值；因不能限制欄位不能重複，故請維護人員藉此介面自行檢查重複的設定值是否適當!" style="text-align: center"></asp:Label>
    </p>
    
<p>
       <asp:TextBox ID="TextCheck" runat="server" ReadOnly="True" Rows="36" 
            TextMode="MultiLine" Width="1000px"></asp:TextBox> &nbsp;</p>
</asp:Content>

