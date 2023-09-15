<%@ Page Title="生命履歷" Language="C#" MasterPageFile="~/IDMS/MasterPage.master" AutoEventWireup="true"  MaintainScrollPositionOnPostback="true"
    CodeFile="LifeLog.aspx.cs" Inherits="Life_LifeLog"  %>

<asp:Content ID="Content1" ContentPlaceHolderID="head" Runat="Server">
    <style type="text/css">
        .style4
        {
            font-size: x-small;
            color: Blue;
            cursor: pointer;
        }
    </style>
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server">
    <p>
        <asp:CheckBox ID="ChkUnit" runat="server" Text="課別" Font-Size="Small" />
        <asp:DropDownList ID="SelUnit" runat="server" ForeColor="#009933" 
            DataSourceID="SqlDataSource1" DataTextField="Item" DataValueField="Item" 
            AutoPostBack="True" onselectedindexchanged="SelUnit_SelectedIndexChanged">
        </asp:DropDownList>
        <asp:SqlDataSource ID="SqlDataSource1" runat="server" 
            ConnectionString="<%$ ConnectionStrings:IDMSConnectionString %>" 
            SelectCommand="SELECT [Item] FROM [Config] WHERE (([Kind] = @Kind) OR ([Kind] = @Kind2)) ORDER BY [mark]">
            <SelectParameters>
                <asp:Parameter DefaultValue="維護群組" Name="Kind" Type="String" />
                <asp:Parameter DefaultValue="數值資訊組" Name="Kind2" Type="String" />
            </SelectParameters>
        </asp:SqlDataSource>&nbsp;
        <asp:CheckBox ID="ChkMt" runat="server" Text="人員" Font-Size="Small" />
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
        <span class="style4" onclick="alert('包含維護人員、原負責人、保管人員、原保管人、異動人員');"><u>說明</u></span>　

        <asp:Label ID="Label4" runat="server" Font-Size="Small" Text="日期"></asp:Label>
        <asp:DropDownList ID="SelYYYY" runat="server" ForeColor="#009933">
        </asp:DropDownList>
        <asp:Label ID="Label6" runat="server" Font-Size="Small" Text="年"></asp:Label>
        <asp:DropDownList ID="SelMM" runat="server" ForeColor="#009933"></asp:DropDownList>
        <asp:Label ID="Label7" runat="server" Font-Size="Small" Text="月"></asp:Label>
        <asp:DropDownList ID="SelDD" runat="server" ForeColor="#009933"></asp:DropDownList>
        <asp:Label ID="Label8" runat="server" Font-Size="Small" Text="日"></asp:Label>　

        <asp:Label ID="Label1" runat="server" Font-Size="Small" Text="表格"></asp:Label>
        <asp:DropDownList ID="SelTbl" runat="server" ForeColor="#006600">
            <asp:ListItem></asp:ListItem>
            <asp:ListItem>實體設備</asp:ListItem>
            <asp:ListItem>作業主機</asp:ListItem>
            <asp:ListItem>系統資源</asp:ListItem> 
            <asp:ListItem>秘總財產</asp:ListItem>
            <asp:ListItem>設備迴路</asp:ListItem>
            <asp:ListItem>系統迴路</asp:ListItem>
            <asp:ListItem>Config</asp:ListItem>
            <asp:ListItem>定位設定</asp:ListItem>
            <asp:ListItem>資產清冊</asp:ListItem>                   
        </asp:DropDownList>　

        <asp:Label ID="Label9" runat="server" Font-Size="Small" Text="主鍵"></asp:Label>
        <asp:TextBox ID="TextPK" runat="server" Columns="2"></asp:TextBox>&nbsp;&nbsp;&nbsp;&nbsp;

        <asp:Label ID="Label10" runat="server" Font-Size="Small" Text="履歷字串"></asp:Label>
        <asp:TextBox ID="TextLife" runat="server" Columns="6"></asp:TextBox>&nbsp;<font color="green">(逗點隔開)</font>&nbsp;&nbsp;
        <asp:CheckBox ID="ChkPage" Checked="true" runat="server" Font-Size="Small" Text="分頁" />&nbsp;&nbsp;&nbsp;&nbsp;

        <asp:Button ID="Button1" runat="server" Text=" 查 詢 " onclick="Button1_Click" />
    </p>
    <table width="100%"><tr><td><p>
        <asp:GridView ID="GridView1" runat="server" AllowPaging="True" 
            AllowSorting="True" AutoGenerateColumns="False" BackColor="#CCCCCC" 
            BorderColor="#999999" BorderStyle="Solid" BorderWidth="3px" CellPadding="4" 
            CellSpacing="2" DataKeyNames="履歷編號" DataSourceID="SqlDataSource3" 
            ForeColor="Black" Font-Size="Small" 
            OnSelectedIndexChanging="GridView1_SelectedIndexChanging" HorizontalAlign="Center">
            <Columns>
                <asp:CommandField ShowSelectButton="True" ButtonType="Button" HeaderText="執行" SelectText="選取" />
                <asp:BoundField DataField="履歷編號" HeaderText="No." ReadOnly="True" 
                    SortExpression="履歷編號" />
                <asp:BoundField DataField="表格名稱" HeaderText="表格名稱" SortExpression="表格名稱" />
                <asp:BoundField DataField="主鍵編號" HeaderText="PK." SortExpression="主鍵編號" />
                <asp:BoundField DataField="生命履歷" HeaderText="生命履歷" SortExpression="生命履歷" />
                <asp:BoundField DataField="維護人員" HeaderText="維護人員" SortExpression="維護人員" />
                <asp:BoundField DataField="異動人員" HeaderText="異動人員" SortExpression="異動人員" />
                <asp:BoundField DataField="異動日期"  DataFormatString="{0:yyyy/MM/dd HH:mm}" HeaderText="異動日期" SortExpression="異動日期" />
                <asp:BoundField DataField="執行機器" HeaderText="執行機器" SortExpression="執行機器" />
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
        <asp:SqlDataSource ID="SqlDataSource3" runat="server" ConnectionString="<%$ ConnectionStrings:IDMSConnectionString %>" SelectCommand="">
        </asp:SqlDataSource>
    </p></td></tr></table>
</asp:Content>

