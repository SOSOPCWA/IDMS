<%@ Page Title="資料查核" Language="C#" MasterPageFile="~/IDMS/MasterPage.master" AutoEventWireup="true"  MaintainScrollPositionOnPostback="true"
    CodeFile="Check.aspx.cs" Inherits="Life_Check" Debug="true"%>

<asp:Content ID="Content1" ContentPlaceHolderID="head" Runat="Server">
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server">
    <p>
        <asp:Label ID="Label1" runat="server" Font-Size="Small" Text="表格"></asp:Label>
        <asp:DropDownList ID="SelTbl" runat="server" ForeColor="#006600">                        
            <asp:ListItem>(全部)</asp:ListItem>                                  
            <asp:ListItem>設備+作業</asp:ListItem> 
            <asp:ListItem>實體設備</asp:ListItem>
            <asp:ListItem>作業主機</asp:ListItem>
            <asp:ListItem>接網迴路</asp:ListItem>
            <asp:ListItem>秘總財產</asp:ListItem> 
            <asp:ListItem>系統資源</asp:ListItem> 
            <asp:ListItem>軟體除外</asp:ListItem>
            <asp:ListItem>軟體管理</asp:ListItem>        
        </asp:DropDownList>&nbsp;&nbsp;

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
        <asp:DropDownList ID="SelMt" runat="server" DataSourceID="SqlDataSource2"  ForeColor="#009933"
            DataTextField="Item" DataValueField="Item">
        </asp:DropDownList>
        <asp:SqlDataSource ID="SqlDataSource2" runat="server" 
            ConnectionString="<%$ ConnectionStrings:IDMSConnectionString %>" 
            SelectCommand="SELECT [Item] FROM [Config] WHERE ([Kind] = @Kind) ORDER BY [mark]">
            <SelectParameters>
                <asp:ControlParameter ControlID="SelUnit" Name="Kind" 
                    PropertyName="SelectedValue" Type="String" />
            </SelectParameters>
        </asp:SqlDataSource>&nbsp;&nbsp;

        <asp:Label ID="Label2" runat="server" Font-Size="Small" Text="查核"></asp:Label>
        <asp:DropDownList ID="SelCheck" runat="server" ForeColor="#006600" DataSourceID="SqlDataSource3" DataTextField="Checks" DataValueField="Item" AutoPostBack="true" OnSelectedIndexChanged="SelCheck_SelectedIndexChanged">            
        </asp:DropDownList>&nbsp;&nbsp;
        <asp:Label ID="lblCheck" runat="server" Font-Size="Small" ForeColor="Green"></asp:Label>
        <asp:SqlDataSource ID="SqlDataSource3" runat="server" 
            ConnectionString="<%$ ConnectionStrings:IDMSConnectionString %>" 
            SelectCommand="SELECT [Item],'('+[Config]+') '+[Item] as [Checks],[Memo] FROM [Config] WHERE [Kind] = '查核項目' ORDER BY [Config],[Mark]">
        </asp:SqlDataSource>&nbsp;&nbsp;

        <asp:CheckBox ID="ChkKey" runat="server" />
        <asp:Label ID="Label10" runat="server" Font-Size="Small" Text="關鍵字"></asp:Label>
        <asp:TextBox ID="TextKey" runat="server" Columns="6"></asp:TextBox>&nbsp;<font color="green">(逗點)</font>&nbsp;&nbsp;
                <asp:CheckBox ID="ChkPage" Checked="true" runat="server" Font-Size="Small" Text="分頁" />&nbsp;&nbsp;&nbsp;&nbsp;

        <asp:Button ID="Button1" runat="server" Text="　查　詢　" onclick="Button1_Click" />        
    </p>
    <div><table width="100%">
      <tr><td align="center">
        <asp:Button ID="Button2" runat="server" Text="重新產製查核暫存資料表" onclick="Button2_Click" ForeColor="#FF5050" />&nbsp;&nbsp;
        <asp:Label ID="Label3" runat="server" Font-Size="Small" Text="(每週一mail前才自動更新，故要手動重新產製才會是最新資料)"></asp:Label>
      </td></tr>
      <tr><td>
        <asp:GridView ID="GridView1" runat="server" AllowPaging="True"
            AllowSorting="True" AutoGenerateColumns="False" CellPadding="4" 
            DataKeyNames="表格名稱,主鍵編號" DataSourceID="SqlDataSourceCheck" 
            ForeColor="#333333" Font-Size="Small" 
            OnSelectedIndexChanging="GridView1_SelectedIndexChanging" 
            Font-Underline="False" GridLines="None" 
            ShowHeaderWhenEmpty="True" HorizontalAlign="Center">
            <AlternatingRowStyle BackColor="White" />
            <Columns>
                <asp:CommandField ShowSelectButton="True" ButtonType="Button" HeaderText="執行" SelectText="選取" />
                <asp:BoundField DataField="表格名稱" HeaderText="表格名稱" SortExpression="表格名稱" />
                <asp:BoundField DataField="主鍵編號" HeaderText="PK." SortExpression="主鍵編號" />
                <asp:BoundField DataField="資料內容" HeaderText="資料內容" SortExpression="資料內容" />
                <asp:BoundField DataField="維護人員" HeaderText="維護人員" SortExpression="維護人員" />
                <asp:BoundField DataField="相關人員" HeaderText="相關人員" SortExpression="相關人員" />
                <asp:BoundField DataField="查核結果" HeaderText="查核結果" SortExpression="查核結果" />
            </Columns>
            <EditRowStyle BackColor="#2461BF" />
            <FooterStyle BackColor="#507CD1" Font-Bold="True" ForeColor="White" />
            <HeaderStyle BackColor="#507CD1" Font-Bold="True" ForeColor="White" />
            <PagerStyle BackColor="#2461BF" ForeColor="White" HorizontalAlign="Center" />
            <RowStyle BackColor="#EFF3FB" />
            <SelectedRowStyle BackColor="#D1DDF1" Font-Bold="True" ForeColor="#333333" />
            <SortedAscendingCellStyle BackColor="#F5F7FB" />
            <SortedAscendingHeaderStyle BackColor="#6D95E1" />
            <SortedDescendingCellStyle BackColor="#E9EBEF" />
            <SortedDescendingHeaderStyle BackColor="#4870BE" />
        </asp:GridView>
        <asp:SqlDataSource ID="SqlDataSourceCheck" runat="server" ConnectionString="<%$ ConnectionStrings:IDMSConnectionString %>">
        </asp:SqlDataSource>
      </td></tr>
    </table></div>
    <br />
    <div style="text-align:center">
        總筆數：<asp:TextBox ID="TextCount" runat="server" BackColor="#CCCCCC" ReadOnly="True" Width="60px" style="text-align:center" Font-Bold="True"></asp:TextBox> &nbsp;&nbsp;
        <asp:Button ID="BtnExcel" runat="server" Text="匯出Excel檔 (html格式)" OnClick="BtnExcel_Click" />   
    </div>
</asp:Content>

