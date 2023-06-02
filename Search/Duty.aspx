<%@ Page Title="權責資料" Language="C#" MasterPageFile="~/IDMS/MasterPage.master" AutoEventWireup="true" MaintainScrollPositionOnPostback="true" CodeFile="Duty.aspx.cs" Inherits="Search_Duty" %>

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
            font-weight: bold;
        }
        .style8
        {
            font-weight: bold;
            text-align:center
        }
    </style>
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1" runat="Server">
    <br />
    <div class="style7">        　
        <asp:DropDownList ID="SelUnit" runat="server" ForeColor="#009933" Visible="true"
                    AutoPostBack="True" DataSourceID="SqlDataSourceUnit" DataTextField="成員" DataValueField="成員">
        </asp:DropDownList>
        <asp:SqlDataSource ID="SqlDataSourceUnit" runat="server" ConnectionString="<%$ ConnectionStrings:IDMSConnectionString %>"
            SelectCommand="Select [成員] from [View_組織架構] where [性質]='課別' order by [代號]"></asp:SqlDataSource>
        &nbsp;&nbsp;        
        <asp:DropDownList ID="SelMan" runat="server" ForeColor="#009933" Visible="true" DataSourceID="SqlDataSourceMan"
            DataTextField="Item" DataValueField="Item">
        </asp:DropDownList>
        <asp:SqlDataSource ID="SqlDataSourceMan" runat="server" ConnectionString="<%$ ConnectionStrings:IDMSConnectionString %>"
            SelectCommand="Select Item from Config where Kind=@課別 order by Mark">
            <SelectParameters>
                <asp:ControlParameter ControlID="SelUnit" Name="課別" PropertyName="SelectedValue" Type="String" />
            </SelectParameters>
        </asp:SqlDataSource>&nbsp;&nbsp;&nbsp;&nbsp;
        <asp:CheckBox ID="ChkPage" Checked="true" runat="server" Font-Size="Small" Text="分頁" Visible="false" />&nbsp;&nbsp;
        <asp:Button ID="BtnSearch" runat="server" Text="　查　詢　" OnClick="BtnSearch_Click" />        
    </div>
    <hr /><!-------------------------------------------------------------------------------------------->
    <div class="style7">
        [維護群組]：<asp:Label ID="lblGroup" runat="server" Font-Size="Small" Font-Bold="True" ForeColor="Green" Text="" />
    </div>
    <br />
    <div class="style7">
        [重複資料]：<asp:Button ID="BtnRepeat" runat="server" Text="　請按此查詢　" OnClick="BtnRepeat_Click" ForeColor="Green" Font-Bold="true" /> 
    </div>
    <br />
    <div class="style7">
        [軟體授權]：<asp:Button ID="BtnSw" runat="server" Text="　請按此查詢　" OnClick="BtnSw_Click" ForeColor="Green" Font-Bold="true" /> 
    </div>
    <br />
    <hr /><!-------------------------------------------------------------------------------------------->
    <div class="style7">
        [資料查核]　筆數：<asp:TextBox ID="TextChkCount" runat="server" BackColor="#CCCCCC" ReadOnly="True" Width="60px" CssClass="style8" Font-Bold="True"></asp:TextBox>&nbsp;&nbsp; 
        <asp:Button ID="BtnChkExcel" runat="server" Text="匯出Excel檔 (html格式)" OnClick="BtnChkExcel_Click" />
    </div>
    <p>
        <asp:GridView ID="GridViewChk" runat="server" AllowPaging="True" AllowSorting="True" AutoGenerateColumns="False"
            BackColor="#CCCCCC" BorderColor="#999999" BorderStyle="Solid" BorderWidth="3px"
            CellPadding="4" CellSpacing="2" DataSourceID="SqlDataSourceChk" DataKeyNames="表格名稱,主鍵編號"
            ForeColor="Black" Font-Size="Small" OnSelectedIndexChanging="GridViewChk_SelectedIndexChanging"
            Font-Underline="False" GridLines="None" 
            ShowHeaderWhenEmpty="True">
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
        <asp:SqlDataSource ID="SqlDataSourceChk" runat="server" ConnectionString="<%$ ConnectionStrings:IDMSConnectionString %>"
            SelectCommand=""></asp:SqlDataSource>
        <br />
    </p>    

    <hr /><!-------------------------------------------------------------------------------------------->
    <div class="style7">
        [系統資源]　筆數：<asp:TextBox ID="TextSysCount" runat="server" BackColor="#CCCCCC" ReadOnly="True" Width="60px" CssClass="style8" Font-Bold="True"></asp:TextBox>&nbsp;&nbsp; 
        <asp:Button ID="BtnSysExcel" runat="server" Text="匯出Excel檔 (html格式)" OnClick="BtnSysExcel_Click" />
    </div>
    <p>
        <asp:GridView ID="GridViewSys" runat="server" AllowPaging="True" AllowSorting="True"
            BackColor="#CCCCCC" BorderColor="#999999" BorderStyle="Solid" BorderWidth="3px"
            CellPadding="4" CellSpacing="2" DataSourceID="SqlDataSourceSys" DataKeyNames="資源編號"
            ForeColor="Black" Font-Size="Small" OnRowCommand="GridViewSys_RowCommand">
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
        <asp:SqlDataSource ID="SqlDataSourceSys" runat="server" ConnectionString="<%$ ConnectionStrings:IDMSConnectionString %>"
            SelectCommand=""></asp:SqlDataSource>
        <br />
    </p>    

    <hr /><!-------------------------------------------------------------------------------------------->
    <div class="style7">
        [作業主機]　筆數：<asp:TextBox ID="TextApCount" runat="server" BackColor="#CCCCCC" ReadOnly="True" Width="60px" CssClass="style8" Font-Bold="True"></asp:TextBox>&nbsp;&nbsp; 
        <asp:Button ID="BtnApExcel" runat="server" Text="匯出Excel檔 (html格式)" OnClick="BtnApExcel_Click" />
    </div>
    <p>
        <asp:GridView ID="GridViewAp" runat="server" AllowPaging="True" AllowSorting="True"
            BackColor="#CCCCCC" BorderColor="#999999" BorderStyle="Solid" BorderWidth="3px"
            CellPadding="4" CellSpacing="2" DataSourceID="SqlDataSourceAp" DataKeyNames="作業編號"
            ForeColor="Black" Font-Size="Small" OnRowCommand="GridViewAp_RowCommand">
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
        <asp:SqlDataSource ID="SqlDataSourceAp" runat="server" ConnectionString="<%$ ConnectionStrings:IDMSConnectionString %>"
            SelectCommand=""></asp:SqlDataSource>
        <br />
    </p>
    
    <hr /><!-------------------------------------------------------------------------------------------->
    <div class="style7">
        [實體設備]　筆數：<asp:TextBox ID="TextDevCount" runat="server" BackColor="#CCCCCC" ReadOnly="True" Width="60px" CssClass="style8" Font-Bold="True"></asp:TextBox>&nbsp;&nbsp; 
        <asp:Button ID="BtnDevExcel" runat="server" Text="匯出Excel檔 (html格式)" OnClick="BtnDevExcel_Click" />
    </div>
    <p>
        <asp:GridView ID="GridViewDev" runat="server" AllowPaging="True" AllowSorting="True"
            BackColor="#CCCCCC" BorderColor="#999999" BorderStyle="Solid" BorderWidth="3px"
            CellPadding="4" CellSpacing="2" DataSourceID="SqlDataSourceDev" DataKeyNames="設備編號"
            ForeColor="Black" Font-Size="Small" OnRowCommand="GridViewDev_RowCommand">
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
        <asp:SqlDataSource ID="SqlDataSourceDev" runat="server" ConnectionString="<%$ ConnectionStrings:IDMSConnectionString %>"
            SelectCommand=""></asp:SqlDataSource>
        <br />
    </p>
</asp:Content>
