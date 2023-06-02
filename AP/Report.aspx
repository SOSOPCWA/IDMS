<%@ Page Language="C#" Title="整合迴路報表" AutoEventWireup="true" CodeFile="Report.aspx.cs" Inherits="Ap_Report" MaintainScrollPositionOnPostback="true" Debug="true" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <style type="text/css">
    </style>
</head>
<body>
<form id="form1" runat="server">
    <div style="text-align: center; font-family: 標楷體;">
        <table width="100%"><tr>
            <td width="200px"></td>
            <td><asp:Label runat="server" Text="整合迴路查詢報表" Style="font-weight: 700; font-size:xx-large" /></td>
            <td width="200px" align="right"><asp:Label ID="lblDT" runat="server" Style="font-size:small" /></td>
        </tr></table>        
    </div>
    <asp:Label runat="server" Text="【查詢條件】：" Style="font-weight: 500; font-size: large" />
    <asp:Label ID="lblSearch" runat="server" Text="" />
    <br /><br />

    <asp:Label ID="lblDev" runat="server" Text="【實體設備】：" Style="font-weight: 500; font-size: large" />
    <asp:Button ID="BtnExcel" runat="server" Text="匯出Excel檔" OnClick="BtnDevExcel_Click" />
    <asp:GridView ID="GridViewDev" runat="server" AllowSorting="True"
            BackColor="White" BorderColor="#CCCCCC" BorderStyle="None" BorderWidth="0px"
            CellPadding="4" DataSourceID="SqlDataSourceDev" DataKeyNames="設備編號"
            ForeColor="Black" Font-Size="Small" GridLines="Horizontal">
            <FooterStyle BackColor="#CCCC99" ForeColor="Black" />
            <HeaderStyle BackColor="LightGray" Font-Bold="True" ForeColor="Black" />
            <PagerStyle BackColor="White" ForeColor="Black" HorizontalAlign="Right" />
            <SelectedRowStyle BackColor="#CC3333" Font-Bold="True" ForeColor="White" />
            <SortedAscendingCellStyle BackColor="#F7F7F7" />
            <SortedAscendingHeaderStyle BackColor="#4B4B4B" />
            <SortedDescendingCellStyle BackColor="#E5E5E5" />
            <SortedDescendingHeaderStyle BackColor="#242121" />            
    </asp:GridView>
    <asp:SqlDataSource ID="SqlDataSourceDev" runat="server" ConnectionString="<%$ ConnectionStrings:IDMSConnectionString %>" SelectCommand=""></asp:SqlDataSource>
    <br />

    <asp:Label ID="lblAp" runat="server" Text="【作業主機】：" Style="font-weight: 500; font-size: large" />
    <asp:Button ID="BtnApExcel" runat="server" Text="匯出Excel檔" OnClick="BtnApExcel_Click" />
    <asp:GridView ID="GridViewAp" runat="server" AllowSorting="True"
            BackColor="White" BorderColor="#CCCCCC" BorderStyle="None" BorderWidth="0px"
            CellPadding="4" DataSourceID="SqlDataSourceAp" DataKeyNames="作業編號"
            ForeColor="Black" Font-Size="Small" GridLines="Horizontal">
            <FooterStyle BackColor="#CCCC99" ForeColor="Black" />
            <HeaderStyle BackColor="LightGray" Font-Bold="True" ForeColor="Black" />
            <PagerStyle BackColor="White" ForeColor="Black" HorizontalAlign="Right" />
            <SelectedRowStyle BackColor="#CC3333" Font-Bold="True" ForeColor="White" />
            <SortedAscendingCellStyle BackColor="#F7F7F7" />
            <SortedAscendingHeaderStyle BackColor="#4B4B4B" />
            <SortedDescendingCellStyle BackColor="#E5E5E5" />
            <SortedDescendingHeaderStyle BackColor="#242121" />            
    </asp:GridView>
    <asp:SqlDataSource ID="SqlDataSourceAp" runat="server" ConnectionString="<%$ ConnectionStrings:IDMSConnectionString %>" SelectCommand=""></asp:SqlDataSource>
    <br />

    <asp:Label ID="lblSys" runat="server" Text="【系統資源】：" Style="font-weight: 500; font-size: large" />
    <asp:Button ID="BtnSysExcel" runat="server" Text="匯出Excel檔" OnClick="BtnSysExcel_Click" />
    <asp:GridView ID="GridViewSys" runat="server" AllowSorting="True"
            BackColor="White" BorderColor="#CCCCCC" BorderStyle="None" BorderWidth="0px"
            CellPadding="4" DataSourceID="SqlDataSourceSys" DataKeyNames="資源編號"
            ForeColor="Black" Font-Size="Small" GridLines="Horizontal">
            <FooterStyle BackColor="#CCCC99" ForeColor="Black" />
            <HeaderStyle BackColor="LightGray" Font-Bold="True" ForeColor="Black" />
            <PagerStyle BackColor="White" ForeColor="Black" HorizontalAlign="Right" />
            <SelectedRowStyle BackColor="#CC3333" Font-Bold="True" ForeColor="White" />
            <SortedAscendingCellStyle BackColor="#F7F7F7" />
            <SortedAscendingHeaderStyle BackColor="#4B4B4B" />
            <SortedDescendingCellStyle BackColor="#E5E5E5" />
            <SortedDescendingHeaderStyle BackColor="#242121" />            
    </asp:GridView>
    <asp:SqlDataSource ID="SqlDataSourceSys" runat="server" ConnectionString="<%$ ConnectionStrings:IDMSConnectionString %>" SelectCommand=""></asp:SqlDataSource>
</form>
</body>
</html>
