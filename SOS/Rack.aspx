<%@ Page Title="機櫃空間" Language="C#" MasterPageFile="~/IDMS/MasterPage.master" AutoEventWireup="true" Debug="true" MaintainScrollPositionOnPostback="true" CodeFile="Rack.aspx.cs" Inherits="SOS_Rack" %>

<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="Server">    
</asp:Content>

<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1" runat="Server">
    <br />
    <font size="4"><b>機櫃空間配置圖</b></font>&nbsp;&nbsp;
    分區：<asp:DropDownList ID="SelRack" runat="server" ForeColor="#009933" AutoPostBack="true">
        <asp:ListItem Value="E_">Ex機櫃</asp:ListItem>
        <asp:ListItem Value="FR">FR機櫃</asp:ListItem>        
        <asp:ListItem Value="GI">GI機櫃</asp:ListItem>
        <asp:ListItem Value="H_">Hx機櫃</asp:ListItem>        
        <asp:ListItem Value="D0">D0機櫃</asp:ListItem>
        <asp:ListItem Value="D1">D1機櫃</asp:ListItem>
		<asp:ListItem Value="M0">M0機櫃</asp:ListItem>
        <asp:ListItem Value="M1">M1機櫃</asp:ListItem>
		<asp:ListItem Value="N0">N0機櫃</asp:ListItem>
        <asp:ListItem Value="N1">N1機櫃</asp:ListItem>
		<asp:ListItem Value="O0">O0機櫃</asp:ListItem>
        <asp:ListItem Value="O1">O1機櫃</asp:ListItem>
		<asp:ListItem Value="P0">P0機櫃</asp:ListItem>
        <asp:ListItem Value="P1">P1機櫃</asp:ListItem>
    </asp:DropDownList>&nbsp;&nbsp;
    
    顏色：<asp:DropDownList ID="SelColor" runat="server" ForeColor="#009933" AutoPostBack="true">
        <asp:ListItem Value="資產價值">資產價值</asp:ListItem>
        <asp:ListItem Value="關機值">關機順序</asp:ListItem>
        <asp:ListItem Value="狀態值">設備狀態</asp:ListItem>
        <asp:ListItem Value="清查值">機器清查</asp:ListItem>
        <asp:ListItem Value="接電數">接電迴路</asp:ListItem>
        <asp:ListItem Value="接網數">接網迴路</asp:ListItem>
    </asp:DropDownList>&nbsp;&nbsp;
    
    課別：<asp:DropDownList ID="SelUnit" runat="server" ForeColor="#009933" DataSourceID="SqlDataSourceUnit" DataTextField="成員" DataValueField="成員" AutoPostBack="true">
    </asp:DropDownList>
    <asp:SqlDataSource ID="SqlDataSourceUnit" runat="server" ConnectionString="<%$ ConnectionStrings:IDMSConnectionString %>"
        SelectCommand="Select [成員] from [View_組織架構] where [性質]='課別' order by [代號]"></asp:SqlDataSource>
    
    <asp:CheckBox ID="ChkMt" runat="server" Font-Size="Small" Text="人員" />
    <asp:DropDownList ID="SelMt" runat="server" ForeColor="#009933" DataSourceID="SqlDataSourceMt" DataTextField="Item" DataValueField="Item">
    </asp:DropDownList>
    <asp:SqlDataSource ID="SqlDataSourceMt" runat="server" ConnectionString="<%$ ConnectionStrings:IDMSConnectionString %>"
        SelectCommand="Select [Item] from [Config] where [Kind]=@Kind order by [Mark]">
        <SelectParameters>
            <asp:ControlParameter ControlID="SelUnit" Name="Kind" PropertyName="SelectedValue" Type="String" />
        </SelectParameters>
    </asp:SqlDataSource>&nbsp;&nbsp;

    <asp:Button ID="Button1" runat="server" Text=" 查 詢 " OnClick="Sel_Changed" />&nbsp;&nbsp;
    <asp:Label runat="server" Font-Size="Small" ForeColor="Green" Text="高度0：未設，*N：相同高度設備數，$N：資產價值，↑N：使用空間大小，顏色：由設備提示訊息可判知，粗體：清查有註記"></asp:Label>
    <br />
    
    <asp:Table ID="tblRack" runat="server" Width="100%" GridLines="Both" BackColor="white">        
    </asp:Table>
</asp:Content>
