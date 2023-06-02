<%@ Page Title="定位設定" Language="C#" MasterPageFile="~/IDMS/MasterPage.master" AutoEventWireup="true"
    CodeFile="Pointer.aspx.cs" Inherits="Config_Pointer" %>

<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="Server">
    <style type="text/css">
        .style17
        {
            width: 100%;
            border-style: solid;
            border-width: 1px;
        }
        .style18
        {
            font-size: xx-large;
            font-family: 標楷體;
            text-align: center;
        }
        .style19
        {
            font-family: 標楷體;
            font-size: large;
            width: 127px;
        }
        .style20
        {
            font-size: small;
        }
        .style21
        {
            font-size: xx-large;
            font-family: 標楷體;
            width: 127px;
            text-align: center;
        }
        .style22
        {
            font-size: xx-large;
            font-family: 標楷體;
            text-align: center;
            width: 273px;
        }
        .style23
        {
            width: 273px;
        }
        .style24
        {
            color: #006600;
            font-size: small;
        }
    </style>
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1" runat="Server">
    <br />
    <table border="1" cellpadding="0" cellspacing="0" class="style17">
        <tr>
            <td rowspan="9" valign="top">
                <asp:TreeView ID="TreeView1" runat="server" OnTreeNodePopulate="TreeView1_TreeNodePopulate"
                    ImageSet="XPFileExplorer" ShowLines="True" 
                    onselectednodechanged="TreeView1_SelectedNodeChanged" NodeIndent="15">
                    <HoverNodeStyle Font-Underline="True" ForeColor="#6666AA" />
                    <Nodes>
                        <asp:TreeNode Text="定位設定" Value="定位設定" PopulateOnDemand="True" ToolTip="由區域名稱與定位名稱展開"
                            SelectAction="SelectExpand"></asp:TreeNode>
                    </Nodes>
                    <NodeStyle Font-Names="Tahoma" Font-Size="8pt" ForeColor="Black" HorizontalPadding="2px"
                        NodeSpacing="0px" VerticalPadding="2px" />
                    <ParentNodeStyle Font-Bold="False" />
                    <SelectedNodeStyle Font-Underline="False" HorizontalPadding="0px" VerticalPadding="0px"
                        Font-Bold="True" BackColor="#B5B5B5" ForeColor="Red" />
                </asp:TreeView>
            </td>
            <td class="style21">
                項目
            </td>
            <td class="style22">
                定位設定
            </td>
            <td class="style18">
                說明
            </td>
        </tr>
        <tr>
            <td class="style19">
                定位編號
            </td>
            <td class="style23">
                <asp:TextBox ID="TextPointerNo" runat="server" Columns="4" ReadOnly="True" BackColor="Silver"
                    Font-Bold="True" Font-Size="Large" ForeColor="Red"></asp:TextBox>
            </td>
            <td class="style20">
                <asp:Label ID="lblPointerNo" runat="server" Font-Size="XX-Small"></asp:Label>
            </td>
        </tr>
        <tr>
            <td class="style19">
                定位方式
            </td>
            <td class="style23">
                <asp:DropDownList ID="SelKind" runat="server" AutoPostBack="True" ForeColor="#009933">
                    <asp:ListItem Value="機櫃">機櫃</asp:ListItem>
                    <asp:ListItem Value="分區">分區</asp:ListItem>
                    <asp:ListItem Selected="True" Value="自訂">自訂</asp:ListItem>
                </asp:DropDownList>
            </td>
            <td class="style20">
                <asp:Label ID="lblKind" runat="server" Font-Size="XX-Small"></asp:Label>
            </td>
        </tr>
        <tr>
            <td class="style19">
                區域名稱
            </td>
            <td class="style23">
                <asp:DropDownList ID="SelPlace" runat="server" AutoPostBack="True" OnSelectedIndexChanged="SelPlace_SelectedIndexChanged"
                    ForeColor="#009933" DataTextField="區域名稱" DataValueField="區域名稱">
                </asp:DropDownList>
                <asp:TextBox ID="TextPlace" runat="server" Columns="16"></asp:TextBox>
                <asp:SqlDataSource ID="SqlDataSource1" runat="server" ConnectionString="<%$ ConnectionStrings:IDMSConnectionString %>"
                    SelectCommand="SELECT DISTINCT [區域名稱] FROM [定位設定]"></asp:SqlDataSource>
            </td>
            <td class="style20" rowspan="2">
                <asp:Label ID="lblPlace" runat="server" Font-Size="XX-Small"></asp:Label>
            </td>
        </tr>
        <tr>
            <td class="style19">
                定位名稱
            </td>
            <td class="style23">
                <asp:DropDownList ID="SelPointer" runat="server" DataTextField="定位名稱" 
                    DataValueField="定位名稱" AutoPostBack="True"
                    ForeColor="#009933" 
                    onselectedindexchanged="SelPointer_SelectedIndexChanged">
                </asp:DropDownList>
                <asp:TextBox ID="TextPointer" runat="server" Columns="16"></asp:TextBox>
                <asp:SqlDataSource ID="SqlDataSource2" runat="server" ConnectionString="<%$ ConnectionStrings:IDMSConnectionString %>"
                    SelectCommand="SELECT DISTINCT [定位名稱] FROM [定位設定] WHERE ([區域名稱] = @區域名稱) and [定位方式]<>'坐標'">
                    <SelectParameters>
                        <asp:ControlParameter ControlID="SelPlace" Name="區域名稱" PropertyName="SelectedValue"
                            Type="String" />
                    </SelectParameters>
                </asp:SqlDataSource>
            </td>
        </tr>
        <tr>
            <td class="style19">
                機櫃容量
            </td>
            <td class="style23">
                <asp:DropDownList ID="SelSpace" runat="server" ForeColor="#009933">
                </asp:DropDownList>
            </td>
            <td class="style20">
                <asp:Label ID="lblSpace" runat="server" Font-Size="XX-Small"></asp:Label>
            </td>
        </tr>
        <tr>
            <td class="style19">
                顯示坐標
            </td>
            <td class="style23">
                X&nbsp; ：<asp:DropDownList ID="SelX" runat="server" ForeColor="#009933">
                </asp:DropDownList>
                Y&nbsp; ：<asp:DropDownList ID="SelY" runat="server" ForeColor="#009933">
                </asp:DropDownList>
            </td>
            <td class="style20">
                <asp:Label ID="lblXY" runat="server" Font-Size="XX-Small"></asp:Label>
            </td>
        </tr>
        <tr>
            <td class="style19">
                範圍(左上)
            </td>
            <td class="style23">
                X1：<asp:DropDownList ID="SelX1" runat="server" ForeColor="#009933">
                </asp:DropDownList>
                Y1：<asp:DropDownList ID="SelY1" runat="server" ForeColor="#009933">
                </asp:DropDownList>
            </td>
            <td class="style20" rowspan="2">
                <asp:Label ID="lblRange" runat="server" Font-Size="XX-Small"></asp:Label>
            </td>
        </tr>
        <tr>
            <td class="style19">
                範圍(右下)
            </td>
            <td class="style23">
                X2：<asp:DropDownList ID="SelX2" runat="server" ForeColor="#009933">
                </asp:DropDownList>
                Y2：<asp:DropDownList ID="SelY2" runat="server" ForeColor="#009933">
                </asp:DropDownList>
            </td>
        </tr>
        <tr>
            <td colspan="4" style="text-align: center">
                <br />
                <asp:Button ID="BtnNew" runat="server" Text="新增" OnClick="BtnNew_Click" />&nbsp;&nbsp;&nbsp;&nbsp;
                <asp:Button ID="BtnEdit" runat="server" Text="修改" OnClick="BtnEdit_Click" />&nbsp;&nbsp;&nbsp;&nbsp;
                <asp:Button ID="BtnDel" runat="server" Text="刪除" OnClientClick="return confirm('刪除前請再一次確認，所有引用到此定位的資料(例如實體設備、迴路或配電盤的放置地點)均已先刪除或修改，您確定要刪除這筆資料嗎？')" OnClick="BtnDel_Click" />&nbsp;&nbsp;&nbsp;&nbsp;
                <asp:Button ID="BtnMap" runat="server" OnClick="BtnMap_Click" Text="設備分佈圖" ToolTip="以定位編號為查詢條件" />&nbsp;&nbsp;
                <asp:TextBox ID="TextPno" runat="server" Columns="2"></asp:TextBox>&nbsp;&nbsp;                
                <asp:Button ID="BtnPlace" runat="server" Text="定位查詢" ToolTip="可以定位編號查詢放置地點" onclick="BtnPlace_Click" />
                <br />
                <br />
                <span class="style24">若異動後樹狀節點顯示有問題，請重新連結定位設定網頁以重讀資料</span>
                <br />
                <span class="style24">詳細說明請參考 3999f10 SSM(環): 機房操作管理系統地點管理相關事務</span>
            </td>
        </tr>
    </table>
</asp:Content>
