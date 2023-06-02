<%@ Page Title="接電迴路" Language="C#" MasterPageFile="~/IDMS/MasterPage.master" AutoEventWireup="true" CodeFile="PowerTree.aspx.cs" Inherits="SOS_PowerTree" MaintainScrollPositionOnPostback="true" Debug="true"  %>

<asp:Content ID="Content1" ContentPlaceHolderID="head" Runat="Server">
    <style type="text/css">
        .style4
        {
            width: 100%;
        }
        .style5
        {
            border: 1px solid #000000;
        }
        .style6
        {
            font-size: large;
            font-weight: bold;
        }
        .style7
        {
            color: #000000;
        }
        .style9
        {
            height: 22px;
            text-align: left;
        }
        .style10
        {
            height: 22px;
            color: #000000;
        }
        .style11
        {
            color: #000000;
            height: 24px;
        }
        .style12
        {
            height: 24px;
            text-align: left;
        }
        .style13
        {
            text-align: left;
        }
        </style>
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server">

    <table border="1" class="style4">
        <tr>
            <td valign="top" width="240px">
                <asp:TreeView ID="TreeViewDn" runat="server" ImageSet="XPFileExplorer" 
                    NodeWrap="True" OnSelectedNodeChanged="TreeViewDn_SelectedNodeChanged"
                    NodeIndent="15" ontreenodepopulate="TreeViewDn_TreeNodePopulate" 
                    ShowLines="True" ExpandDepth="1" 
                    ontreenodecollapsed="TreeViewDn_TreeNodeCollapsed" 
                    ontreenodeexpanded="TreeViewDn_TreeNodeExpanded">
                    <HoverNodeStyle Font-Underline="True" ForeColor="#6666AA" />
                    <Nodes>
                        <asp:TreeNode PopulateOnDemand="True" SelectAction="SelectExpand" Text="<font color='blue' size='large'><b>接電迴路(往下)</b></font>" Value="-2" ToolTip="電力骨幹請對照EMS、實際電力標籤及教育訓練系統電力部分"></asp:TreeNode>
                    </Nodes>
                    <NodeStyle Font-Names="Tahoma" Font-Size="10pt" ForeColor="Black" HorizontalPadding="2px" NodeSpacing="0px" VerticalPadding="2px" />
                    <ParentNodeStyle Font-Bold="False" />
                    <SelectedNodeStyle BackColor="#B5B5B5" Font-Underline="False" HorizontalPadding="0px" VerticalPadding="0px" ForeColor="Red" />
                </asp:TreeView>
            </td>

            <td valign="top">
                <div class="style3" style="text-align: center">                    
                    <br />
                    <table class="style5" border="1" cellpadding="0">
                        <tr class="style7">
                            <td class="style6">
                                欄位名稱</td>
                            <td class="style6">
                                電力節點設定值</td>
                            <td class="style6">
                                執行</td>
                        </tr>
                        <tr>
                            <td class="style7">
                                設備編號</td>
                            <td class="style13">
                                <asp:Label ID="lblDevNo" runat="server"></asp:Label>
                            &nbsp;</td>
                            <td align="left" rowspan="12" valign="bottom" style="text-align: center">
                                &nbsp;　<asp:Button ID="BtnEdit" runat="server" Text="編輯設備" onclick="BtnEdit_Click" />　&nbsp;
                                <br />
                                &nbsp;　<asp:Button ID="BtnConn" runat="server" Text="編輯迴路" onclick="BtnConn_Click" />　&nbsp;
                                <br />
                                &nbsp;　<asp:Button ID="BtnPlace" runat="server" Text="放置地點" onclick="BtnPlace_Click" />　&nbsp;
                                <br />
                                &nbsp;　<asp:Button ID="BtnTree" runat="server" Text="以下節點" onclick="BtnTree_Click" />　&nbsp;
                            </td>
                        </tr>                        
                        <tr>
                            <td class="style7">
                                設備名稱</td>
                            <td class="style13">
                                <asp:Label ID="lblDevName" runat="server"></asp:Label>
                            &nbsp;</td>
                        </tr>
                        <tr>
                            <td class="style7">
                                設備種類</td>
                            <td class="style13">
                                <asp:Label ID="lblDevKind" runat="server"></asp:Label>
                            &nbsp;</td>
                        </tr>
                        <tr>
                            <td class="style7">
                                設備用途</td>
                            <td class="style13">
                                <asp:Label ID="lblPurpose" runat="server"></asp:Label>
                            &nbsp;</td>
                        </tr>
                        <tr>
                            <td class="style11">
                                放置地點</td>
                            <td class="style12">
                                <asp:Label ID="lblPlace" runat="server"></asp:Label>
                            &nbsp;</td>
                        </tr>
                        <tr>
                            <td class="style7">
                                用電電壓</td>
                            <td class="style13">
                                <asp:Label ID="lblVoltage" runat="server"></asp:Label>
                            &nbsp;</td>
                        </tr>
                        <tr>
                            <td class="style7">
                                額定電流</td>
                            <td class="style13">
                                <asp:Label ID="lblCurrent" runat="server"></asp:Label>
                            &nbsp;</td>
                        </tr>
                        <tr>
                            <td class="style7">
                                維護人員</td>
                            <td class="style13">
                                <asp:Label ID="lblMaintainor" runat="server"></asp:Label>
                            &nbsp;</td>
                        </tr>
                        <tr>
                            <td class="style7">
                                設備狀態</td>
                            <td class="style13">
                                <asp:Label ID="lblStatus" runat="server"></asp:Label>
                            &nbsp;</td>
                        </tr>
                        <tr>
                            <td class="style7">
                                備註說明</td>
                            <td class="style13">
                                <asp:Label ID="lblMemo" runat="server"></asp:Label>
                            &nbsp;</td>
                        </tr>
                    </table>
                    <br />
                </div>
            </td>

            <td valign="top" width="240px">
                <asp:TreeView ID="TreeViewUp" runat="server" ImageSet="XPFileExplorer" NodeWrap="True" NodeIndent="15" ShowLines="True" ExpandDepth="1" 
                    ontreenodepopulate="TreeViewUp_TreeNodePopulate" ontreenodeexpanded="TreeViewUp_TreeNodePopulate">
                    <HoverNodeStyle Font-Underline="True" ForeColor="#6666AA" />
                    <Nodes>
                        <asp:TreeNode PopulateOnDemand="True" SelectAction="SelectExpand" Text="<font color='blue' size='large'><b>接電迴路(往上)</b></font>" Value="-2"></asp:TreeNode>
                    </Nodes>
                    <NodeStyle Font-Names="Tahoma" Font-Size="10pt" ForeColor="Black" HorizontalPadding="2px" NodeSpacing="0px" VerticalPadding="2px" />
                    <ParentNodeStyle Font-Bold="False" />
                    <SelectedNodeStyle BackColor="#B5B5B5" Font-Underline="False" HorizontalPadding="0px" VerticalPadding="0px" ForeColor="Red" />
                </asp:TreeView>
            </td>
        </tr>
    </table>
</asp:Content>


