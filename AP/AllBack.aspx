<%@ Page Title="整合回溯" Language="C#" MasterPageFile="~/IDMS/MasterPage.master" AutoEventWireup="true"
    CodeFile="AllBack.aspx.cs" Inherits="AP_AllBack" MaintainScrollPositionOnPostback="true" Debug="true" %>

<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="Server">
    <style type="text/css">
        .style1
        {
            text-align: center;
        }
        .style2
        {
            width: 100%;
            border-style: solid;
            border-width: 1px;
        }
    </style>
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1" runat="Server">
    <asp:Label runat="server" Font-Size="Large" Font-Bold="True" Text="查詢條件" />：
    <asp:TextBox ID="TextPkNo" runat="server" ReadOnly="True" BackColor="Silver" Columns="2"
        AutoPostBack="true" Font-Bold="True" Font-Size="Small" CssClass="style1" ForeColor="Red"></asp:TextBox>&nbsp;&nbsp; 
    <asp:Label ID="lbltbl" runat="server" Font-Size="Large" ForeColor="Red" Text="" ToolTip="" />&nbsp;&nbsp;&nbsp;&nbsp;    
    <asp:Label ID="lblName" runat="server" Font-Size="Large" ForeColor="Red" Text="(尚未查詢)"></asp:Label>&nbsp;&nbsp;
    <asp:Label ID="lblFunc" runat="server" Font-Size="Large" ForeColor="Red" Text="請先選取某節點當查詢條件後按查詢，以產生可能影響查詢條件之上游，再選取其它(根)節點並展開，或產生報表，以查看根因節點(非根因節點不顯示)"></asp:Label>    
    <br /><br />
    <table border="1" cellpadding="0" cellspacing="0" class="style2">
        <tr>
            <td width="25%">
                <asp:LinkButton ID="BtnSysOpen" runat="server" Font-Size="Small" Text="展開" ToolTip="自所選節點展開全部樹狀結構以檢視系統根因節點(非根因節點不顯示)" OnClick="BtnSysOpen_Click" />&nbsp;&nbsp;
                <asp:LinkButton ID="BtnSysClose" runat="server" Font-Size="Small" Text="閉合" ToolTip="閉合系統迴路樹狀結構至所選節點" OnClick="BtnSysClose_Click" />
            </td>

            <td width="25%">
                <asp:LinkButton ID="BtnApOpen" runat="server" Font-Size="Small" Text="展開" ToolTip="自所選節點展開全部樹狀結構以檢視主機根因節點(非根因節點不顯示)" OnClick="BtnApOpen_Click" />&nbsp;&nbsp;
                <asp:LinkButton ID="BtnApClose" runat="server" Font-Size="Small" Text="閉合" ToolTip="閉合作業主機樹狀結構至所選節點" OnClick="BtnApClose_Click" />
            </td>            

            <td width="25%">
                <asp:LinkButton ID="BtnNetOpen" runat="server" Font-Size="Small" Text="展開" ToolTip="自所選節點展開全部樹狀結構以檢視網路根因節點(非根因節點不顯示)" OnClick="BtnNetOpen_Click" />&nbsp;&nbsp;
                <asp:LinkButton ID="BtnNetClose" runat="server" Font-Size="Small" Text="閉合" ToolTip="閉合接網迴路樹狀結構至所選節點" OnClick="BtnNetClose_Click" />
            </td>

            <td width="25%">
                <asp:LinkButton ID="BtnPowerOpen" runat="server" Font-Size="Small" Text="展開" ToolTip="自所選節點展開全部樹狀結構以檢視電力根因節點(非根因節點不顯示)" OnClick="BtnPowerOpen_Click" />&nbsp;&nbsp;
                <asp:LinkButton ID="BtnPowerClose" runat="server" Font-Size="Small" Text="閉合" ToolTip="閉合接電迴路樹狀結構至所選節點" OnClick="BtnPowerClose_Click" />
            </td>
        </tr>

        <tr>
            <td valign="top" width="25%">
                <asp:TreeView ID="TreeViewSys" runat="server" ImageSet="XPFileExplorer" NodeWrap="True"
                    OnSelectedNodeChanged="TreeViewSys_SelectedNodeChanged" NodeIndent="15" OnTreeNodePopulate="TreeViewSys_TreeNodePopulate"
                    ShowLines="True" ExpandDepth="1">
                    <HoverNodeStyle Font-Underline="True" ForeColor="#6666AA" />
                    <Nodes>
                        <asp:TreeNode PopulateOnDemand="True" SelectAction="SelectExpand" Text="<font color='blue' size='large'><b>系統迴路</b></font>" Value="0" ToolTip="請對照機房教育訓練系統中的系統架構圖"></asp:TreeNode>
                    </Nodes>
                    <NodeStyle Font-Names="Tahoma" Font-Size="10pt" ForeColor="Black" HorizontalPadding="2px"
                        NodeSpacing="0px" VerticalPadding="2px" />
                    <ParentNodeStyle Font-Bold="False" />
                    <SelectedNodeStyle BackColor="#B5B5B5" Font-Underline="False" HorizontalPadding="0px"
                        VerticalPadding="0px" ForeColor="Red" />
                </asp:TreeView>
            </td>
            
            <td valign="top" width="25%">
                <asp:TreeView ID="TreeViewAp" runat="server" ImageSet="XPFileExplorer" NodeWrap="True"
                    OnSelectedNodeChanged="TreeViewAp_SelectedNodeChanged" NodeIndent="15" OnTreeNodePopulate="TreeViewAp_TreeNodePopulate"
                    ShowLines="True" ExpandDepth="1">
                    <HoverNodeStyle Font-Underline="True" ForeColor="#6666AA" />
                    <Nodes>
                        <asp:TreeNode PopulateOnDemand="True" SelectAction="SelectExpand" Text="<font color='blue' size='large'><b>作業主機</b></font>" Value="0" ToolTip="資訊系統請資產清冊"></asp:TreeNode>
                    </Nodes>
                    <NodeStyle Font-Names="Tahoma" Font-Size="10pt" ForeColor="Black" HorizontalPadding="2px"
                        NodeSpacing="0px" VerticalPadding="2px" />
                    <ParentNodeStyle Font-Bold="False" />
                    <SelectedNodeStyle BackColor="#B5B5B5" Font-Underline="False" HorizontalPadding="0px"
                        VerticalPadding="0px" ForeColor="Red" />
                </asp:TreeView>
            </td>            

            <td valign="top" width="25%">
                <asp:TreeView ID="TreeViewNet" runat="server" ImageSet="XPFileExplorer" NodeWrap="True"
                    OnSelectedNodeChanged="TreeViewNet_SelectedNodeChanged" NodeIndent="15" OnTreeNodePopulate="TreeViewNet_TreeNodePopulate"
                    ShowLines="True" ExpandDepth="1">
                    <HoverNodeStyle Font-Underline="True" ForeColor="#6666AA" />
                    <Nodes>
                        <asp:TreeNode PopulateOnDemand="True" SelectAction="SelectExpand" Text="<font color='blue' size='large'><b>接網迴路</b></font>" Value="-3" ToolTip="網路骨幹請對照Solarwinds及教育訓練系統的網路架構圖"></asp:TreeNode>
                    </Nodes>
                    <NodeStyle Font-Names="Tahoma" Font-Size="10pt" ForeColor="Black" HorizontalPadding="2px"
                        NodeSpacing="0px" VerticalPadding="2px" />
                    <ParentNodeStyle Font-Bold="False" />
                    <SelectedNodeStyle BackColor="#B5B5B5" Font-Underline="False" HorizontalPadding="0px"
                        VerticalPadding="0px" ForeColor="Red" />
                </asp:TreeView>
            </td>

            <td valign="top" width="25%">
                <asp:TreeView ID="TreeViewPower" runat="server" ImageSet="XPFileExplorer" NodeWrap="True"
                    OnSelectedNodeChanged="TreeViewPower_SelectedNodeChanged" NodeIndent="15" OnTreeNodePopulate="TreeViewPower_TreeNodePopulate"
                    ShowLines="True" ExpandDepth="1">
                    <HoverNodeStyle Font-Underline="True" ForeColor="#6666AA" />
                    <Nodes>
                        <asp:TreeNode PopulateOnDemand="True" SelectAction="SelectExpand" Text="<font color='blue' size='large'><b>接電迴路</b></font>" Value="-2" ToolTip="電力骨幹請對照EMS、實際電力標籤及教育訓練系統電力部分"></asp:TreeNode>
                    </Nodes>
                    <NodeStyle Font-Names="Tahoma" Font-Size="10pt" ForeColor="Black" HorizontalPadding="2px"
                        NodeSpacing="0px" VerticalPadding="2px" />
                    <ParentNodeStyle Font-Bold="False" />
                    <SelectedNodeStyle BackColor="#B5B5B5" Font-Underline="False" HorizontalPadding="0px"
                        VerticalPadding="0px" ForeColor="Red" />
                </asp:TreeView>
            </td>
        </tr>
    </table>
    <br />
    <div align="right">        
        <asp:Button ID="BtnSearch" runat="server" ForeColor="Red" Font-Size="Large" Font-Bold="true" Text="　查　詢　" ToolTip="依選取的節點作為查詢條件往上查詢整合回溯，以及產生根因節點(非根因節點不產生)" OnClick="BtnSearch_Click" />&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;        
        <asp:LinkButton ID="LinkEdit" runat="server" Font-Size="Small" OnClick="LinkEdit_Click" Text="編輯" ToolTip="可連結到編輯介面以檢視或編輯所選取節點之詳細資料"></asp:LinkButton>&nbsp;&nbsp;
        <asp:LinkButton ID="LinkTree" runat="server" Font-Size="Small" OnClick="LinkTree_Click" Text="樹狀" ToolTip="可連結到樹狀迴路介面以檢視所選取節點之上下游關係"></asp:LinkButton>&nbsp;&nbsp;
        <asp:LinkButton ID="LinkConn" runat="server" Font-Size="Small" OnClick="LinkConn_Click" Text="迴路" ToolTip="可連結到迴路設定介面以設定所選取節點之迴路上下游"></asp:LinkButton>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;         
        <asp:CheckBox ID="ChkPower" runat="server" Text="列出電力" Checked="true" Enabled="false" Font-Size="Small" ToolTip="報表是否需列出電力設備之資料" />
        <asp:CheckBox ID="ChkNet" runat="server" Text="列出網路" Checked="true" Enabled="false" Font-Size="Small" ToolTip="報表是否需列出網路設備之資料" />
        <asp:Button ID="BtnReport" runat="server" Font-Size="Large" Text="　報　表　" ToolTip="列出所有上游根因節點之清單，分為查詢條件、實體設備、作業主機、系統資源等四大區塊" OnClick="BtnReport_Click" />&nbsp;&nbsp;&nbsp;&nbsp;
        <asp:Button ID="BtnMap" runat="server" Font-Size="Large" OnClick="BtnMap_Click" Text="設備分佈圖" ToolTip="在機房平面圖點出所有上游根因節點設備" />
    </div>
</asp:Content>
