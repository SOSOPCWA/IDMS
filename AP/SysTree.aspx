<%@ Page Title="系統迴路" Language="C#" MasterPageFile="~/IDMS/MasterPage.master" AutoEventWireup="true" CodeFile="SysTree.aspx.cs" Inherits="AP_SysTree" MaintainScrollPositionOnPostback="true" Debug="true" %>

<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="Server">
    <style type="text/css">
        .style1
        {
            text-align: center;
        }
        .style17
        {
            width: 100%;
            border-style: solid;
            border-width: 1px;
        }
        .style19
        {
            font-family: 標楷體;
            font-size: large;
        }
        .style20
        {
            font-size: small;
        }
        .style23
        {
            font-family: 標楷體;
            font-size: XX-large;
        }
    </style>
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1" runat="Server">
    <br />
    <table border="1" cellpadding="0" cellspacing="0" class="style17">
        <tr>
            <td valign="top" width="280px">
                <asp:TreeView ID="TreeViewDn" runat="server" ImageSet="XPFileExplorer" NodeWrap="True"
                    OnSelectedNodeChanged="TreeViewDn_SelectedNodeChanged" NodeIndent="15" OnTreeNodePopulate="TreeViewDn_TreeNodePopulate"
                    ShowLines="True" ExpandDepth="1" OnTreeNodeCollapsed="TreeViewDn_TreeNodeCollapsed"
                    OnTreeNodeExpanded="TreeViewDn_TreeNodeExpanded">
                    <HoverNodeStyle Font-Underline="True" ForeColor="#6666AA" />
                    <Nodes>
                        <asp:TreeNode PopulateOnDemand="True" SelectAction="SelectExpand" Text="<font color='blue' size='large'><b>系統迴路(往下)</b></font>"
                            ToolTip="0. 請對照機房教育訓練系統的系統架構圖" Value="0"></asp:TreeNode>
                    </Nodes>
                    <NodeStyle Font-Names="Tahoma" Font-Size="10pt" ForeColor="Black" HorizontalPadding="2px"
                        NodeSpacing="0px" VerticalPadding="2px" />
                    <ParentNodeStyle Font-Bold="False" />
                    <SelectedNodeStyle BackColor="#B5B5B5" Font-Underline="False" HorizontalPadding="0px"
                        VerticalPadding="0px" ForeColor="Red" />
                </asp:TreeView>
            </td>
            <td valign="top">
                <table border="1" cellpadding="0" cellspacing="0" width="100%">
                    <tr>
                        <td valign="top">
                            <table width="100%">
                                <tr>
                                    <td width="50%" valign="top">
                                        系統資源(<font color="red">中游</font>)：
                                        <asp:TextBox ID="TextSysNo" runat="server" Columns="4" ReadOnly="True" BackColor="Silver"
                                            AutoPostBack="true" Font-Bold="True" Font-Size="Large" CssClass="style1" ForeColor="Red"></asp:TextBox>&nbsp;&nbsp;
                                        <asp:LinkButton ID="LinkSys" runat="server" Font-Size="Small" OnClick="LinkSys_Click" Text="編輯"></asp:LinkButton>&nbsp;&nbsp;
                                        <asp:LinkButton ID="BtnLife" runat="server" Font-Size="Small" OnClick="BtnLife_Click" Text="履歷"></asp:LinkButton>
                                        <br />
                                        <asp:Label ID="lblSys" runat="server" Font-Size="Small" ForeColor="Green"></asp:Label>
                                    </td>
                                    <td width="50%" valign="top">
                                        <asp:Label ID="helpSysTree" runat="server" Font-Size="Small"></asp:Label>
                                    </td>
                                </tr>
                            </table>
                        </td>
                        <td rowspan="2" valign="top" class="style20" width="240px">
                            <asp:TreeView ID="TreeViewUp" runat="server" ImageSet="XPFileExplorer" NodeWrap="True" NodeIndent="15" OnTreeNodePopulate="TreeViewUp_TreeNodePopulate" ShowLines="True" ExpandDepth="1">
                                <HoverNodeStyle Font-Underline="True" ForeColor="#6666AA" />
                                <Nodes>
                                    <asp:TreeNode PopulateOnDemand="True" SelectAction="SelectExpand"></asp:TreeNode>
                                </Nodes>
                                <NodeStyle Font-Names="Tahoma" Font-Size="10pt" ForeColor="Black" HorizontalPadding="2px"
                                    NodeSpacing="0px" VerticalPadding="2px" />
                                <ParentNodeStyle Font-Bold="False" />
                                <SelectedNodeStyle BackColor="#B5B5B5" Font-Underline="False" HorizontalPadding="0px"
                                    VerticalPadding="0px" ForeColor="Red" />
                            </asp:TreeView>
                        </td>
                    </tr>
                    <tr>
                        <td>
                            <table border="1" width="100%">
                                <tr align="center">
                                    <td>系統功能待選清單</td>
                                    <td width="80px" align="center">執行</td>
                                    <td><font color="red"><b>上游系統功能</b></font></td>
                                </tr>
                                <tr align="center">
                                    <td rowspan="3">
                                        <asp:ListBox ID="ListSysTree" runat="server" ForeColor="#009933" AutoPostBack="true" OnSelectedIndexChanged="ListSysTree_SelectedIndexChanged" AppendDataBoundItems="False" Rows="10">
                                        </asp:ListBox>
                                    </td>
                                    <td align="center">
                                        <asp:LinkButton ID="LinkUpSysTreeAdd" runat="server" Font-Size="Small" OnClick="LinkUpSysTreeAdd_Click"
                                            Text=">> 新增 >>" ToolTip="新增左方所選取之系統到右方上游"></asp:LinkButton><br />
                                        <br /><br />
                                        <asp:LinkButton ID="LinkUpSysTreeDel" runat="server" Font-Size="Small" OnClick="LinkUpSysTreeDel_Click"
                                            Text="<< 刪除 <<" ToolTip="刪除右方上游所選取之系統迴路"></asp:LinkButton>
                                    </td>
                                    <td>
                                        <asp:ListBox ID="ListUpSysTree" runat="server" ForeColor="#009933" AutoPostBack="true" OnSelectedIndexChanged="ListUpSysTree_SelectedIndexChanged">
                                        </asp:ListBox>
                                    </td>
                                </tr>
                                <tr align="center">
                                    <td>&nbsp;</td>
                                    <td><font color="red"><b>下游系統功能</b></font></td>
                                </tr>
                                <tr align="center">
                                    <td align="center">
                                        <asp:LinkButton ID="LinkDnSysTreeAdd" runat="server" Font-Size="Small" OnClick="LinkDnSysTreeAdd_Click"
                                            Text=">> 新增 >>" ToolTip="新增左方所選取之系統到右方下游"></asp:LinkButton><br />
                                        <br /><br />
                                        <asp:LinkButton ID="LinkDnSysTreeDel" runat="server" Font-Size="Small" OnClick="LinkDnSysTreeDel_Click"
                                            Text="<< 刪除 <<" ToolTip="刪除右方下游所選取之系統迴路"></asp:LinkButton>
                                    </td>
                                    <td>
                                        <asp:ListBox ID="ListDnSysTree" runat="server" ForeColor="#009933" AutoPostBack="true" OnSelectedIndexChanged="ListDnSysTree_SelectedIndexChanged">
                                        </asp:ListBox>
                                    </td>
                                </tr>
                                <tr><td colspan="3"><asp:Label ID="lblSysTree" runat="server" Font-Size="Small"></asp:Label></td></tr>

                                <tr align="center">
                                    <td>
                                        <asp:DropDownList ID="SelSys" runat="server" AutoPostBack="True" ForeColor="#009933" DataSourceID="SqlDataSourceSys" DataTextField="系統全名" DataValueField="資源編號">
                                        </asp:DropDownList>
                                        <asp:SqlDataSource ID="SqlDataSourceSys" runat="server" ConnectionString="<%$ ConnectionStrings:IDMSConnectionString %>"
                                            SelectCommand="SELECT [系統全名],[資源編號] FROM [View_系統資源] where [資源種類]='系統' order by [系統全名]">
                                        </asp:SqlDataSource>
                                    </td>
                                    <td width="80px" align="center">執行</td>
                                    <td>
                                        <font color="red"><b>功能相同的主機</b></font>
                                        (<asp:LinkButton ID="LinkHostHelp" runat="server" Font-Size="Small" OnClick="LinkHostHelp_Click" Text="說明" ToolTip="功能主機的設定說明"></asp:LinkButton>)
                                    </td>
                                </tr>
                                <tr align="center">
                                    <td>                                
                                        <asp:ListBox ID="ListHosts1" runat="server" ForeColor="#009933" DataSourceID="SqlDataSourceHosts" AutoPostBack="true" OnSelectedIndexChanged="ListHosts1_SelectedIndexChanged"
                                            DataTextField="主機名稱" DataValueField="作業編號" AppendDataBoundItems="False" Rows="10">
                                        </asp:ListBox>
                                        <asp:SqlDataSource ID="SqlDataSourceHosts" runat="server" ConnectionString="<%$ ConnectionStrings:IDMSConnectionString %>"
                                            SelectCommand="SELECT [主機名稱],[作業編號] FROM [作業主機] where [系統編號]=@Sys order by [主機名稱]">
                                            <SelectParameters>
                                                <asp:ControlParameter ControlID="SelSys" Name="Sys" PropertyName="SelectedValue" Type="Int32" />
                                            </SelectParameters>
                                        </asp:SqlDataSource>
                                    </td>
                                    <td align="center">
                                        <asp:LinkButton ID="LinkHostsAdd"   runat="server" Font-Size="Small" OnClick="LinkHostsAdd_Click"   Text=">> 新增 >>" ToolTip="新增左方所選取之功能主機到右方"></asp:LinkButton><br/><br/><br/>
                                        <asp:LinkButton ID="LinkHostsDel"   runat="server" Font-Size="Small" OnClick="LinkHostsDel_Click"   Text="<< 刪除 <<" ToolTip="刪除右方所選取之功能主機"></asp:LinkButton><br/><br/><br/>                        
                                        <asp:LinkButton ID="LinkHostsEdit"  runat="server" Font-Size="Small" OnClick="LinkHostsEdit_Click"  Text=">> 編輯 >>" ToolTip="編輯右方所選取之功能主機"></asp:LinkButton>
                                    </td>
                                    <td>
                                        <asp:ListBox ID="ListHosts2" runat="server" ForeColor="#009933" Rows="10" AutoPostBack="true" OnSelectedIndexChanged="ListHosts2_SelectedIndexChanged"></asp:ListBox>                             
                                    </td>
                                </tr>
                                <tr><td colspan="3"><asp:Label ID="lblHosts" runat="server" Font-Size="Small"></asp:Label></td></tr>
                            </table>                            
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
    </table>
</asp:Content>
