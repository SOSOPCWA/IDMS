<%@ Page Title="系統資源" Language="C#" MasterPageFile="~/IDMS/MasterPage.master" AutoEventWireup="true" CodeFile="SysEdit.aspx.cs" Inherits="AP_SysEdit" MaintainScrollPositionOnPostback="true" Debug="true" %>

<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="Server">
    <style type="text/css">
        .style1
        {
            text-align: center;
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
        }
        .style20
        {
            font-size: small;
        }
        .style21
        {
            font-size: xx-large;
            font-family: 標楷體;
            width: 120px;
            text-align: center;
        }
        .style22
        {
            font-size: xx-large;
            font-family: 標楷體;
            text-align: center;
            width: 320px;
        }
        .style23
        {
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
    <asp:Table ID="Table1" runat="server" Width="100%" GridLines="Both">
        <asp:TableRow ID="rowTitle" runat="server">
            <asp:TableCell runat="server" Width="280px" RowSpan="15" VerticalAlign="Top">
                <asp:TreeView ID="TreeView1" runat="server" ImageSet="XPFileExplorer" NodeWrap="True"
                    OnSelectedNodeChanged="TreeView1_SelectedNodeChanged" NodeIndent="15" OnTreeNodePopulate="TreeView1_TreeNodePopulate"
                    ShowLines="True" ExpandDepth="1" OnTreeNodeCollapsed="TreeView1_TreeNodeCollapsed"
                    OnTreeNodeExpanded="TreeView1_TreeNodeExpanded">
                    <HoverNodeStyle Font-Underline="True" ForeColor="#6666AA" />
                    <Nodes>
                        <asp:TreeNode PopulateOnDemand="True" SelectAction="SelectExpand" Text="<font color='blue' size='large'><b>系統資源</b></font>"
                            ToolTip="0. 請對照機房教育訓練系統的系統架構圖" Value="0"></asp:TreeNode>
                    </Nodes>
                    <NodeStyle Font-Names="Tahoma" Font-Size="10pt" ForeColor="Black" HorizontalPadding="2px"
                        NodeSpacing="0px" VerticalPadding="2px" />
                    <ParentNodeStyle Font-Bold="False" />
                    <SelectedNodeStyle BackColor="#B5B5B5" Font-Underline="False" HorizontalPadding="0px"
                        VerticalPadding="0px" ForeColor="Red" />
                </asp:TreeView>
            </asp:TableCell>
            <asp:TableCell runat="server" CssClass="style21">欄位</asp:TableCell>
            <asp:TableCell runat="server" CssClass="style22">
                系統資源 <asp:LinkButton ID="LinkHideHelp" runat="server" Font-Size="Small" ForeColor="Blue" OnClick="LinkHideHelp_Click" ToolTip="隱藏右方說明欄位，以簡化操作介面">隱藏說明</asp:LinkButton>
            </asp:TableCell>
            <asp:TableCell runat="server" CssClass="style18">設定說明</asp:TableCell>
        </asp:TableRow>

        <asp:TableRow runat="server">
            <asp:TableCell runat="server" CssClass="style19" ForeColor="Red">資源編號</asp:TableCell>
            <asp:TableCell runat="server" CssClass="style23">
                <asp:TextBox ID="TextSysNo" runat="server" Columns="4" ReadOnly="True" BackColor="Silver"
                    Font-Bold="True" Font-Size="Large" CssClass="style1" ForeColor="Red"></asp:TextBox>
            </asp:TableCell>
            <asp:TableCell runat="server" CssClass="style20">
                <asp:Label ID="helpSysNo" runat="server" Font-Size="Small"></asp:Label>
            </asp:TableCell>
        </asp:TableRow>

        <asp:TableRow runat="server">
            <asp:TableCell runat="server" CssClass="style19">
                資源種類(★)
            </asp:TableCell>
            <asp:TableCell runat="server" CssClass="style23">
                <asp:DropDownList ID="SelKind" runat="server" AutoPostBack="True" ForeColor="#009933"
                    OnSelectedIndexChanged="SelKind_SelectedIndexChanged" DataSourceID="SqlDataSourceKind"
                    DataTextField="Item" DataValueField="Item">
                </asp:DropDownList>
                &nbsp;&nbsp;
                <asp:Label ID="lblKind" runat="server" Font-Size="Small"></asp:Label>
                <asp:SqlDataSource ID="SqlDataSourceKind" runat="server" ConnectionString="<%$ ConnectionStrings:IDMSConnectionString %>"
                    SelectCommand="SELECT [Item] FROM [Config] WHERE ([Kind] = '資源種類') ORDER BY [mark]">
                </asp:SqlDataSource>
            </asp:TableCell>
            <asp:TableCell runat="server" CssClass="style20">
                <asp:Label ID="helpKind" runat="server" Font-Size="Small"></asp:Label>
            </asp:TableCell>
        </asp:TableRow>

        <asp:TableRow runat="server">
            <asp:TableCell runat="server" CssClass="style19">
                資源名稱(★)
            </asp:TableCell>
            <asp:TableCell runat="server" CssClass="style23">
                <asp:TextBox ID="TextSysName" runat="server" Width="300px" ForeColor="Green"></asp:TextBox>
            </asp:TableCell>
            <asp:TableCell runat="server" CssClass="style20">
                <asp:Label ID="helpSysName" runat="server" Font-Size="Small"></asp:Label>
            </asp:TableCell>
        </asp:TableRow>

        <asp:TableRow runat="server">
            <asp:TableCell runat="server" CssClass="style19">
                資源功能(★)
            </asp:TableCell>
            <asp:TableCell runat="server" CssClass="style23">
                <asp:TextBox ID="TextFunc" runat="server" Columns="48" Width="300px" Height="100px" Rows="5" TextMode="MultiLine" ForeColor="Green"></asp:TextBox>
            </asp:TableCell>
            <asp:TableCell runat="server" CssClass="style20">
                <asp:Label ID="helpFunc" runat="server" Font-Size="Small"></asp:Label>
            </asp:TableCell>
        </asp:TableRow>      

        <asp:TableRow runat="server">
            <asp:TableCell runat="server" CssClass="style19">
                <font color="red">系統架構(★)</font>
            </asp:TableCell>
            <asp:TableCell runat="server" CssClass="style23">上層所屬：
                <asp:DropDownList ID="SelBelong" runat="server" AutoPostBack="True" AppendDataBoundItems="true" 
                    ForeColor="#009933" DataSourceID="SqlDataSourceBelong" DataTextField="資源名稱" DataValueField="資源編號">                    
                    <asp:ListItem Value="0">(系統資源根目錄)</asp:ListItem>
                </asp:DropDownList>
                <asp:SqlDataSource ID="SqlDataSourceBelong" runat="server" ConnectionString="<%$ ConnectionStrings:IDMSConnectionString %>" SelectCommand="">
                </asp:SqlDataSource>
            </asp:TableCell>
            <asp:TableCell runat="server" CssClass="style20">
                <asp:Label ID="helpBelong" runat="server" Font-Size="Small"></asp:Label>
            </asp:TableCell>
        </asp:TableRow>

        <asp:TableRow runat="server">
            <asp:TableCell runat="server" CssClass="style19">
                <font color="red">系統功能(★)</font>
            </asp:TableCell>
            <asp:TableCell runat="server" CssClass="style23">
                功能主機：<asp:Label ID="lblHosts" runat="server" Font-Size="Small" Text="(無)" ForeColor="green"></asp:Label>
                <br /><br />
                <table><tr><td>
                    <asp:Button ID="BtnSysFun" runat="server" Text=" 設 定 " CausesValidation="False" OnClick="BtnSysTree_Click" />&nbsp;&nbsp;
                    <asp:Label ID="Label1" runat="server" Font-Size="Small" Text="(設定路徑同系統迴路)" ForeColor="blue"></asp:Label>
                </td></tr></table> 
                <br />
                系統名稱主機：<asp:Label ID="lblSysNames" runat="server" Font-Size="Small" Text="0 筆 (異動請至作業主機介面)" ForeColor="green" ToolTip="資源種類是資訊系統才有"></asp:Label>               
            </asp:TableCell>
            <asp:TableCell runat="server" CssClass="style20">
                <asp:Label ID="helpHosts" runat="server" Font-Size="Small"></asp:Label>
            </asp:TableCell>
        </asp:TableRow>

        <asp:TableRow runat="server">
            <asp:TableCell runat="server" CssClass="style19">
                <font color="red">系統迴路(★)</font>
            </asp:TableCell>
            <asp:TableCell runat="server" CssClass="style20">
                <table>
                    <tr>
                        <td><asp:Button ID="BtnSysTree" runat="server" Text=" 設 定 " CausesValidation="False" OnClick="BtnSysTree_Click" />&nbsp;&nbsp;&nbsp;&nbsp;</td>
                        <td>
                            上游：<asp:Label ID="lblUpSysTree" runat="server" Font-Size="Small" ForeColor="green"></asp:Label>
                            <br />
                            下游：<asp:Label ID="lblDnSysTree" runat="server" Font-Size="Small" ForeColor="green"></asp:Label>
                        </td>
                    </tr>
                </table>
            </asp:TableCell>
            <asp:TableCell runat="server" CssClass="style20">
                <asp:Label ID="helpSysTree" runat="server" Font-Size="Small"></asp:Label>
            </asp:TableCell>
        </asp:TableRow>      

        <asp:TableRow runat="server">
            <asp:TableCell runat="server" CssClass="style19">
                資產編號
            </asp:TableCell>
            <asp:TableCell runat="server" CssClass="style23">
                <asp:DropDownList ID="SelAssets" runat="server" ForeColor="#009933" AutoPostBack="True" OnSelectedIndexChanged="SelAssets_SelectedIndexChanged">
                </asp:DropDownList><br /><br />
                <asp:Label ID="lblAssets" runat="server" Font-Size="Small"></asp:Label>
            </asp:TableCell>
            <asp:TableCell runat="server" CssClass="style20">
                <asp:Label ID="helpAssetsNo" runat="server" Font-Size="Small"></asp:Label>
            </asp:TableCell>
        </asp:TableRow>
        
        <asp:TableRow runat="server">
            <asp:TableCell runat="server" CssClass="style19">
                維護人員(★)
            </asp:TableCell>
            <asp:TableCell runat="server" CssClass="style23">
                <asp:DropDownList ID="SelMtUnit" runat="server" AutoPostBack="True" ForeColor="#009933"
                    DataSourceID="SqlDataSourceMtUnit" DataTextField="Item" DataValueField="Item">
                </asp:DropDownList>
                <asp:SqlDataSource ID="SqlDataSourceMtUnit" runat="server" ConnectionString="<%$ ConnectionStrings:IDMSConnectionString %>"
                    SelectCommand="SELECT [Item] FROM [Config] WHERE (Kind='組織架構' and Item=@MIC) or ([Kind] = @Kind or Kind=@MIC) ORDER BY [mark]">
                    <SelectParameters>
                        <asp:Parameter DefaultValue="維護群組" Name="Kind" Type="String" />
                        <asp:Parameter DefaultValue="資訊中心" Name="MIC" Type="String" />
                    </SelectParameters>
                </asp:SqlDataSource>
                &nbsp;
                <asp:DropDownList ID="SelMt" runat="server" AutoPostBack="True" ForeColor="#009933"
                    DataSourceID="SqlDataSourceMt" DataTextField="Item" DataValueField="Item">
                </asp:DropDownList>
                <asp:SqlDataSource ID="SqlDataSourceMt" runat="server" ConnectionString="<%$ ConnectionStrings:IDMSConnectionString %>"
                    SelectCommand="SELECT [Item] FROM [Config] WHERE ([Kind] = @Kind) and [Config]<>'測試' ORDER BY [mark]">
                    <SelectParameters>
                        <asp:ControlParameter ControlID="SelMtUnit" Name="Kind" PropertyName="SelectedValue"
                            Type="String" />
                    </SelectParameters>
                </asp:SqlDataSource>
                &nbsp;
                <asp:Label ID="lblMt" runat="server" Font-Size="Small"></asp:Label>
            </asp:TableCell>
            <asp:TableCell runat="server" CssClass="style20">
                <asp:Label ID="helpMt" runat="server" Font-Size="Small"></asp:Label>
            </asp:TableCell>
        </asp:TableRow>

        <asp:TableRow runat="server">
            <asp:TableCell runat="server" CssClass="style19">
                備註說明
            </asp:TableCell>
            <asp:TableCell runat="server" CssClass="style23">
                <asp:TextBox ID="TextMemo" runat="server" Columns="48" Width="300px" Height="100px" ForeColor="Green"
                    Rows="5" TextMode="MultiLine"></asp:TextBox>
            </asp:TableCell>
            <asp:TableCell runat="server" CssClass="style20">
                <asp:Label ID="helpMemo" runat="server" Font-Size="Small"></asp:Label>
            </asp:TableCell>
        </asp:TableRow>

        <asp:TableRow runat="server">
            <asp:TableCell runat="server" CssClass="style19">
                資料建立
            </asp:TableCell>
            <asp:TableCell runat="server" CssClass="style23">
                <asp:Label ID="LblCreater" runat="server" Font-Size="Small"></asp:Label>
                &nbsp;
                <asp:Label ID="LblCreateDT" runat="server" Font-Size="Small"></asp:Label>
            </asp:TableCell>
            <asp:TableCell runat="server" CssClass="style20">
                <asp:Label ID="helpCreate" runat="server" Font-Size="Small"></asp:Label>
            </asp:TableCell>
        </asp:TableRow>

        <asp:TableRow runat="server">
            <asp:TableCell runat="server" CssClass="style19">
                資料修改
            </asp:TableCell>
            <asp:TableCell runat="server" CssClass="style23">
                <asp:Label ID="LblUpdater" runat="server" Font-Size="Small"></asp:Label>
                &nbsp;
                <asp:Label ID="LblUpdateDT" runat="server" Font-Size="Small"></asp:Label>
            </asp:TableCell>
            <asp:TableCell runat="server" CssClass="style20">
                <asp:Label ID="helpUpdate" runat="server" Font-Size="Small"></asp:Label>
            </asp:TableCell>
        </asp:TableRow>

        <asp:TableRow runat="server">
            <asp:TableCell runat="server" CssClass="style19">
                資料查核
            </asp:TableCell>
            <asp:TableCell runat="server" CssClass="style23">
                <asp:Label ID="lblChecks" runat="server" ForeColor="red" Text="" Font-Bold="True"></asp:Label>&nbsp;
            </asp:TableCell>
            <asp:TableCell runat="server" CssClass="style20">
                <asp:Label ID="helpChecks" runat="server" Font-Size="Small"></asp:Label>
            </asp:TableCell>
        </asp:TableRow>

        <asp:TableRow runat="server">
            <asp:TableCell runat="server" ColumnSpan="4" HorizontalAlign="Center">
                <br />
                <asp:Button ID="BtnNew" runat="server" Text="新增" OnClick="BtnNew_Click" />&nbsp;&nbsp;&nbsp;&nbsp;
                <asp:Button ID="BtnEdit" runat="server" Text="修改" OnClick="BtnEdit_Click" />&nbsp;&nbsp;&nbsp;&nbsp;
                <asp:Button ID="BtnDel" runat="server" Text="刪除" OnClientClick="return confirm('刪除前請確認，所有引用到此的系統資源(例如：所屬系統、系統迴路、作業主機)均已先刪除或修改，您確定要刪除這筆資料嗎？')"
                    OnClick="BtnDel_Click" />&nbsp;&nbsp;&nbsp;&nbsp;
                <asp:Button ID="BtnLife" runat="server" Text=" 生命履歷 "   OnClick="BtnLife_Click" CausesValidation="False" />&nbsp;&nbsp;&nbsp;&nbsp;
                <asp:Button ID="BtnBlank" runat="server" Text="空白頁" OnClick="BtnBlank_Click" />
                <br />
                <br />
                <span class="style24">(★：必填欄位　　紅色：連結欄位)　　異動後樹狀節點顯示不會同步更改，請重新連結系統資源網頁以重讀資料</span>
            </asp:TableCell>
        </asp:TableRow>
    </asp:Table>
</asp:Content>
