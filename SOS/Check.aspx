<%@ Page Title="機器清查" Language="C#" MasterPageFile="~/IDMS/MasterPage.master" AutoEventWireup="true" Debug="true"
    MaintainScrollPositionOnPostback="true" CodeFile="Check.aspx.cs" Inherits="SOS_Check"
    %>

<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="Server">
    <style type="text/css">
        .style4
        {
            font-size: x-small;
            color: Blue;
            cursor: pointer;
        }
        
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
            text-align: center;
            font-weight: bold;
        }
    </style>
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1" runat="Server">    
    <br />
    <asp:CheckBox ID="ChkYear" Checked="true" runat="server" />
    <asp:Label ID="Label1" runat="server" Font-Size="Small" Text="清查年度"></asp:Label>
    <asp:DropDownList ID="SelChkYear" runat="server" ForeColor="#009933" Visible="true">
    </asp:DropDownList>&nbsp;&nbsp;&nbsp;&nbsp;
    
    <asp:Label ID="Label5" runat="server" Font-Size="Small" Text="確認碼(清查年度)："></asp:Label>
    <asp:TextBox ID="TextConfirm" runat="server" Columns="4" CssClass="style7"></asp:TextBox>&nbsp;&nbsp;
    <asp:Button ID="BtnChkList" runat="server" Text="產　生　或　同　步　清　查　清　單" OnClick="BtnChkList_Click" />&nbsp;&nbsp; 
        <span class="style4" onclick="alert('本功能將執行下列動作：\n\n1.產生本年度的機器清查清單 \n2.將清單項目與實體設備同步(無刪有增) \n\n清查範圍可參考：[新增資料]按鈕->機器清查編輯介面->清查年度說明')">
            <u style="cursor:pointer">說明</u>
        </span> &nbsp;&nbsp;&nbsp;&nbsp;
    <asp:Button ID="BtnChkDel" runat="server" Text="刪　除　清　單　全　部　資　料" OnClick="BtnChkDel_Click" />
    <br />
    <asp:Label ID="Label10" runat="server" Text="清查流程：產生年度清查清單 -> 實際清查 -> 記錄不符狀況、相符資料結案 -> 未結案的不符資料經相關人員確認 -> 更正實體設備資料 -> 清單設備全部結案" Font-Size="Small" ForeColor="Red"></asp:Label>    
    <br />
    <asp:Label ID="Label11" runat="server" Text="清查範圍：門禁管制區(前、一、二、HPC、外圍)的接電設備(系統、網路、環境、週邊)" Font-Size="Small" ForeColor="Red"></asp:Label> 

    <table cellpadding="0" cellspacing="0" class="style6" border="1">
        <tr>
            <td class="style7"> 設備 </td>
            <td class="style7"> 電力 </td>
            <td class="style7"> 維護 </td>
            <td class="style7"> 保管 </td>
            <td class="style7"> 執行人員 </td>
            <td class="style7"> 狀態 </td>
        </tr>
        <tr>
            <td>
                <asp:CheckBox ID="ChkType" runat="server" />
                <asp:Label ID="Label8" runat="server" Font-Size="Small" Text="型態"></asp:Label>
                <asp:DropDownList ID="SelType" runat="server" ForeColor="#009933" Visible="true"
                    AutoPostBack="True" DataSourceID="SqlDataSourceType" DataTextField="Item" DataValueField="Item">
                </asp:DropDownList>
                <asp:SqlDataSource ID="SqlDataSourceType" runat="server" ConnectionString="<%$ ConnectionStrings:IDMSConnectionString %>"
                    SelectCommand="select Item from Config where Kind=@Kind order by mark">
                    <SelectParameters>
                        <asp:Parameter DefaultValue="設備型態" Name="Kind" Type="String" />
                    </SelectParameters>
                </asp:SqlDataSource>
                <asp:CheckBox ID="ChkKind" runat="server" />
                <asp:Label ID="Label3" runat="server" Font-Size="Small" Text="種類"></asp:Label>
                <asp:DropDownList ID="SelKind" runat="server" ForeColor="#009933" Visible="true"
                    DataSourceID="SqlDataSourceKind" DataTextField="Item" DataValueField="Item">
                </asp:DropDownList>
                <asp:SqlDataSource ID="SqlDataSourceKind" runat="server" ConnectionString="<%$ ConnectionStrings:IDMSConnectionString %>"
                    SelectCommand="select Item from Config where Kind=@Kind order by mark">
                    <SelectParameters>
                        <asp:ControlParameter ControlID="SelType" Name="Kind" PropertyName="SelectedValue"
                            Type="String" />
                    </SelectParameters>
                </asp:SqlDataSource>
                <br />
                <asp:CheckBox ID="ChkPlace" runat="server" />
                <asp:Label ID="Label2" runat="server" Font-Size="Small" Text="地點"></asp:Label>
                <asp:DropDownList ID="SelPlace" runat="server" DataSourceID="SqlDataSourcePlace"
                    DataTextField="區域名稱" ForeColor="#009933" DataValueField="區域名稱" AutoPostBack="True"
                    OnSelectedIndexChanged="SelPlace_SelectedIndexChanged">
                </asp:DropDownList>
                <asp:SqlDataSource ID="SqlDataSourcePlace" runat="server" ConnectionString="<%$ ConnectionStrings:IDMSConnectionString %>"
                    SelectCommand="SELECT [Item] as [區域名稱] from [Config] where [Kind]='門禁管制區' ORDER BY [mark]">
                </asp:SqlDataSource>
                <asp:CheckBox ID="ChkPointer" runat="server" />
                <asp:Label ID="Label4" runat="server" Font-Size="Small" Text="定位"></asp:Label>
                <asp:DropDownList ID="SelPointer" runat="server" DataSourceID="SqlDataSourcePointer"
                    DataTextField="定位名稱" ForeColor="#009933" DataValueField="定位編號">
                </asp:DropDownList>
                <asp:SqlDataSource ID="SqlDataSourcePointer" runat="server" ConnectionString="<%$ ConnectionStrings:IDMSConnectionString %>"
                    SelectCommand="SELECT [定位名稱], [定位編號] FROM [定位設定] WHERE (([定位方式] &lt;&gt; @定位方式) AND ([區域名稱] = @區域名稱)) ORDER BY [定位名稱]">
                    <SelectParameters>
                        <asp:Parameter DefaultValue="坐標" Name="定位方式" Type="String" />
                        <asp:ControlParameter ControlID="SelPlace" Name="區域名稱" PropertyName="SelectedValue"
                            Type="String" />
                    </SelectParameters>
                </asp:SqlDataSource>
            </td>
            <td>
                <asp:CheckBox ID="ChkCircuit" runat="server" />
                <asp:DropDownList ID="SelPanel" runat="server" ForeColor="#009933" Visible="true"
                    AutoPostBack="True" DataSourceID="SqlDataSourcePanel" DataTextField="設備名稱" DataValueField="設備編號"
                    OnSelectedIndexChanged="SelPanel_SelectedIndexChanged" AppendDataBoundItems="True">
                </asp:DropDownList>
                <asp:SqlDataSource ID="SqlDataSourcePanel" runat="server" ConnectionString="<%$ ConnectionStrings:IDMSConnectionString %>"
                    SelectCommand="SELECT [設備編號],[設備名稱] FROM [實體設備] where [設備種類]='配電盤' or [設備種類]='PDC' or [設備種類]='電源' ORDER BY [設備種類] desc,[設備名稱]">
                </asp:SqlDataSource>
                &nbsp;
                <asp:DropDownList ID="SelCircuit" runat="server" ForeColor="#009933" Visible="true"
                    DataSourceID="SqlDataSourceCircuit" DataTextField="設備名稱" DataValueField="設備編號">
                </asp:DropDownList>
                <asp:SqlDataSource ID="SqlDataSourceCircuit" runat="server" ConnectionString="<%$ ConnectionStrings:IDMSConnectionString %>"
                    SelectCommand=""></asp:SqlDataSource>
                <br />
                <asp:CheckBox ID="ChkPowerNum" runat="server" />
                <asp:DropDownList ID="SelPowerNum" runat="server" ForeColor="#009933" Visible="true">
                    <asp:ListItem Value="0">未接迴路</asp:ListItem>
                    <asp:ListItem Value="1">接單迴路</asp:ListItem>
                    <asp:ListItem Value="2">接雙迴路</asp:ListItem>
                    <asp:ListItem Value="*">接多迴路</asp:ListItem>
                </asp:DropDownList>
            </td>
            <td>
                <asp:CheckBox ID="ChkHwUnit" runat="server" />課別
                <asp:DropDownList ID="SelHwUnit" runat="server" ForeColor="#009933" Visible="true"
                    AutoPostBack="True" DataSourceID="SqlDataSourceUnit" DataTextField="成員" DataValueField="成員">
                </asp:DropDownList>
                <asp:SqlDataSource ID="SqlDataSourceUnit" runat="server" ConnectionString="<%$ ConnectionStrings:IDMSConnectionString %>"
                    SelectCommand="Select [成員] from [View_組織架構] where [性質]='課別' order by [代號]"></asp:SqlDataSource>
                <br />
                <asp:CheckBox ID="ChkHw" runat="server" />人員
                <asp:DropDownList ID="SelHw" runat="server" ForeColor="#009933" Visible="true" DataSourceID="SqlDataSourceHw"
                    DataTextField="Item" DataValueField="Item">
                </asp:DropDownList>
                <asp:SqlDataSource ID="SqlDataSourceHw" runat="server" ConnectionString="<%$ ConnectionStrings:IDMSConnectionString %>"
                    SelectCommand="Select Item from Config where Kind=@課別 order by Mark">
                    <SelectParameters>
                        <asp:ControlParameter ControlID="SelHwUnit" Name="課別" PropertyName="SelectedValue"
                            Type="String" />
                    </SelectParameters>
                </asp:SqlDataSource>
            </td>
            <td>
                <asp:CheckBox ID="ChkKeepUnit" runat="server" />課別
                <asp:DropDownList ID="SelKeepUnit" runat="server" ForeColor="#009933" Visible="true"
                    OnSelectedIndexChanged="SelKeepUnit_SelectedIndexChanged" AutoPostBack="True"
                    DataSourceID="SqlDataSourceUnit" DataTextField="成員" DataValueField="成員">
                </asp:DropDownList>
                <br />
                <asp:CheckBox ID="ChkKeeper" runat="server" />人員
                <asp:DropDownList ID="SelKeeper" runat="server" ForeColor="#009933" Visible="true"
                    DataSourceID="SqlDataSourceKeeper" DataTextField="Item" DataValueField="Item">
                </asp:DropDownList>
                <asp:SqlDataSource ID="SqlDataSourceKeeper" runat="server" ConnectionString="<%$ ConnectionStrings:IDMSConnectionString %>"
                    SelectCommand="Select Item from Config where Kind=@課別 order by Mark">
                    <SelectParameters>
                        <asp:ControlParameter ControlID="SelKeepUnit" Name="課別" PropertyName="SelectedValue"
                            Type="String" />
                    </SelectParameters>
                </asp:SqlDataSource>
            </td>
            <td>
                <asp:CheckBox ID="Chker" runat="server" />清查
                <asp:DropDownList ID="SelChker" runat="server" ForeColor="#009933" Visible="true"
                    DataSourceID="SqlDataSourceEvner" DataTextField="Item" DataValueField="Item">
                </asp:DropDownList>
                <br />
                <asp:CheckBox ID="ChkUpdater" runat="server" />更新
                <asp:DropDownList ID="SelUpdater" runat="server" ForeColor="#009933" Visible="true"
                    DataSourceID="SqlDataSourceEvner" DataTextField="Item" DataValueField="Item">
                </asp:DropDownList>
                <asp:SqlDataSource ID="SqlDataSourceEvner" runat="server" ConnectionString="<%$ ConnectionStrings:IDMSConnectionString %>"
                    SelectCommand="Select Item from Config where Kind='作業管控科' order by Mark"></asp:SqlDataSource>
            </td>

            <td>
                <asp:CheckBox ID="ChkDevStatus" runat="server" />設備
                <asp:DropDownList ID="SelDevStatus" runat="server" ForeColor="#009933" DataSourceID="SqlDataSourceDevStatus"
                    DataTextField="Item" DataValueField="Item">
                </asp:DropDownList>
                <asp:SqlDataSource ID="SqlDataSourceDevStatus" runat="server" ConnectionString="<%$ ConnectionStrings:IDMSConnectionString %>"
                    SelectCommand="SELECT [Item] FROM [Config] WHERE [Kind] = '設備狀態' ORDER BY [mark]">
                </asp:SqlDataSource>
                <br />

                <asp:CheckBox ID="ChkStatus" runat="server" />清查
                <asp:DropDownList ID="SelStatus" runat="server" ForeColor="#009933" DataSourceID="SqlDataSourceStatus"
                    DataTextField="Item" DataValueField="Item">
                </asp:DropDownList>
                <asp:SqlDataSource ID="SqlDataSourceStatus" runat="server" ConnectionString="<%$ ConnectionStrings:IDMSConnectionString %>"
                    SelectCommand="SELECT [Item] FROM [Config] WHERE [Kind] = '清查狀態' ORDER BY [mark]">
                </asp:SqlDataSource>                
            </td>
        </tr>
        <tr>
            <td colspan="5">
                <asp:CheckBox ID="ChkCheck" runat="server" />
                <asp:Label ID="Label9" runat="server" Font-Size="Small" Text="結果："></asp:Label>
                <asp:CheckBoxList ID="ChkChecks" runat="server" DataSourceID="SqlDataSourceCheck" Font-Size="Small"
                    DataTextField='Item' DataValueField="Item" RepeatDirection="Horizontal" RepeatLayout="Flow">
                </asp:CheckBoxList>
                <asp:SqlDataSource ID="SqlDataSourceCheck" runat="server" ConnectionString="<%$ ConnectionStrings:IDMSConnectionString %>"
                    SelectCommand="SELECT [Item] FROM [Config] WHERE ([Kind] = @Kind) ORDER BY [mark]">
                    <SelectParameters>
                        <asp:Parameter DefaultValue="清查結果" Name="Kind" Type="String" />
                    </SelectParameters>
                </asp:SqlDataSource>&nbsp;&nbsp;
                <span class="style5">(系統設定可改變項目)</span>
            </td> 
            
            <td>
                <asp:CheckBox ID="ChkAdd" runat="server" />僅新增資料
            </td>            
        </tr>
        <tr>
            <td colspan="6">
                <asp:CheckBox ID="ChkSQL" runat="server" />
                <asp:Label ID="Label7" runat="server" Font-Size="Small" Text="SQL"></asp:Label>
                <asp:TextBox ID="TextSQL" Width="450px" runat="server">有特殊查詢需求，請洽機房 (本段文字僅供說明，請勿執行查詢)</asp:TextBox>
                <asp:DropDownList ID="SelSQL" runat="server" ForeColor="#009933" Visible="true" OnSelectedIndexChanged="SelSQL_SelectedIndexChanged"
                    AutoPostBack="True" DataSourceID="SqlDataSourceSQL" DataTextField="Item" DataValueField="Memo">
                </asp:DropDownList>
                &nbsp;
                <asp:SqlDataSource ID="SqlDataSourceSQL" runat="server" ConnectionString="<%$ ConnectionStrings:IDMSConnectionString %>"
                    SelectCommand="SELECT [Item],[Memo] from Config where Kind='SQL範本' and Config='清查查詢' order by Mark">
                </asp:SqlDataSource>
                &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                <asp:CheckBox ID="ChkKey" runat="server" />
                <asp:Label ID="Label12" runat="server" Font-Size="Small" Text="關鍵字"></asp:Label>
                <asp:TextBox ID="TextKey" runat="server"></asp:TextBox>&nbsp;<span class="style5">(逗點隔開)</span>&nbsp;&nbsp;
                <asp:CheckBox ID="ChkPage" Checked="true" runat="server" />
                <asp:Label ID="Label6" runat="server" Font-Size="Small" Text="分頁"></asp:Label>&nbsp;&nbsp;&nbsp;&nbsp;
                顯示欄位：<asp:DropDownList ID="SelShow" runat="server" ForeColor="#009933" Visible="true"
                    DataSourceID="SqlDataSourceShow" DataTextField="Item" DataValueField="Memo">
                </asp:DropDownList>
                &nbsp;
                <asp:SqlDataSource ID="SqlDataSourceShow" runat="server" ConnectionString="<%$ ConnectionStrings:IDMSConnectionString %>"
                    SelectCommand="SELECT [Item],[Memo] from Config where Kind='清查查詢' order by Mark">
                </asp:SqlDataSource>
                <span class="style4" onclick="alert('顯示欄位如需客製化設定，請洽機房');"><u style="cursor: pointer">說明</u> </span>
            </td>
        </tr>
    </table>
    <table width="100%">
        <tr>
            <td>
                <asp:Button ID="BtnAddChk" runat="server" Text="　新　增　資　料　" OnClick="BtnAddChk_Click" />
            </td>
            <td align="right">                
                <asp:Button ID="Button1" runat="server" Text="　查　詢　" OnClick="BtnSearch_Click" />
            </td>
        </tr>
    </table>

    <p>
        <asp:GridView ID="GridView1" runat="server" AllowPaging="True" AllowSorting="True"
            BackColor="#CCCCCC" BorderColor="#999999" BorderStyle="Solid" BorderWidth="3px"
            CellPadding="4" CellSpacing="2" DataSourceID="SqlDataSourceSearch" DataKeyNames="設備編號"
            ForeColor="Black" Font-Size="Small" OnSelectedIndexChanging="GridView1_SelectedIndexChanging" OnRowCommand="GridView1_RowCommand">
            <Columns>
                <asp:ButtonField ButtonType="Button" CommandName="ChkDev" Text="設備" HeaderText="設備" />
                <asp:ButtonField ButtonType="Button" CommandName="ChkXX" Text="不符註記" HeaderText="清查不符" />
                <asp:ButtonField ButtonType="Button" CommandName="ChkOK" Text="予以結案" HeaderText="清查符合"  />
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
        <asp:SqlDataSource ID="SqlDataSourceSearch" runat="server" ConnectionString="<%$ ConnectionStrings:IDMSConnectionString %>"
            SelectCommand=""></asp:SqlDataSource>
        <br />
    </p>

    <div class="style7">
        總筆數：<asp:TextBox ID="TextCount" runat="server" BackColor="#CCCCCC" ReadOnly="True"
            Width="60px" CssClass="style7" Font-Bold="True"></asp:TextBox>&nbsp;&nbsp; 總價：<asp:TextBox
                ID="TextTotal" runat="server" BackColor="#CCCCCC" ReadOnly="True" Width="120px"
                CssClass="style7" Font-Bold="True"></asp:TextBox>&nbsp;&nbsp; 總額定電流：<asp:TextBox
                    ID="TextSumCurrent" runat="server" BackColor="#CCCCCC" ReadOnly="True" Width="60px"
                    CssClass="style7" Font-Bold="True">0</asp:TextBox>&nbsp;&nbsp;&nbsp;&nbsp;
        <asp:Button ID="BtnExcel" runat="server" Text="匯出Excel檔 (html格式)" OnClick="BtnExcel_Click" />
    </div>
    <span class="style5">設定顯示欄位：[系統設定] -> [顯示欄位] ，所有可設定的欄位名稱在[系統設定] -> [一般設定] -> [顯示欄位] 可查到</span>
</asp:Content>
