<%@ Page Title="軟體查詢" Language="C#" MasterPageFile="~/IDMS/MasterPage.master" AutoEventWireup="true"
    MaintainScrollPositionOnPostback="true" CodeFile="Search.aspx.cs" Inherits="Software_Search"
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
    <p style="font-size: small">
        <table cellpadding="0" cellspacing="0" class="style6" border="1">
            <tr>
                <td class="style7">課別人員</td>
                <td>
                    <asp:CheckBox ID="ChkUnitSw" runat="server" />軟體課別
                    <asp:DropDownList ID="SelUnitSw" runat="server" ForeColor="#009933" Visible="true" AutoPostBack="True"
                        DataSourceID="SqlDataSourceUnitSw" DataTextField="成員" DataValueField="成員">
                    </asp:DropDownList>
                    <asp:SqlDataSource ID="SqlDataSourceUnitSw" runat="server" ConnectionString="<%$ ConnectionStrings:IDMSConnectionString %>"
                        SelectCommand="Select [成員] from [View_組織架構] where [性質]='課別'and [單位]='數值資訊組' order by [代號]"></asp:SqlDataSource>                    
                    &nbsp;&nbsp;
                    <asp:CheckBox ID="ChkUserSw" runat="server" />軟體人員
                    <asp:DropDownList ID="SelUserSw" runat="server" ForeColor="#009933" Visible="true"
                        DataSourceID="SqlDataSourceUserSw" DataTextField="Item" DataValueField="Item">
                    </asp:DropDownList>
                    <asp:SqlDataSource ID="SqlDataSourceUserSw" runat="server" ConnectionString="<%$ ConnectionStrings:IDMSConnectionString %>"
                        SelectCommand="Select Item from Config where Kind=@課別 order by Mark">
                        <SelectParameters>
                            <asp:ControlParameter ControlID="SelUnitSw" Name="課別" PropertyName="SelectedValue" Type="String" />
                        </SelectParameters>
                    </asp:SqlDataSource>
                    &nbsp;&nbsp;&nbsp;&nbsp;
                    
                    <asp:CheckBox ID="ChkUnitLcs" runat="server" />授權課別
                    <asp:DropDownList ID="SelUnitLcs" runat="server" ForeColor="#009933" Visible="true" AutoPostBack="True"
                        DataSourceID="SqlDataSourceUnitLcs" DataTextField="成員" DataValueField="成員">
                    </asp:DropDownList>
                    <asp:SqlDataSource ID="SqlDataSourceUnitLcs" runat="server" ConnectionString="<%$ ConnectionStrings:IDMSConnectionString %>"
                        SelectCommand="Select [成員] from [View_組織架構] where [性質]='課別'and [單位]='數值資訊組' order by [代號]"></asp:SqlDataSource>                    
                    &nbsp;&nbsp;
                    <asp:CheckBox ID="ChkUserLcs" runat="server" />授權人員
                    <asp:DropDownList ID="SelUserLcs" runat="server" ForeColor="#009933" Visible="true"
                        DataSourceID="SqlDataSourceUserLcs" DataTextField="Item" DataValueField="Item">
                    </asp:DropDownList>
                    <asp:SqlDataSource ID="SqlDataSourceUserLcs" runat="server" ConnectionString="<%$ ConnectionStrings:IDMSConnectionString %>"
                        SelectCommand="Select Item from Config where Kind=@課別 order by Mark">
                        <SelectParameters>
                            <asp:ControlParameter ControlID="SelUnitLcs" Name="課別" PropertyName="SelectedValue" Type="String" />
                        </SelectParameters>
                    </asp:SqlDataSource>
                    &nbsp;&nbsp;&nbsp;&nbsp;

                    <asp:CheckBox ID="ChkUnitUse" runat="server" />使用課別
                    <asp:DropDownList ID="SelUnitUse" runat="server" ForeColor="#009933" Visible="true" AutoPostBack="True"
                        DataSourceID="SqlDataSourceUnitUse" DataTextField="成員" DataValueField="成員">
                    </asp:DropDownList>
                    <asp:SqlDataSource ID="SqlDataSourceUnitUse" runat="server" ConnectionString="<%$ ConnectionStrings:IDMSConnectionString %>"
                        SelectCommand="Select [成員] from [View_組織架構] where [性質]='課別'and [單位]='數值資訊組' order by [代號]"></asp:SqlDataSource>                    
                    &nbsp;&nbsp;
                    <asp:CheckBox ID="ChkUserUse" runat="server" />使用人員
                    <asp:DropDownList ID="SelUserUse" runat="server" ForeColor="#009933" Visible="true"
                        DataSourceID="SqlDataSourceUserUse" DataTextField="Item" DataValueField="Item">
                    </asp:DropDownList>
                    <asp:SqlDataSource ID="SqlDataSourceUserUse" runat="server" ConnectionString="<%$ ConnectionStrings:IDMSConnectionString %>"
                        SelectCommand="Select Item from Config where Kind=@課別 order by Mark">
                        <SelectParameters>
                            <asp:ControlParameter ControlID="SelUnitUse" Name="課別" PropertyName="SelectedValue" Type="String" />
                        </SelectParameters>
                    </asp:SqlDataSource>&nbsp;&nbsp; 
                    <span class="style4" onclick="alert('非個人作業系統取作業維護人員當使用人員，個人作業系統或沒有系統作業主機則取保管人員當使用人員');">說明</span>
                </td>
            </tr>
            <tr>
                <td class="style7">狀態</td>
                <td>
                    <asp:CheckBox ID="ChkLcsBuy" runat="server" />購買授權
                    <asp:DropDownList ID="SelLcsBuy" runat="server" ForeColor="#009933" 
                        DataSourceID="SqlDataSourceLcsKind" DataTextField="Item" DataValueField="Item">
                    </asp:DropDownList>
                    <asp:SqlDataSource ID="SqlDataSourceLcsKind" runat="server" ConnectionString="<%$ ConnectionStrings:IDMSConnectionString %>"
                        SelectCommand="SELECT [Item] FROM [Config] WHERE [Kind] = '授權方式' ORDER BY [mark]">
                    </asp:SqlDataSource>
                    &nbsp;&nbsp;&nbsp;&nbsp;
                    <asp:CheckBox ID="ChkLcsAsk" runat="server" />申請授權
                    <asp:DropDownList ID="SelLcsAsk" runat="server" ForeColor="#009933"
                        DataSourceID="SqlDataSourceLcsKind" DataTextField="Item" DataValueField="Item">
                    </asp:DropDownList>
                    &nbsp;&nbsp;&nbsp;&nbsp;
                    <asp:CheckBox ID="ChkSwStatus" runat="server" />軟體狀態
                    <asp:DropDownList ID="SelSwStatus" runat="server" ForeColor="#009933"
                        DataSourceID="SqlDataSourceSwStatus" DataTextField="Item" DataValueField="Item">
                    </asp:DropDownList>
                    <asp:SqlDataSource ID="SqlDataSourceSwStatus" runat="server" ConnectionString="<%$ ConnectionStrings:IDMSConnectionString %>"
                        SelectCommand="SELECT [Item] FROM [Config] WHERE [Kind] = '軟體狀態' ORDER BY [mark]">
                    </asp:SqlDataSource>
                    &nbsp;&nbsp;&nbsp;&nbsp;
                    <asp:CheckBox ID="ChkAskStatus" runat="server" />授權狀態
                    <asp:DropDownList ID="SelAskStatus" runat="server" ForeColor="#009933"
                        DataSourceID="SqlDataSourceAskStatus" DataTextField="Item" DataValueField="Item">
                    </asp:DropDownList>
                    <asp:SqlDataSource ID="SqlDataSourceAskStatus" runat="server" ConnectionString="<%$ ConnectionStrings:IDMSConnectionString %>"
                        SelectCommand="SELECT [Item] FROM [Config] WHERE [Kind] = '授權狀態' ORDER BY [mark]">
                    </asp:SqlDataSource>
                    &nbsp;&nbsp;&nbsp;&nbsp;                    
                    <asp:CheckBox ID="ChkDevStatus" runat="server" />設備狀態
                    <asp:DropDownList ID="SelDevStatus" runat="server" ForeColor="#009933"
                        DataSourceID="SqlDataSourceDevStatus" DataTextField="Item" DataValueField="Item">
                    </asp:DropDownList>
                    <asp:SqlDataSource ID="SqlDataSourceDevStatus" runat="server" ConnectionString="<%$ ConnectionStrings:IDMSConnectionString %>"
                        SelectCommand="SELECT [Item] FROM [Config] WHERE [Kind] = '設備狀態' ORDER BY [mark]">
                    </asp:SqlDataSource>                   
                    &nbsp;&nbsp;&nbsp;&nbsp;                    
                    <asp:CheckBox ID="ChkApStatus" runat="server" />作業狀態
                    <asp:DropDownList ID="SelApStatus" runat="server" ForeColor="#009933"
                        DataSourceID="SqlDataSourceApStatus" DataTextField="Item" DataValueField="Item">
                    </asp:DropDownList>
                    <asp:SqlDataSource ID="SqlDataSourceApStatus" runat="server" ConnectionString="<%$ ConnectionStrings:IDMSConnectionString %>"
                        SelectCommand="SELECT [Item] FROM [Config] WHERE [Kind] = '作業狀態' ORDER BY [mark]">
                    </asp:SqlDataSource>  
                </td>
            </tr>
            <tr>
                <td class="style7">不一致</td>
                <td>
                    <asp:CheckBox ID="ChkApNo" runat="server" />作業編號&nbsp;&nbsp;
                    <asp:CheckBox ID="ChkUnit" runat="server" />單位&nbsp;&nbsp;
                    <asp:CheckBox ID="ChkUser" runat="server" />人員&nbsp;&nbsp;
                    <asp:CheckBox ID="ChkHost" runat="server" />主機名稱&nbsp;&nbsp;
                    <asp:CheckBox ID="ChkIP" runat="server" />IP&nbsp;&nbsp;
                    <asp:CheckBox ID="ChkProp" runat="server" />財產編號&nbsp;&nbsp;
                    <asp:CheckBox ID="ChkBrand" runat="server" />廠牌&nbsp;&nbsp;
                    <asp:CheckBox ID="ChkStyle" runat="server" />型式&nbsp;&nbsp;&nbsp;&nbsp;
                    <span class="style5">(聯集查詢，但欄位空白不檢核)</span>
                </td>
            </tr>
            <tr>
                <td class="style7">SQL</td>
                <td>
                    <asp:CheckBox ID="ChkSQL" runat="server" />
                    <asp:TextBox ID="TextSQL" Columns="100" runat="server">有特殊查詢需求，請洽機房 (本段文字僅供說明，請勿執行查詢)</asp:TextBox>&nbsp;
                    <asp:DropDownList ID="SelSQL" runat="server" ForeColor="#009933" Visible="true" OnSelectedIndexChanged="SelSQL_SelectedIndexChanged"
                        AutoPostBack="True" DataSourceID="SqlDataSourceSQL" DataTextField="Item" DataValueField="Memo">
                    </asp:DropDownList>                    
                    <asp:SqlDataSource ID="SqlDataSourceSQL" runat="server" ConnectionString="<%$ ConnectionStrings:IDMSConnectionString %>"
                        SelectCommand="SELECT [Item],[Memo] from Config where Kind='SQL範本' and (Config='軟體查詢') order by Mark">
                    </asp:SqlDataSource>
                </td>
            </tr>
            <tr>
                <td class="style7">關鍵字</td>
                <td>                                        
                    <asp:TextBox ID="TextKey" runat="server"></asp:TextBox>&nbsp;<span class="style5">(逗點隔開)</span>&nbsp;&nbsp;
                    <asp:CheckBox ID="ChkSwKey" Checked="true" runat="server" />保管&nbsp;&nbsp;
                    <asp:CheckBox ID="ChkAskKey" Checked="true" runat="server" />授權&nbsp;&nbsp;
                    <asp:CheckBox ID="ChkApKey" runat="server" />作業&nbsp;&nbsp;
                    <asp:CheckBox ID="ChkDevKey" runat="server" />設備&nbsp;
                    <span class="style5">(聯集查詢)</span>&nbsp;&nbsp;&nbsp;&nbsp;
                    <asp:CheckBox ID="ChkPage" Checked="true" runat="server" />
                    <asp:Label ID="Label6" runat="server" Font-Size="Small" Text="分頁"></asp:Label>&nbsp;&nbsp;&nbsp;&nbsp;                    
                    顯示欄位：<asp:DropDownList ID="SelShow" runat="server" ForeColor="#009933" Visible="true" DataSourceID="SqlDataSourceShow" DataTextField="Item" DataValueField="Memo">
                    </asp:DropDownList>
                    &nbsp;
                    <asp:SqlDataSource ID="SqlDataSourceShow" runat="server" ConnectionString="<%$ ConnectionStrings:IDMSConnectionString %>"
                        SelectCommand="SELECT [Item],[Memo] from Config where [Kind]='軟體查詢' and [Config]='軟體查詢' order by [Mark]">
                    </asp:SqlDataSource>
                    <span class="style4" onclick="alert('顯示欄位如需客製化設定，請洽機房');">
                        <u style="cursor:pointer">說明</u>
                    </span
                </td>
            </tr>
        </table> 
        <div align="right"><asp:Button ID="Button1" runat="server" Text="　查　詢　" OnClick="Button1_Click" /></div>
    </p>
    <p>
        <asp:GridView ID="GridView1" runat="server" AllowPaging="True" AllowSorting="True"
            BackColor="#CCCCCC" BorderColor="#999999" BorderStyle="Solid"
            BorderWidth="3px" CellPadding="4" CellSpacing="2" DataSourceID="SqlDataSource1"
            DataKeyNames="授權編號" ForeColor="Black" Font-Size="Small" OnSelectedIndexChanging="GridView1_SelectedIndexChanging" PageSize="100">
            <Columns>
                <asp:CommandField ButtonType="Button" HeaderText="申請單" SelectText="編輯" ShowSelectButton="True" />
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
        <asp:SqlDataSource ID="SqlDataSource1" runat="server" ConnectionString="<%$ ConnectionStrings:IDMSConnectionString %>"
            SelectCommand=""></asp:SqlDataSource>
        <br />
    </p>
    <div class="style7">
        總筆數：<asp:TextBox ID="TextCount" runat="server" BackColor="#CCCCCC" ReadOnly="True" Width="60px" CssClass="style7" Font-Bold="True"></asp:TextBox>&nbsp;&nbsp;
        總價：<asp:TextBox ID="TextTotal" runat="server" BackColor="#CCCCCC" ReadOnly="True"
            Width="120px" CssClass="style7" Font-Bold="True"></asp:TextBox>&nbsp;&nbsp;&nbsp;&nbsp;
        <asp:Button ID="BtnExcel" runat="server" Text="匯出Excel檔 (html格式)" OnClick="BtnExcel_Click" />
    </div>
    <span class="style5">設定顯示欄位：[系統設定] -> [顯示欄位] ，所有可設定的欄位名稱在[系統設定] -> [一般設定] -> [顯示欄位] 可查到</span>
</asp:Content>
