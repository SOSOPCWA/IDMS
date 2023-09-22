<%@ Page Language="C#" Title="作業編輯" AutoEventWireup="true" CodeFile="ApEdit.aspx.cs" Inherits="Device_ApEdit" MaintainScrollPositionOnPostback="true" Debug="true" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <style type="text/css">
        .style1
        {
            text-align: center;
        }
    </style>
</head>
<body bgcolor="LightGray">
    <form id="form1" runat="server">
    <div style="text-align: center; font-family: 標楷體;">
        <asp:Label ID="LblApEdit" runat="server" Text="作業主機編輯介面" Style="font-weight: 700;
            color: #0000CC; font-size: xx-large">
        </asp:Label>
    </div>
    
    <div class="style1">
        <asp:Table ID="Table1" runat="server" Width="100%" GridLines="Both">
            <asp:TableRow ID="rowTitle" runat="server" Font-Bold="True" ForeColor="#800000">
                <asp:TableCell ID="TableCell1" runat="server" Width="150px">欄位名稱</asp:TableCell>
                <asp:TableCell ID="TableCell2" runat="server">
                    設定 (<asp:LinkButton ID="LinkHideHelp" runat="server" Font-Size="Small" ForeColor="Blue" OnClick="LinkHideHelp_Click" ToolTip="隱藏右方說明欄位，以簡化操作介面">隱藏說明</asp:LinkButton>)
                </asp:TableCell>
                <asp:TableCell ID="TableCell3" runat="server" Width="40%">說明</asp:TableCell>
            </asp:TableRow>
            <asp:TableRow ID="rowApNo" runat="server" HorizontalAlign="Left" VerticalAlign="Middle">
                <asp:TableCell runat="server" Font-Bold="True" ForeColor="Red">作業編號</asp:TableCell>
                <asp:TableCell runat="server" Font-Size="Small">
                    <asp:TextBox ID="TextApNo" runat="server" BackColor="Silver" Columns="2" Font-Size="Large"
                        ForeColor="Red" ReadOnly="True" CssClass="style1"></asp:TextBox> &nbsp;&nbsp;
                    <asp:LinkButton ID="LinkLcs" runat="server" Font-Size="Small" ForeColor="Blue" OnClick="LinkLcs_Click">軟體授權查詢</asp:LinkButton>
                </asp:TableCell>
                <asp:TableCell runat="server" Font-Size="Small">
                    <asp:Label ID="helpApNo" runat="server" Font-Size="Small" ForeColor="#003300"></asp:Label>
                </asp:TableCell>
            </asp:TableRow>

            <asp:TableRow ID="rowDevNo" runat="server" HorizontalAlign="Left" VerticalAlign="Middle">
                <asp:TableCell runat="server" Font-Bold="True" ForeColor="Red">設備編號(★)</asp:TableCell>
                <asp:TableCell runat="server" Font-Size="Small">
                    <asp:DropDownList ID="SelHostType" runat="server" ForeColor="#009933" DataSourceID="SqlDataSourceHostType" AppendDataBoundItems="true"
                        DataTextField="維護人員" DataValueField="設備編號" AutoPostBack="True" OnSelectedIndexChanged="SelHostType_SelectedIndexChanged">
                        <asp:ListItem Value="0">尚未設定</asp:ListItem>
                        <asp:ListItem Value="-1">一般主機</asp:ListItem>
                    </asp:DropDownList>&nbsp;&nbsp;
                    <asp:SqlDataSource ID="SqlDataSourceHostType" runat="server" ConnectionString="<%$ ConnectionStrings:IDMSConnectionString %>"
                        SelectCommand="SELECT [設備編號],'虛擬主機('+[維護人員]+')' as [維護人員] FROM [實體設備] where [設備種類]='虛擬主機'">
                    </asp:SqlDataSource>&nbsp;&nbsp;
                    <asp:TextBox ID="TextDevNo" runat="server" Columns="2" Font-Size="Large" ForeColor="Red"
                        CssClass="style1" OnKeyDown="if(event.keyCode <=57 | event.keyCode >=96 & event.keyCode <=105) {event.returnValue=true;} else{event.returnValue=false;}"></asp:TextBox>
                    &nbsp;&nbsp;
                    <asp:Label ID="LblDevName" runat="server"></asp:Label>&nbsp; 
                    <span onclick="window.open('../Device/Device.aspx?SelMode=Y','_blank');">
                        <font style="cursor: pointer; font-size: Small; color: Blue"><u>選取設備</u></font> 
                    </span>&nbsp;&nbsp; 
                    <span onclick="if(TextDevNo.value != '' & TextDevNo.value != '0') window.open('../Device/DevEdit.aspx?DevNo='+form1.TextDevNo.value,'_self');else alert('您尚未選取安裝系統作業的設備！');">
                        <font style="cursor: pointer; font-size: Small; color: Blue"><u>編輯設備</u></font>
                    </span>
                </asp:TableCell>
                <asp:TableCell runat="server" Font-Size="Small">
                    <asp:Label ID="helpDevNo" runat="server" Font-Size="Small" ForeColor="#003300"></asp:Label>
                </asp:TableCell>
            </asp:TableRow>

            <asp:TableRow ID="rowPowerOff" runat="server" HorizontalAlign="Left">
                <asp:TableCell runat="server" Font-Bold="True">關機順序(★)</asp:TableCell>
                <asp:TableCell runat="server" Font-Size="Small">
                    <asp:DropDownList ID="SelPowerOff" runat="server" ForeColor="#009933" DataSourceID="SqlDataSourcePowerOff" DataTextField="Item" DataValueField="Item" AutoPostBack="true" OnSelectedIndexChanged="SelPowerOff_SelectedIndexChanged">
                    </asp:DropDownList>&nbsp;&nbsp;
                    <asp:SqlDataSource ID="SqlDataSourcePowerOff" runat="server" ConnectionString="<%$ ConnectionStrings:IDMSConnectionString %>"
                        SelectCommand="SELECT [Item] FROM [Config] WHERE [Kind] = '關機順序' ORDER BY [mark]">
                    </asp:SqlDataSource>&nbsp;&nbsp;
                    <asp:Button ID="BtnPowerOff" runat="server" Text=" 更 新 " OnClick="BtnPowerOff_Click" />&nbsp;&nbsp;
                    <asp:Label ID="Label1" runat="server" ForeColor="Red" Text="(本欄位單獨更新至實體設備，虛擬主機不必設定)"></asp:Label> 
                    <br /><br />
                    <asp:Label ID="lblPowerOff" runat="server" Text=""></asp:Label>
                </asp:TableCell>
                <asp:TableCell runat="server" Font-Size="Small">
                    <asp:Label ID="helpPowerOff" runat="server" Font-Size="Small" ForeColor="#003300"></asp:Label>
                </asp:TableCell>
            </asp:TableRow>

            <asp:TableRow ID="rowHostName" runat="server" HorizontalAlign="Left" VerticalAlign="Middle">
                <asp:TableCell runat="server" Font-Bold="True">主機名稱(★)</asp:TableCell>
                <asp:TableCell runat="server" Font-Size="Small">
                    <asp:TextBox ID="TextHostName" runat="server" MaxLength="20" ForeColor="Green"></asp:TextBox>&nbsp;&nbsp;
                    <asp:LinkButton ID="LinkCopyDevName" runat="server" Font-Size="Small" OnClick="LinkCopyDevName_Click">複製設備名稱</asp:LinkButton>&nbsp;&nbsp;
                </asp:TableCell>
                <asp:TableCell runat="server" Font-Size="Small">
                    <asp:Label ID="helpHostName" runat="server" Font-Size="Small" ForeColor="#003300"></asp:Label>
                </asp:TableCell>
            </asp:TableRow>

            <asp:TableRow ID="rowSysNo" runat="server" HorizontalAlign="Left" VerticalAlign="Middle">
                <asp:TableCell runat="server" Font-Bold="True">系統名稱(★)</asp:TableCell>
                <asp:TableCell runat="server" Font-Size="Small">
                    <asp:DropDownList ID="SelSysNo" runat="server" ForeColor="#009933" DataSourceID="SqlDataSourceSys" AppendDataBoundItems="true"
                        DataTextField="系統選項" DataValueField="系統編號" AutoPostBack="True" OnSelectedIndexChanged="SelSysNo_SelectedIndexChanged">
                        <asp:ListItem Text="(空白)" Value="0"></asp:ListItem>
                    </asp:DropDownList>&nbsp;&nbsp; 
                    <span onclick="window.open('SysEdit.aspx?SysNo='+form1.SelSysNo.value,'_self');">
                        <font style="cursor: pointer; font-size: Small; color: Blue"><u>編輯系統</u></font>
                    </span>&nbsp;&nbsp;
                    <asp:Label ID="lblSysAlias" runat="server"></asp:Label> &nbsp;&nbsp;
                    <asp:SqlDataSource ID="SqlDataSourceSys" runat="server" ConnectionString="<%$ ConnectionStrings:IDMSConnectionString %>"
                        SelectCommand="SELECT [系統全名] as [系統選項],[資源編號] as [系統編號] FROM [View_系統資源] where [資源種類]='系統' order by [系統全名]">
                    </asp:SqlDataSource>
                    <br /><br />
                    系統資產：<asp:Label ID="lblAssets" runat="server"></asp:Label> 
                </asp:TableCell>
                <asp:TableCell runat="server" Font-Size="Small">
                    <asp:Label ID="helpSysNo" runat="server" Font-Size="Small" ForeColor="#003300"></asp:Label>
                </asp:TableCell>
            </asp:TableRow>

            <asp:TableRow ID="rowFunctions" runat="server" HorizontalAlign="Left" VerticalAlign="Middle">
                <asp:TableCell runat="server" Font-Bold="True">主要功能(★)</asp:TableCell>
                <asp:TableCell runat="server" Font-Size="Small">
                    <asp:TextBox ID="TextFunctions" runat="server" Columns="64" ForeColor="Green"></asp:TextBox>&nbsp;&nbsp;
                    <asp:LinkButton ID="LinkCopyDevFunc" runat="server" Font-Size="Small" OnClick="LinkCopyDevFunc_Click">複製設備用途</asp:LinkButton>
                </asp:TableCell>
                <asp:TableCell runat="server" Font-Size="Small">
                    <asp:Label ID="helpFunctions" runat="server" Font-Size="Small" ForeColor="#003300"></asp:Label>
                </asp:TableCell>
            </asp:TableRow>

            <asp:TableRow ID="rowSysFun" runat="server" HorizontalAlign="Left" VerticalAlign="Middle">
                <asp:TableCell runat="server" Font-Bold="True" ForeColor="Red">系統功能(★)</asp:TableCell>
                <asp:TableCell runat="server" Font-Size="Small">                    
                    <asp:Label ID="lblSysFun" runat="server" Font-Size="Small" ForeColor="green" Text="(無)"></asp:Label>  
                </asp:TableCell>
                <asp:TableCell runat="server" Font-Size="Small">
                    <asp:Label ID="helpSysFun" runat="server" Font-Size="Small" ForeColor="#003300"></asp:Label>
                </asp:TableCell>
            </asp:TableRow>

            <asp:TableRow ID="rowOS" runat="server" HorizontalAlign="Left" VerticalAlign="Middle">
                <asp:TableCell runat="server" Font-Bold="True">作業平台(★)</asp:TableCell>
                <asp:TableCell runat="server" Font-Size="Small">
                    作業平台：<asp:DropDownList ID="SelOS" runat="server" ForeColor="#009933" AutoPostBack="true" OnSelectedIndexChanged="SelOS_SelectedIndexChanged"
                        DataSourceID="SqlDataSource1" DataTextField="OsItem" DataValueField="Item">
                    </asp:DropDownList>
                    <asp:SqlDataSource ID="SqlDataSource1" runat="server" ConnectionString="<%$ ConnectionStrings:IDMSConnectionString %>"
                        SelectCommand="SELECT [Item],[Config]+' ： '+[Item] as [OsItem] FROM [Config] where [Kind]='作業平台' order by [Config],[Item]"></asp:SqlDataSource>
                    <br /><br />
                    核心版本：<asp:TextBox ID="TextKernel" runat="server" ForeColor="Green"></asp:TextBox>&nbsp;
                    <asp:DropDownList ID="SelKernel" runat="server" ForeColor="#009933" AutoPostBack="true"  OnSelectedIndexChanged="SelKernel_SelectedIndexChanged">                        
                    </asp:DropDownList>
                    
                </asp:TableCell>
                <asp:TableCell runat="server" Font-Size="Small">
                    <asp:Label ID="helpOS" runat="server" Font-Size="Small" ForeColor="#003300"></asp:Label>
                </asp:TableCell>
            </asp:TableRow>

            <asp:TableRow ID="rowIP" runat="server" HorizontalAlign="Left" VerticalAlign="Middle">
                <asp:TableCell runat="server" Font-Bold="True">IP</asp:TableCell>
                <asp:TableCell runat="server" Font-Size="Small">
                    <asp:TextBox ID="TextIP" runat="server" Columns="64" ForeColor="Green"></asp:TextBox> (可多筆)
                    <br /><br />
                    監控IP：<asp:TextBox ID="TextMosIP" runat="server" ForeColor="Green"></asp:TextBox> (單筆)&nbsp;
                    <asp:RegularExpressionValidator ID="RegularExpressionValidator1" runat="server" Display="Dynamic"
                        ErrorMessage="輸入格式有誤" ValidationExpression="\b(?:(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.){3}(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\b"
                        ControlToValidate="TextMosIP" SetFocusOnError="True">
                    </asp:RegularExpressionValidator>
                </asp:TableCell>
                <asp:TableCell runat="server" Font-Size="Small">
                    <asp:Label ID="helpIP" runat="server" Font-Size="Small" ForeColor="#003300"></asp:Label>
                </asp:TableCell>
            </asp:TableRow>

            <asp:TableRow ID="rowSaveIO" runat="server" HorizontalAlign="Left">
                <asp:TableCell runat="server" Font-Bold="True">安內外(★)</asp:TableCell>
                <asp:TableCell runat="server" Font-Size="Small">
                    <asp:DropDownList ID="SelSaveIO" runat="server" ForeColor="#009933">
                        <asp:ListItem Value="X">非網(X)</asp:ListItem>
                        <asp:ListItem Value="I">安內(I)</asp:ListItem>
                        <asp:ListItem Value="O">安外(O)</asp:ListItem>
                        <asp:ListItem Value="E">外網(E)</asp:ListItem>
						<asp:ListItem Value="D">DMZ區(D)</asp:ListItem>
                    </asp:DropDownList>
                </asp:TableCell>
                <asp:TableCell runat="server" Font-Size="Small">
                    <asp:Label ID="helpSaveIO" runat="server" Font-Size="Small" ForeColor="#003300"></asp:Label>
                </asp:TableCell>
            </asp:TableRow>

            <asp:TableRow ID="rowOnCall" runat="server" HorizontalAlign="Left">
                <asp:TableCell runat="server" Font-Bold="True">緊急程度(★)</asp:TableCell>
                <asp:TableCell runat="server" Font-Size="Small">
                    <asp:DropDownList ID="SelOnCall" runat="server" ForeColor="#009933">
                        <asp:ListItem Value="不用">不用</asp:ListItem>
                        <asp:ListItem Value="上班">上班</asp:ListItem>
                        <asp:ListItem Value="白天">白天</asp:ListItem>
                        <asp:ListItem Value="立即">立即</asp:ListItem>
                    </asp:DropDownList>
                </asp:TableCell>
                <asp:TableCell runat="server" Font-Size="Small">
                    <asp:Label ID="helpOnCall" runat="server" Font-Size="Small" ForeColor="#003300"></asp:Label>
                </asp:TableCell>
            </asp:TableRow>

            <asp:TableRow ID="rowApStatus" runat="server" HorizontalAlign="Left">
                <asp:TableCell runat="server" Font-Bold="True">作業狀態(★)</asp:TableCell>
                <asp:TableCell runat="server" Font-Size="Small">
                    <asp:DropDownList ID="SelApStatus" runat="server" ForeColor="#009933" DataSourceID="SqlDataSourceApStatus"
                        DataTextField="Item" DataValueField="Item">
                    </asp:DropDownList>
                    <asp:SqlDataSource ID="SqlDataSourceApStatus" runat="server" ConnectionString="<%$ ConnectionStrings:IDMSConnectionString %>"
                        SelectCommand="SELECT [Item] FROM [Config] WHERE [Kind] = '作業狀態' ORDER BY [mark]">
                    </asp:SqlDataSource>
                </asp:TableCell>
                <asp:TableCell runat="server" Font-Size="Small">
                    <asp:Label ID="helpApStatus" runat="server" Font-Size="Small" ForeColor="#003300"></asp:Label>
                </asp:TableCell>
            </asp:TableRow>

            <asp:TableRow ID="rowMaintainor" runat="server" HorizontalAlign="Left" VerticalAlign="Middle">
                <asp:TableCell runat="server" Font-Bold="True">維護人員(★)</asp:TableCell>
                <asp:TableCell runat="server" Font-Size="Small">
                    <asp:DropDownList ID="SelSwUnit" runat="server" AutoPostBack="True" ForeColor="#009933"
                        DataSourceID="SqlDataSource5" DataTextField="Item" DataValueField="Item">
                    </asp:DropDownList>
                    <asp:SqlDataSource ID="SqlDataSource5" runat="server" ConnectionString="<%$ ConnectionStrings:IDMSConnectionString %>"
                        SelectCommand="SELECT [Item],mark FROM [Config] 
                        WHERE (Kind='組織架構' and Item='數值資訊組') or 
                        Kind='維護群組' or Kind='數值資訊組' or Kind='署內單位'
                        ORDER BY CASE WHEN mark LIKE '%b%' THEN 9 END, mark asc">
                        <SelectParameters>
                            <asp:Parameter DefaultValue="維護群組" Name="Kind" Type="String" />
                            <asp:Parameter DefaultValue="數值資訊組" Name="MIC" Type="String" />
                            <asp:Parameter DefaultValue="署內單位" Name="CWA" Type="String" />
                        </SelectParameters>
                    </asp:SqlDataSource>
                    &nbsp;&nbsp;
                    <asp:DropDownList ID="SelMaintainor" runat="server" AutoPostBack="True" ForeColor="#009933"
                        DataSourceID="SqlDataSource6" DataTextField="Item" DataValueField="Item">
                    </asp:DropDownList>
                    <asp:SqlDataSource ID="SqlDataSource6" runat="server" ConnectionString="<%$ ConnectionStrings:IDMSConnectionString %>"
                        SelectCommand="SELECT [Item] FROM [Config] WHERE ([Kind] = @Kind) and [Config]<>'測試' ORDER BY [mark]">
                        <SelectParameters>
                            <asp:ControlParameter ControlID="SelSwUnit" Name="Kind" PropertyName="SelectedValue"
                                Type="String" />
                        </SelectParameters>
                    </asp:SqlDataSource>
                    <asp:Label ID="LblMaintainor" runat="server"></asp:Label>
                    <br />
                    <br />
                    <asp:Label ID="Label5" runat="server" Font-Size="Small" ForeColor="Red">建議盡量設成群組，若無該選項，請向機房反映</asp:Label>
                </asp:TableCell>
                <asp:TableCell runat="server" Font-Size="Small">
                    <asp:Label ID="helpMaintainor" runat="server" Font-Size="Small" ForeColor="#003300"></asp:Label>
                </asp:TableCell>
            </asp:TableRow>

            <asp:TableRow ID="rowOnLine" runat="server" HorizontalAlign="Left" VerticalAlign="Middle">
                <asp:TableCell runat="server" Font-Bold="True">上線日期</asp:TableCell>
                <asp:TableCell runat="server" Font-Size="Small">
                    <asp:DropDownList ID="SelYYYY" runat="server" ForeColor="#009933" AppendDataBoundItems="true">
                        <asp:ListItem Value="1900">1900</asp:ListItem>
                    </asp:DropDownList>年 &nbsp;&nbsp;
                    <asp:DropDownList ID="SelMM" runat="server" ForeColor="#009933"></asp:DropDownList>月 &nbsp;&nbsp;
                    <asp:DropDownList ID="SelDD" runat="server" ForeColor="#009933"></asp:DropDownList>日 
                </asp:TableCell>
                <asp:TableCell runat="server" Font-Size="Small">
                    <asp:Label ID="helpOnLine" runat="server" Font-Size="Small" ForeColor="#003300"></asp:Label>
                </asp:TableCell>
            </asp:TableRow>

            <asp:TableRow ID="rowMemo" runat="server" HorizontalAlign="Left" VerticalAlign="Middle">
                <asp:TableCell runat="server" Font-Bold="True">備註說明</asp:TableCell>
                <asp:TableCell runat="server" Font-Size="Small">
                    <asp:TextBox ID="TextMemo" runat="server" Columns="48" Width="500px" Height="100px" ForeColor="Green"
                        Rows="5" TextMode="MultiLine"></asp:TextBox>
                </asp:TableCell>
                <asp:TableCell runat="server" Font-Size="Small">
                    <asp:Label ID="helpMemo" runat="server" Font-Size="Small" ForeColor="#003300"></asp:Label>
                </asp:TableCell>
            </asp:TableRow>

            <asp:TableRow ID="rowCreate" runat="server" HorizontalAlign="Left" VerticalAlign="Middle">
                <asp:TableCell runat="server" Font-Bold="True">資料建立</asp:TableCell>
                <asp:TableCell runat="server" Font-Size="Small">
                    <asp:Label ID="LblCreater" runat="server" Font-Size="Small"></asp:Label>
                    &nbsp;
                    <asp:Label ID="LblCreateDT" runat="server" Font-Size="Small"></asp:Label>
                </asp:TableCell>
                <asp:TableCell runat="server" Font-Size="Small">
                    <asp:Label ID="helpCreate" runat="server" Font-Size="Small" ForeColor="#003300"></asp:Label>
                </asp:TableCell>
            </asp:TableRow>

            <asp:TableRow ID="rowUpdate" runat="server" HorizontalAlign="Left" VerticalAlign="Middle">
                <asp:TableCell ID="TableCell27" runat="server" Font-Bold="True">資料修改</asp:TableCell>
                <asp:TableCell runat="server" Font-Size="Small">
                    <asp:Label ID="LblUpdater" runat="server" Font-Size="Small"></asp:Label>
                    &nbsp;
                    <asp:Label ID="LblUpdateDT" runat="server" Font-Size="Small"></asp:Label>
                </asp:TableCell>
                <asp:TableCell runat="server" Font-Size="Small">
                    <asp:Label ID="helpUpdate" runat="server" Font-Size="Small" ForeColor="#003300"></asp:Label>
                </asp:TableCell>
            </asp:TableRow>

            <asp:TableRow ID="rowChecks" runat="server" HorizontalAlign="Left" VerticalAlign="Middle">
                <asp:TableCell ID="TableCell4" runat="server" Font-Bold="True">資料查核</asp:TableCell>
                <asp:TableCell ID="TableCell5" runat="server" Font-Size="Small">
                    <asp:Label ID="lblChecks" runat="server" ForeColor="red" Text="" Font-Bold="True"></asp:Label>
                    <br />
                    <asp:Label ID="lblRepeat" runat="server" ForeColor="red" Text="" Font-Bold="True"></asp:Label>&nbsp;
                </asp:TableCell>
                <asp:TableCell ID="TableCell6" runat="server" Font-Size="Small">
                    <asp:Label ID="helpChecks" runat="server" Font-Size="Small" ForeColor="#003300"></asp:Label>
                </asp:TableCell>
            </asp:TableRow>
        </asp:Table>
    </div>
    <br />
    <div class="style1">        
        <asp:Button ID="BtnAdd" runat="server" Text="　新　增　" OnClick="BtnAdd_Click" OnClientClick="return confirm('您確定要新增這筆資料嗎？')" ForeColor="Red" />&nbsp;&nbsp;&nbsp;&nbsp;
        <asp:Button ID="BtnEdit" runat="server" Text="　修　改　" OnClick="BtnEdit_Click" OnClientClick="return confirm('您確定要更新這筆資料嗎？')" />&nbsp;&nbsp;&nbsp;&nbsp;        
        <asp:Button ID="BtnDel" runat="server" Text="　刪　除　" OnClientClick="return confirm('請注意若是換機安裝，請修改設備編號即可！ 您仍確定要永久停止供應該項作業服務而要刪除這筆資料嗎？')"
            OnClick="BtnDel_Click" CausesValidation="False" />&nbsp;&nbsp;&nbsp;&nbsp;
        <asp:Button ID="BtnLife" runat="server" Text=" 生命履歷 "   OnClick="BtnLife_Click" CausesValidation="False" />&nbsp;&nbsp;&nbsp;&nbsp;
        <asp:Button ID="BtnLcs" runat="server" Text=" 軟體授權 "   OnClick="LinkLcs_Click" CausesValidation="False" />&nbsp;&nbsp;&nbsp;&nbsp;
        <asp:Button ID="BtnExit" runat="server" Text="　關　閉　" OnClientClick="window.close();" CausesValidation="False" /> &nbsp;&nbsp;&nbsp;&nbsp;
        <asp:HyperLink ID="HyperLink1" runat="server" NavigateUrl="~/IDMS/Help/權限管理與生命履歷.txt" Target="_blank" Font-Size="X-Small">權限說明</asp:HyperLink>
        <br />
        <asp:Label ID="Label2" runat="server" Font-Size="Small" ForeColor="#006600" Text="(★：必填欄位　　紅色：連結欄位)"></asp:Label>        
    </div>
    </form>
</body>
</html>
