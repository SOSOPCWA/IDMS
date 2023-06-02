<%@ Page Language="C#" Title="設備編輯" AutoEventWireup="true" CodeFile="DevEdit.aspx.cs" Inherits="Device_DevEdit" MaintainScrollPositionOnPostback="true" Debug="true" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>設備編輯</title>
    <style type="text/css">
        .style1
        {
            text-align: center;
        }
        .style2
        {
            font-size: Small;
        }
    </style>
</head>
<body bgcolor="LightGray">
    <form id="form1" runat="server">
    <div style="text-align: center; font-family: 標楷體;">
        <asp:Label ID="LblDevEdit" runat="server" Text="實體設備編輯介面" Style="font-weight: 700;
            color: #0000CC; font-size: xx-large"></asp:Label>
    </div>
    <div class="style1">
        <asp:Table ID="Table1" runat="server" Width="100%" GridLines="Both">
            <asp:TableRow ID="rowTitle" runat="server" Font-Bold="True" ForeColor="#800000">
                <asp:TableCell runat="server" Width="150px">欄位名稱</asp:TableCell>
                <asp:TableCell runat="server">
                    設定 (<asp:LinkButton ID="LinkHideHelp" runat="server" Font-Size="Small" ForeColor="Blue" OnClick="LinkHideHelp_Click" ToolTip="隱藏右方說明欄位，以簡化操作介面">隱藏說明</asp:LinkButton>)
                </asp:TableCell>
                <asp:TableCell runat="server" Width="40%">說明</asp:TableCell>
            </asp:TableRow>

            <asp:TableRow ID="rowDevNo" runat="server" HorizontalAlign="Left" VerticalAlign="Middle">
                <asp:TableCell runat="server" Font-Bold="True" ForeColor="Red">設備編號</asp:TableCell>
                <asp:TableCell runat="server" Font-Size="Small">
                    <asp:TextBox ID="TextDevNo" runat="server" BackColor="Silver" Columns="2" Font-Size="Large"
                        ForeColor="Red" ReadOnly="True" CssClass="style1"></asp:TextBox>
                    &nbsp;&nbsp;
                </asp:TableCell>
                <asp:TableCell runat="server" Font-Size="Small">
                    <asp:Label ID="helpDevNo" runat="server" Font-Size="Small" ForeColor="#003300"></asp:Label>
                </asp:TableCell>
            </asp:TableRow>

            <asp:TableRow ID="rowSysAp" runat="server" HorizontalAlign="Left" VerticalAlign="Middle">
                <asp:TableCell runat="server" Font-Bold="True" ForeColor="Red">作業主機</asp:TableCell>
                <asp:TableCell runat="server" Font-Size="Small">
                    <asp:Label ID="lblAP" runat="server"  ForeColor="Red" Font-Size="Large" Text="(無)"></asp:Label> &nbsp;&nbsp;
                    <asp:LinkButton ID="LinkAP" runat="server" onclick="LinkAP_Click">編輯主機</asp:LinkButton>                     
                </asp:TableCell>
                <asp:TableCell ID="TableCell6" runat="server" Font-Size="Small">
                    <asp:Label ID="helpSysAp" runat="server" Font-Size="Small" ForeColor="#003300"></asp:Label>
                </asp:TableCell>
            </asp:TableRow>

            <asp:TableRow ID="rowAssetsNo" runat="server" HorizontalAlign="Left" VerticalAlign="Middle">
                <asp:TableCell runat="server" Font-Bold="True">資產編號(★)</asp:TableCell>
                <asp:TableCell runat="server" Font-Size="Small">
                    <asp:DropDownList ID="SelAssets" runat="server" ForeColor="#009933" AutoPostBack="True" OnSelectedIndexChanged="SelAssets_SelectedIndexChanged">
                    </asp:DropDownList><br /><br />
                    <asp:Label ID="lblAssets" runat="server" Font-Size="Small"></asp:Label>
                </asp:TableCell>
                <asp:TableCell runat="server" Font-Size="Small">
                    <asp:Label ID="helpAssetsNo" runat="server" Font-Size="Small" ForeColor="#003300"></asp:Label>
                </asp:TableCell>
            </asp:TableRow>

            <asp:TableRow ID="rowDevType" runat="server" HorizontalAlign="Left" VerticalAlign="Middle">
                <asp:TableCell runat="server" Font-Bold="True">設備型態(★)</asp:TableCell>
                <asp:TableCell runat="server" Font-Size="Small">
                    <asp:DropDownList ID="SelDevType" runat="server" AutoPostBack="True" ForeColor="#009933"
                        DataSourceID="SqlDataSource13" DataTextField="Item" DataValueField="Item">
                    </asp:DropDownList>
                    <asp:SqlDataSource ID="SqlDataSource13" runat="server" ConnectionString="<%$ ConnectionStrings:IDMSConnectionString %>"
                        SelectCommand="SELECT [Item] FROM [Config] WHERE [Kind] = '設備型態' ORDER BY [mark]">
                    </asp:SqlDataSource>
                </asp:TableCell>
                <asp:TableCell runat="server" Font-Size="Small">
                    <asp:Label ID="helpDevType" runat="server" Font-Size="Small" ForeColor="#003300"></asp:Label>
                </asp:TableCell>
            </asp:TableRow>

            <asp:TableRow ID="rowDevKind" runat="server" HorizontalAlign="Left" VerticalAlign="Middle">
                <asp:TableCell runat="server" Font-Bold="True">設備種類(★)</asp:TableCell>
                <asp:TableCell runat="server" Font-Size="Small">
                    <asp:DropDownList ID="SelDevKind" runat="server" AutoPostBack="True" ForeColor="#009933"
                        DataSourceID="SqlDataSource3" DataTextField="Item" DataValueField="Item">
                    </asp:DropDownList>
                    <asp:SqlDataSource ID="SqlDataSource3" runat="server" ConnectionString="<%$ ConnectionStrings:IDMSConnectionString %>"
                        SelectCommand="SELECT [Item] FROM [Config] WHERE ([Kind] = @設備型態) ORDER BY [mark]">
                        <SelectParameters>
                            <asp:ControlParameter ControlID="SelDevType" DefaultValue="系統設備" Name="設備型態" PropertyName="SelectedValue"
                                Type="String" />
                        </SelectParameters>
                    </asp:SqlDataSource>
                </asp:TableCell>
                <asp:TableCell runat="server" Font-Size="Small">
                    <asp:Label ID="helpDevKind" runat="server" Font-Size="Small" ForeColor="#003300"></asp:Label>
                </asp:TableCell>
            </asp:TableRow>

            <asp:TableRow ID="rowDevName" runat="server" HorizontalAlign="Left" VerticalAlign="Middle">
                <asp:TableCell runat="server" Font-Bold="True">設備名稱(★)</asp:TableCell>
                <asp:TableCell runat="server" Font-Size="Small">
                    <asp:TextBox ID="TextDevName" runat="server" ForeColor="Green"></asp:TextBox>
                    <asp:RequiredFieldValidator ID="RequiredFieldValidator1" runat="server" ControlToValidate="TextDevName"
                        ErrorMessage="未輸入設備名稱" Font-Size="Small" SetFocusOnError="True" Display="Dynamic"></asp:RequiredFieldValidator>
                </asp:TableCell>
                <asp:TableCell runat="server" Font-Size="Small">
                    <asp:Label ID="helpDevName" runat="server" Font-Size="Small" ForeColor="#003300"></asp:Label>
                </asp:TableCell>
            </asp:TableRow>

            <asp:TableRow ID="rowPurpose" runat="server" HorizontalAlign="Left" VerticalAlign="Middle">
                <asp:TableCell runat="server" Font-Bold="True">設備用途(★)</asp:TableCell>
                <asp:TableCell runat="server" Font-Size="Small">
                    <asp:TextBox ID="TextPurpose" runat="server" Columns="48" ForeColor="Green"></asp:TextBox>
                </asp:TableCell>
                <asp:TableCell runat="server" Font-Size="Small">
                    <asp:Label ID="helpPurpose" runat="server" Font-Size="Small" ForeColor="#003300"></asp:Label>
                </asp:TableCell>
            </asp:TableRow>

            <asp:TableRow ID="rowPropNo" runat="server" HorizontalAlign="Left" VerticalAlign="Middle">
                <asp:TableCell runat="server" Font-Bold="True" ForeColor="Red">財產編號(★)</asp:TableCell>
                <asp:TableCell runat="server" Font-Size="Small">
                    A(11)：<asp:TextBox ID="TextPropNoA" runat="server" Columns="12" MaxLength="11" ForeColor="Green"></asp:TextBox> &nbsp;&nbsp; 
                    B(9 or 10)：<asp:TextBox ID="TextPropNoB" runat="server" Columns="12" MaxLength="10" ForeColor="Green"></asp:TextBox> &nbsp;&nbsp;                    
                    <span><font color="blue" style="cursor:pointer" onclick="TextPropNoA.value='00000000000'"><u>不列管</u></font></span>&nbsp;&nbsp;
                    <asp:LinkButton ID="LinkPropIn" runat="server" onclick="LinkPropIn_Click">查詢</asp:LinkButton> <br /><br />
                    財產名稱：<asp:Label ID="lblProp" runat="server"></asp:Label> <br />
                    財產附屬：<asp:Label ID="lblSpec" runat="server" Font-Size="Small"></asp:Label> <br />
                    數量單位：<asp:Label ID="lblAmounts" runat="server"></asp:Label> <br />
                    價值　　：<asp:Label ID="lblPrice" runat="server"></asp:Label> <br />
                    購買日期：<asp:Label ID="lblBuyDate" runat="server"></asp:Label> <br />
                    使用年限：<asp:Label ID="lblUseLife" runat="server"></asp:Label> <br />
                </asp:TableCell>
                <asp:TableCell runat="server" Font-Size="Small">
                    <asp:Label ID="helpPropNo" runat="server" Font-Size="Small" ForeColor="#003300"></asp:Label>
                </asp:TableCell>
            </asp:TableRow>

            <asp:TableRow ID="rowBrand" runat="server" HorizontalAlign="Left" VerticalAlign="Middle">
                <asp:TableCell runat="server" Font-Bold="True">廠牌</asp:TableCell>
                <asp:TableCell runat="server" Font-Size="Small">
                    <asp:TextBox ID="TextBrand" runat="server" ForeColor="Green"></asp:TextBox> &nbsp;&nbsp;
                    <asp:Label ID="Label4" runat="server" ForeColor="Red" Font-Bold="True" Text="請保持空白，除非需要！"></asp:Label> &nbsp;&nbsp;
                    秘總：<asp:Label ID="lblBrand" runat="server"></asp:Label>
                </asp:TableCell>
                <asp:TableCell runat="server" Font-Size="Small">
                    <asp:Label ID="helpBrand" runat="server" Font-Size="Small" ForeColor="#003300"></asp:Label>
                </asp:TableCell>
            </asp:TableRow>

            <asp:TableRow ID="rowStyle" runat="server" HorizontalAlign="Left" VerticalAlign="Middle">
                <asp:TableCell runat="server" Font-Bold="True">型式</asp:TableCell>
                <asp:TableCell runat="server" Font-Size="Small">
                    <asp:TextBox ID="TextStyle" runat="server" ForeColor="Green"></asp:TextBox> &nbsp;&nbsp;
                    <asp:Label ID="Label5" runat="server" ForeColor="Red" Font-Bold="True" Text="請保持空白，除非需要！"></asp:Label> &nbsp;&nbsp;
                    秘總：<asp:Label ID="lblStyle" runat="server"></asp:Label>
                </asp:TableCell>
                <asp:TableCell runat="server" Font-Size="Small">
                    <asp:Label ID="helpStyle" runat="server" Font-Size="Small" ForeColor="#003300"></asp:Label>
                </asp:TableCell>
            </asp:TableRow>

            <asp:TableRow ID="rowKeeper" runat="server" HorizontalAlign="Left" VerticalAlign="Middle">
                <asp:TableCell runat="server" Font-Bold="True">保管人員</asp:TableCell>
                <asp:TableCell runat="server" Font-Size="Small">
                    設備保管：<asp:TextBox ID="TextKeeper" runat="server"  Columns="6" ForeColor="Green"></asp:TextBox> &nbsp;&nbsp;
                    <asp:DropDownList ID="SelKeeperUnit" runat="server" AutoPostBack="True" ForeColor="#009933" DataSourceID="SqlDataSource10" DataTextField="課別" DataValueField="課別">
                    </asp:DropDownList>
                    &nbsp;&nbsp;
                    <asp:SqlDataSource ID="SqlDataSource10" runat="server" ConnectionString="<%$ ConnectionStrings:IDMSConnectionString %>"
                        SelectCommand="SELECT [課別] from View_課別 order by [代號]">
                        <SelectParameters>
                        </SelectParameters>
                    </asp:SqlDataSource>
                    <asp:DropDownList ID="SelKeeper" runat="server" ForeColor="#009933" DataSourceID="SqlDataSource11" OnSelectedIndexChanged="SelKeeper_SelectedIndexChanged" AutoPostBack="true"
                        DataTextField="Item" DataValueField="Item">
                    </asp:DropDownList>
                    <asp:SqlDataSource ID="SqlDataSource11" runat="server" ConnectionString="<%$ ConnectionStrings:IDMSConnectionString %>"
                        SelectCommand="SELECT [Item] FROM [Config] WHERE ([Kind] = @Kind) and [Config]<>'測試' ORDER BY [mark]">
                        <SelectParameters>
                            <asp:ControlParameter ControlID="SelKeeperUnit" Name="Kind" PropertyName="SelectedValue"
                                Type="String" />
                        </SelectParameters>
                    </asp:SqlDataSource> &nbsp;&nbsp;
                    <asp:LinkButton ID="LinkKeeperClear" runat="server" onclick="LinkKeeperClear_Click">空白</asp:LinkButton>&nbsp;&nbsp;
                    <asp:Label ID="Label7" runat="server" ForeColor="Red" Font-Bold="True" Text="請保持空白，除非需要！"></asp:Label> <br /><br />
                    財產保管：<asp:Label ID="lblKeeper" runat="server"></asp:Label> <br /><br />
                    使用人員：<asp:Label ID="lblUser" runat="server"></asp:Label> 
                </asp:TableCell>
                <asp:TableCell runat="server" Font-Size="Small">
                    <asp:Label ID="helpKeeper" runat="server" Font-Size="Small" ForeColor="#003300"></asp:Label>
                </asp:TableCell>
            </asp:TableRow>

            <asp:TableRow ID="rowDevStatus" runat="server" HorizontalAlign="Left">
                <asp:TableCell ID="TableCell1" runat="server" Font-Bold="True">設備狀態(★)</asp:TableCell>
                <asp:TableCell ID="TableCell2" runat="server" Font-Size="Small">
                    <asp:DropDownList ID="SelDevStatus" runat="server" ForeColor="#009933"
                        DataSourceID="SqlDataSourceDevStatus" DataTextField="Item" DataValueField="Item">
                    </asp:DropDownList>
                    <asp:SqlDataSource ID="SqlDataSourceDevStatus" runat="server" ConnectionString="<%$ ConnectionStrings:IDMSConnectionString %>"
                        SelectCommand="SELECT [Item] FROM [Config] WHERE [Kind] = '設備狀態' ORDER BY [mark]">
                    </asp:SqlDataSource>&nbsp;&nbsp;
                    秘總：<asp:Label ID="lblStatus" runat="server"></asp:Label>
                </asp:TableCell>
                <asp:TableCell ID="TableCell3" runat="server" Font-Size="Small">
                    <asp:Label ID="helpDevStatus" runat="server" Font-Size="Small" ForeColor="#003300"></asp:Label>
                </asp:TableCell>
            </asp:TableRow>

            <asp:TableRow ID="rowPlace" runat="server" HorizontalAlign="Left" VerticalAlign="Middle">
                <asp:TableCell runat="server" Font-Bold="True">放置地點(★)</asp:TableCell>
                <asp:TableCell runat="server" Font-Size="Small">
                    <asp:DropDownList ID="SelPlace" runat="server" AutoPostBack="True" OnSelectedIndexChanged="SelPointer_SelectedIndexChanged"
                        ForeColor="#009933" DataSourceID="SqlDataSource1" DataTextField="區域名稱" DataValueField="區域名稱">
                    </asp:DropDownList>
                    <asp:SqlDataSource ID="SqlDataSource1" runat="server" ConnectionString="<%$ ConnectionStrings:IDMSConnectionString %>"
                        SelectCommand="SELECT DISTINCT [區域名稱] FROM [定位設定] WHERE ([定位方式] &lt;&gt; @定位方式) order by [區域名稱]">
                        <SelectParameters>
                            <asp:Parameter DefaultValue="坐標" Name="定位方式" Type="String" />
                        </SelectParameters>
                    </asp:SqlDataSource>
                    &nbsp;
                    <asp:DropDownList ID="SelPointer" runat="server" AutoPostBack="True" OnSelectedIndexChanged="SelPointer_SelectedIndexChanged"
                        ForeColor="#009933" DataSourceID="SqlDataSource2" DataTextField="定位名稱" DataValueField="定位編號">
                    </asp:DropDownList>
                    <asp:SqlDataSource ID="SqlDataSource2" runat="server" ConnectionString="<%$ ConnectionStrings:IDMSConnectionString %>"
                        SelectCommand="SELECT [定位編號],[定位名稱] FROM [定位設定] WHERE (([定位方式] &lt;&gt; @定位方式) AND ([區域名稱] = @區域名稱)) order by [定位名稱]">
                        <SelectParameters>
                            <asp:Parameter DefaultValue="坐標" Name="定位方式" Type="String" />
                            <asp:ControlParameter ControlID="SelPlace" Name="區域名稱" PropertyName="SelectedValue"
                                Type="String" />
                        </SelectParameters>
                    </asp:SqlDataSource>
                    <span onclick="PointerNo=form1.TextPointerNo.value;if(PointerNo=='') PointerNo=form1.SelPointer.value;window.open('../Lib/map.aspx?PointerNo='+PointerNo,'_blank','location=no;menubar=no;resizable=yes;scrollbars=no;status=no;toolbar=no');">
                        &nbsp; &nbsp; 或 &nbsp; &nbsp;<font style="cursor: pointer; font-size: Small; color: Blue"><u>坐標定位</u></font>
                    </span>&nbsp;&nbsp;
                    <asp:TextBox ID="TextPointerNo" runat="server" BackColor="Silver" Columns="4" OnKeyDown="event.returnValue=false;"></asp:TextBox>   
                    <br /><br />
                    機櫃高度： <asp:DropDownList ID="SelRU" runat="server" ForeColor="#009933"></asp:DropDownList> (低) &nbsp;&nbsp;&nbsp;&nbsp;               
                    空間大小： <asp:DropDownList ID="SelSpace" runat="server" ForeColor="#009933"></asp:DropDownList> (U)  
                </asp:TableCell>
                <asp:TableCell runat="server" Font-Size="Small">
                    <asp:Label ID="helpPlace" runat="server" Font-Size="Small" ForeColor="#003300"></asp:Label>
                </asp:TableCell>
            </asp:TableRow>

            <asp:TableRow ID="rowSource" runat="server" HorizontalAlign="Left" VerticalAlign="Middle">
                <asp:TableCell runat="server" Font-Bold="True">取得來源</asp:TableCell>
                <asp:TableCell runat="server" Font-Size="Small">
                    <asp:TextBox ID="TextSource" runat="server" ForeColor="Green"></asp:TextBox>
                    &nbsp;&nbsp;
                    <asp:DropDownList ID="SelSource" runat="server" AutoPostBack="True" OnSelectedIndexChanged="SelSource_SelectedIndexChanged"
                        DataSourceID="SqlDataSource12" DataTextField="取得來源" DataValueField="取得來源" ForeColor="#009933">
                    </asp:DropDownList>
                    <asp:SqlDataSource ID="SqlDataSource12" runat="server" ConnectionString="<%$ ConnectionStrings:IDMSConnectionString %>"
                        SelectCommand="SELECT DISTINCT [取得來源] FROM [實體設備] WHERE ([設備型態] = @設備型態) order by [取得來源]">
                        <SelectParameters>
                            <asp:ControlParameter ControlID="SelDevType" DefaultValue="系統設備" Name="設備型態" PropertyName="SelectedValue"
                                Type="String" />
                        </SelectParameters>
                    </asp:SqlDataSource>
                </asp:TableCell>
                <asp:TableCell runat="server" Font-Size="Small">
                    <asp:Label ID="helpSource" runat="server" Font-Size="Small" ForeColor="#003300"></asp:Label>
                </asp:TableCell>
            </asp:TableRow>            

            <asp:TableRow ID="rowHostName" runat="server" HorizontalAlign="Left" VerticalAlign="Middle">
                <asp:TableCell runat="server" Font-Bold="True">主機名稱(★)</asp:TableCell>
                <asp:TableCell runat="server" Font-Size="Small">
                    <asp:TextBox ID="TextHostName" runat="server" ForeColor="Green"></asp:TextBox>&nbsp;&nbsp;
                    <asp:Label ID="lblApNo" Text="" runat="server" Font-Size="Small"></asp:Label>
                    <asp:RequiredFieldValidator ID="RequiredFieldValidator2" runat="server" ControlToValidate="TextHostName"
                        ErrorMessage="未輸入主機名稱" Font-Size="Small" SetFocusOnError="True" Display="Dynamic"></asp:RequiredFieldValidator>
                </asp:TableCell>
                <asp:TableCell runat="server" Font-Size="Small">
                    <asp:Label ID="helpHostName" runat="server" Font-Size="Small" ForeColor="#003300"></asp:Label>
                </asp:TableCell>
            </asp:TableRow>

            <asp:TableRow ID="rowIP" runat="server" HorizontalAlign="Left" VerticalAlign="Middle">
                <asp:TableCell runat="server" Font-Bold="True">IP</asp:TableCell>
                <asp:TableCell runat="server" Font-Size="Small">
                    <asp:TextBox ID="TextIP" runat="server" Columns="64" ForeColor="Green"></asp:TextBox> (可多筆)
                    <br /><br />
                    監控IP：<asp:TextBox ID="TextMosIP" runat="server" ForeColor="Green"></asp:TextBox> (單筆)&nbsp;                    
                    <asp:RegularExpressionValidator ID="RegularExpressionValidator1" runat="server" Display="Dynamic"
                        ErrorMessage="IP輸入格式有誤" ValidationExpression="\b(?:(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.){3}(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\b"
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

            <asp:TableRow ID="rowSpec" runat="server" HorizontalAlign="Left" VerticalAlign="Middle">
                <asp:TableCell runat="server" Font-Bold="True">規格</asp:TableCell>
                <asp:TableCell runat="server" Font-Size="Small">
                    <asp:TextBox ID="TextSpec" runat="server" Columns="48" ForeColor="Green"></asp:TextBox> 
                </asp:TableCell>
                <asp:TableCell runat="server" Font-Size="Small">
                    <asp:Label ID="helpSpec" runat="server" Font-Size="Small" ForeColor="#003300"></asp:Label>
                </asp:TableCell>
            </asp:TableRow>

            <asp:TableRow ID="rowSerial" runat="server" HorizontalAlign="Left" VerticalAlign="Middle">
                <asp:TableCell runat="server" Font-Bold="True">硬體序號</asp:TableCell>
                <asp:TableCell runat="server" Font-Size="Small">
                    <asp:TextBox ID="TextSerial" runat="server" ForeColor="Green"></asp:TextBox>
                </asp:TableCell>
                <asp:TableCell runat="server" Font-Size="Small">
                    <asp:Label ID="helpSerial" runat="server" Font-Size="Small" ForeColor="#003300"></asp:Label>
                </asp:TableCell>
            </asp:TableRow>

            <asp:TableRow ID="rowiLoIP" runat="server" HorizontalAlign="Left" VerticalAlign="Middle">
                <asp:TableCell runat="server" Font-Bold="True">iLoIP</asp:TableCell>
                <asp:TableCell runat="server" Font-Size="Small">
                    <asp:TextBox ID="TextiLoIP" runat="server" ForeColor="Green"></asp:TextBox>
                </asp:TableCell>
                <asp:TableCell runat="server" Font-Size="Small">
                    <asp:Label ID="helpiLoIP" runat="server" Font-Size="Small" ForeColor="#003300"></asp:Label>
                </asp:TableCell>
            </asp:TableRow>

            <asp:TableRow ID="rowVoltage" runat="server" HorizontalAlign="Left">
                <asp:TableCell runat="server" Font-Bold="True">用電電壓</asp:TableCell>
                <asp:TableCell runat="server" Font-Size="Small">
                    <asp:TextBox ID="TextVoltage" runat="server" Columns="6" MaxLength="4" ForeColor="Green"
                        OnKeyDown="if(event.keyCode <=57 | event.keyCode >=96 & event.keyCode <=105) {event.returnValue=true;} else{event.returnValue=false;}"> </asp:TextBox>
                    <span class="style2">(伏特)</span>
                    <asp:RangeValidator ID="RangeValidator3" runat="server" ControlToValidate="TextVoltage"
                        Display="Dynamic" ErrorMessage="用電電壓輸入有誤" Font-Size="Small" MaximumValue="1000"
                        MinimumValue="0" SetFocusOnError="True" Type="Integer">
                    </asp:RangeValidator>
                    <asp:RequiredFieldValidator ID="RequiredFieldValidator3" runat="server" ControlToValidate="TextVoltage"
                        ErrorMessage="未輸入用電電壓" Font-Size="Small" SetFocusOnError="True" Display="Dynamic">
                    </asp:RequiredFieldValidator>
                </asp:TableCell>
                <asp:TableCell runat="server" Font-Size="Small">
                    <asp:Label ID="helpVoltage" runat="server" Font-Size="Small" ForeColor="#003300"></asp:Label>
                </asp:TableCell>
            </asp:TableRow>

            <asp:TableRow ID="rowSpecCurrent" runat="server" HorizontalAlign="Left" VerticalAlign="Middle">
                <asp:TableCell runat="server" Font-Bold="True">額定電流</asp:TableCell>
                <asp:TableCell runat="server" Font-Size="Small">
                    <asp:TextBox ID="TextSpecCurrent" runat="server" Columns="6" MaxLength="6" ForeColor="Green"
                    OnKeyDown="if(event.keyCode <=57 | event.keyCode >=96 & event.keyCode <=105 | event.keyCode ==110 | event.keyCode ==190) {event.returnValue=true;} else{event.returnValue=false;}"></asp:TextBox>
                    <span class="style2">(安培)</span>
                    <asp:RangeValidator ID="RangeValidator1" runat="server" ControlToValidate="TextSpecCurrent"
                        Display="Dynamic" ErrorMessage="額定電流輸入有誤" Font-Size="Small" MaximumValue="1500.0"
                        MinimumValue="0.0" SetFocusOnError="True" Type="Double">
                    </asp:RangeValidator>
                    <asp:RequiredFieldValidator ID="RequiredFieldValidator4" runat="server" ControlToValidate="TextSpecCurrent"
                        ErrorMessage="未輸入額定電流" Font-Size="Small" SetFocusOnError="True" Display="Dynamic">
                    </asp:RequiredFieldValidator>
                </asp:TableCell>
                <asp:TableCell runat="server" Font-Size="Small">
                    <asp:Label ID="helpSpecCurrent" runat="server" Font-Size="Small" ForeColor="#003300"></asp:Label>
                </asp:TableCell>
            </asp:TableRow>

            <asp:TableRow ID="rowConn" runat="server" HorizontalAlign="Left" VerticalAlign="Middle">
                <asp:TableCell runat="server" Font-Bold="True" ForeColor="Red">設備迴路(★)</asp:TableCell>
                <asp:TableCell runat="server" Font-Size="Small">  
                    <table>
                        <tr>
                            <td><asp:Button ID="BtnDevTree" runat="server" Text=" 設 定 " CausesValidation="False" OnClick="BtnDevTree_Click" />&nbsp;&nbsp;&nbsp;&nbsp;</td>
                            <td>
                                上游供電：<asp:Label ID="lblUpPower" runat="server" Font-Size="Small"></asp:Label>
                                <br />
                                下游用電：<asp:Label ID="lblDnPower" runat="server" Font-Size="Small"></asp:Label>
                                <br />
                                上游供網：<asp:Label ID="lblUpNet" runat="server" Font-Size="Small"></asp:Label>
                                <br />
                                下游用網：<asp:Label ID="lblDnNet" runat="server" Font-Size="Small"></asp:Label>
                            </td>
                        </tr>
                    </table>                                       
                </asp:TableCell>
                <asp:TableCell runat="server" Font-Size="Small">
                    <asp:Label ID="helpConn" runat="server" Font-Size="Small" ForeColor="#003300"></asp:Label>
                </asp:TableCell>
            </asp:TableRow>

            <asp:TableRow ID="rowRepair" runat="server" HorizontalAlign="Left" VerticalAlign="Middle">
                <asp:TableCell runat="server" Font-Bold="True">維護廠商</asp:TableCell>
                <asp:TableCell runat="server" Font-Size="Small">
                    <asp:TextBox ID="TextRepair" runat="server" MaxLength="20" ForeColor="Green"></asp:TextBox>
                    &nbsp;&nbsp;
                    <asp:DropDownList ID="SelRepair" runat="server" AutoPostBack="True" OnSelectedIndexChanged="SelRepair_SelectedIndexChanged"
                        ForeColor="#009933" DataSourceID="SqlDataSourceRepair" DataTextField="維護廠商" DataValueField="維護廠商">
                    </asp:DropDownList>
                    <asp:SqlDataSource ID="SqlDataSourceRepair" runat="server" ConnectionString="<%$ ConnectionStrings:IDMSConnectionString %>"
                        SelectCommand="SELECT DISTINCT [維護廠商] FROM [實體設備] WHERE ([設備型態] = @設備型態)">
                        <SelectParameters>
                            <asp:ControlParameter ControlID="SelDevType" DefaultValue="系統設備" Name="設備型態" PropertyName="SelectedValue"
                                Type="String" />
                        </SelectParameters>
                    </asp:SqlDataSource>
                    &nbsp;&nbsp;
                    <asp:Label ID="LblRepair" runat="server"></asp:Label>
                </asp:TableCell>
                <asp:TableCell runat="server" Font-Size="Small">
                    <asp:Label ID="helpRepair" runat="server" Font-Size="Small" ForeColor="#003300"></asp:Label>
                </asp:TableCell>
            </asp:TableRow>

            <asp:TableRow ID="rowMaintainor" runat="server" HorizontalAlign="Left" VerticalAlign="Middle">
                <asp:TableCell runat="server" Font-Bold="True">維護人員(★)</asp:TableCell>
                <asp:TableCell runat="server" Font-Size="Small">
                    <asp:DropDownList ID="SelHwUnit" runat="server" AutoPostBack="True" ForeColor="#009933"
                        DataSourceID="SqlDataSource5" DataTextField="Item" DataValueField="Item">
                    </asp:DropDownList>
                    &nbsp;&nbsp;
                    <asp:SqlDataSource ID="SqlDataSource5" runat="server" ConnectionString="<%$ ConnectionStrings:IDMSConnectionString %>"
                        SelectCommand="SELECT [Item] FROM [Config] WHERE (Kind='組織架構' and Item=@MIC) or ([Kind] = @Kind or Kind=@MIC) ORDER BY [mark]">
                        <SelectParameters>
                            <asp:Parameter DefaultValue="維護群組" Name="Kind" Type="String" />
                            <asp:Parameter DefaultValue="資訊中心" Name="MIC" Type="String" />
                        </SelectParameters>
                    </asp:SqlDataSource>
                    <asp:DropDownList ID="SelMaintainor" runat="server" AutoPostBack="True" ForeColor="#009933"
                        DataSourceID="SqlDataSource6" DataTextField="Item" DataValueField="Item">
                    </asp:DropDownList>
                    <asp:SqlDataSource ID="SqlDataSource6" runat="server" ConnectionString="<%$ ConnectionStrings:IDMSConnectionString %>"
                        SelectCommand="SELECT [Item] FROM [Config] WHERE ([Kind] = @Kind) and [Config]<>'測試' ORDER BY [mark]">
                        <SelectParameters>
                            <asp:ControlParameter ControlID="SelHwUnit" Name="Kind" PropertyName="SelectedValue"
                                Type="String" />
                        </SelectParameters>
                    </asp:SqlDataSource>
                    <br />
                    <br />
                    <asp:Label ID="LblMaintainor" runat="server"></asp:Label>
                    &nbsp;&nbsp;
                    <asp:RequiredFieldValidator ID="RequiredFieldValidator5" runat="server" ControlToValidate="SelMaintainor"
                        ErrorMessage="未選擇維護人員" Font-Size="Small" SetFocusOnError="True" Display="Dynamic"></asp:RequiredFieldValidator>
                    <br />
                    <br />
                    <asp:Label ID="Label6" runat="server" Font-Size="Small" ForeColor="Red">建議盡量設成群組，若無該選項，請向機房反映</asp:Label>
                </asp:TableCell>
                <asp:TableCell runat="server" Font-Size="Small">
                    <asp:Label ID="helpMaintainor" runat="server" ForeColor="#003300"></asp:Label>
                </asp:TableCell>
            </asp:TableRow>

            <asp:TableRow ID="rowPowerOff" runat="server" HorizontalAlign="Left">
                <asp:TableCell runat="server" Font-Bold="True">關機順序(★)</asp:TableCell>
                <asp:TableCell runat="server" Font-Size="Small">
                    <asp:DropDownList ID="SelPowerOff" runat="server" ForeColor="#009933" DataSourceID="SqlDataSourcePowerOff" DataTextField="Item" DataValueField="Item" AutoPostBack="true" OnSelectedIndexChanged="SelPowerOff_SelectedIndexChanged">
                    </asp:DropDownList>
                    <asp:SqlDataSource ID="SqlDataSourcePowerOff" runat="server" ConnectionString="<%$ ConnectionStrings:IDMSConnectionString %>"
                        SelectCommand="SELECT [Item] FROM [Config] WHERE [Kind] = '關機順序' ORDER BY [mark]">
                    </asp:SqlDataSource>&nbsp;&nbsp;
                    <asp:Label ID="lblPowerOff" runat="server" Text=""></asp:Label>
                </asp:TableCell>
                <asp:TableCell runat="server" Font-Size="Small">
                    <asp:Label ID="helpPowerOff" runat="server" Font-Size="Small" ForeColor="#003300"></asp:Label>
                </asp:TableCell>
            </asp:TableRow>
            
            <asp:TableRow ID="rowMemo" runat="server" HorizontalAlign="Left" VerticalAlign="Middle">
                <asp:TableCell runat="server" Font-Bold="True">備註說明</asp:TableCell>
                <asp:TableCell runat="server" Font-Size="Small">
                    <asp:TextBox ID="TextMemo" runat="server" Columns="48" Width="500px" Height="100px" Rows="5" TextMode="MultiLine" ForeColor="Green"></asp:TextBox>                    
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
                <asp:TableCell runat="server" Font-Bold="True">資料修改</asp:TableCell>
                <asp:TableCell runat="server" Font-Size="Small">
                    <asp:Label ID="LblUpdater" runat="server" Font-Size="Small"></asp:Label>
                    &nbsp;
                    <asp:Label ID="LblUpdateDT" runat="server" Font-Size="Small"></asp:Label>
                </asp:TableCell>
                <asp:TableCell runat="server" Font-Size="Small">
                    <asp:Label ID="helpUpdate" runat="server" Font-Size="Small" ForeColor="#003300"></asp:Label>
                </asp:TableCell>
            </asp:TableRow>

            <asp:TableRow ID="rowDevCheck" runat="server" HorizontalAlign="Left" VerticalAlign="Middle">
                <asp:TableCell runat="server" Font-Bold="True">機器清查</asp:TableCell>
                <asp:TableCell runat="server" Font-Size="Small">
                    <asp:Label ID="lblDevCheck" runat="server"></asp:Label>&nbsp;&nbsp;
                    <asp:Button ID="BtnDevCheck" runat="server" Text="清查設定" OnClick="BtnDevCheck_Click" CausesValidation="False" ToolTip="僅環境小組有權限設定" />　
                </asp:TableCell>
                <asp:TableCell runat="server" Font-Size="Small">
                    <asp:Label ID="helpDevCheck" runat="server" Font-Size="Small" ForeColor="#003300"></asp:Label>
                </asp:TableCell>
            </asp:TableRow>

            <asp:TableRow ID="rowChecks" runat="server" HorizontalAlign="Left" VerticalAlign="Middle">
                <asp:TableCell runat="server" Font-Bold="True">資料查核</asp:TableCell>
                <asp:TableCell runat="server" Font-Size="Small">
                    <asp:Label ID="lblChecks" runat="server" ForeColor="red" Text="" Font-Bold="True"></asp:Label>
                    <br />
                    <asp:Label ID="lblRepeat" runat="server" ForeColor="red" Text="" Font-Bold="True"></asp:Label>&nbsp;
                </asp:TableCell>
                <asp:TableCell runat="server" Font-Size="Small">
                    <asp:Label ID="helpChecks" runat="server" Font-Size="Small" ForeColor="#003300"></asp:Label>
                </asp:TableCell>
            </asp:TableRow>
        </asp:Table>
    </div>    
    <br />
    <div class="style1">
        <asp:Button ID="BtnAdd" runat="server" Text="　新　增　" OnClick="BtnAdd_Click" 
            OnClientClick="return confirm('您確定要新增這筆資料嗎？')" ForeColor="Red" />　　
        <asp:Button ID="BtnEdit" runat="server" Text="　修　改　" OnClick="BtnEdit_Click" OnClientClick="return confirm('您確定要修改這筆資料嗎？')" />　　
        <asp:Button ID="BtnDel" runat="server" Text="　刪　除　"  CausesValidation="False" OnClick="BtnDel_Click"
            OnClientClick="return confirm('請注意作業主機與接電迴路將一併刪除！若要保留，請先於作業主機介面清除本設備編號！您確定要刪除這筆資料嗎？')" />　　　
        <asp:Button ID="BtnLife" runat="server" Text=" 生命履歷 "   OnClick="BtnLife_Click" CausesValidation="False" />　　　
        <asp:Button ID="BtnIn" runat="server" Text=" 列印移入單 "   OnClick="BtnIn_Click" CausesValidation="False" />　　
        <asp:Button ID="BtnExit" runat="server" Text="　關　閉　" OnClientClick="window.close();" CausesValidation="False" />　　
        <asp:HyperLink ID="HyperLink1" runat="server" NavigateUrl="~/IDMS/Help/權限管理與生命履歷.txt" Target="_blank" Font-Size="X-Small">權限說明</asp:HyperLink>
        <br />
        <asp:Label ID="Label1" runat="server" Font-Size="Small" ForeColor="#006600" Text="(★：必填欄位　　紅色：連結欄位)"></asp:Label>
    </div>    
    </form>
</body>
</html>
