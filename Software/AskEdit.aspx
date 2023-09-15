<%@ Page Language="C#" Title="授權編輯" AutoEventWireup="true" CodeFile="AskEdit.aspx.cs" Inherits="Software_AskEdit" MaintainScrollPositionOnPostback="true" Debug="true" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>授權編輯</title>
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
        <asp:Label ID="LblAskEdit" runat="server" Text="軟體使用申請單編輯介面" Style="font-weight: 700;
            color: #0000CC; font-size: xx-large"></asp:Label>
    </div>
    <div class="style1">
        <asp:Table ID="Table1" runat="server" Width="100%" GridLines="Both">
            <asp:TableRow ID="rowTitle" runat="server" Font-Bold="True" ForeColor="#800000">
                <asp:TableCell ID="TableCell16" runat="server" Width="150px">欄位名稱</asp:TableCell>
                <asp:TableCell ID="TableCell17" runat="server">
                    設定 (<asp:LinkButton ID="LinkHideHelp" runat="server" Font-Size="Small" ForeColor="Blue" OnClick="LinkHideHelp_Click" ToolTip="隱藏右方說明欄位，以簡化操作介面">隱藏說明</asp:LinkButton>)
                </asp:TableCell>
                <asp:TableCell ID="TableCell20" runat="server" Width="35%">說明</asp:TableCell>
            </asp:TableRow>

            <asp:TableRow ID="rowAskNo" runat="server" HorizontalAlign="Left" VerticalAlign="Middle">
                <asp:TableCell runat="server" Font-Bold="True" ForeColor="Red">授權編號</asp:TableCell>
                <asp:TableCell runat="server" Font-Size="Small">
                    <asp:TextBox ID="TextAskNo" runat="server" BackColor="Silver" Columns="2" Font-Size="Large"
                        ForeColor="Red" ReadOnly="True" CssClass="style1"></asp:TextBox>
                </asp:TableCell>
                <asp:TableCell runat="server" Font-Size="Small">
                    <asp:Label ID="helpAskNo" runat="server" Font-Size="Small" ForeColor="#003300"></asp:Label>
                </asp:TableCell>
            </asp:TableRow>            

            <asp:TableRow ID="rowAskFormNo" runat="server" HorizontalAlign="Left" VerticalAlign="Middle">
                <asp:TableCell runat="server" Font-Bold="True" ForeColor="Red">申請單列表(★)</asp:TableCell>
                <asp:TableCell runat="server" Font-Size="Small">
                    <table border="1">
                        <tr align="center">
                            <td>申請單列表</td>
                            <td width="80px">執行</td>
                            <td>申請單資料</td>
                        </tr>
                        <tr align="center">
                            <td>                                
                                <asp:ListBox ID="SelAskFormNo" runat="server" ForeColor="#009933" DataSourceID="SqlDataSourceAskFormNo" DataTextField="表單項目" DataValueField="表單編號" 
                                    AutoPostBack="True" OnSelectedIndexChanged="SelAskFormNo_SelectedIndexChanged" AppendDataBoundItems="False" Rows="5">
                                </asp:ListBox>
                                <asp:SqlDataSource ID="SqlDataSourceAskFormNo" runat="server" ConnectionString="<%$ ConnectionStrings:IDMSConnectionString %>"
                                    SelectCommand="SELECT *,[表單編號]+' '+[申請事項] as [表單項目] FROM [申請表單] where [表單種類]='軟體申請' and [主鍵編號]=@授權編號 order by [表單編號] desc">
                                    <SelectParameters>
                                        <asp:ControlParameter ControlID="TextAskNo" Name="授權編號" PropertyName="Text" Type="Int16" />
                                    </SelectParameters>
                                </asp:SqlDataSource>
                            </td>
                            <td align="center">
                                <asp:LinkButton ID="LinkAskFormAdd" runat="server" Font-Size="Small" OnClick="LinkAskFormAdd_Click" Text="<< 新增 <<"></asp:LinkButton><br/><br/>
                                <asp:LinkButton ID="LinkAskFormDel" runat="server" Font-Size="Small" OnClientClick="return confirm('您確定要刪除這筆資料嗎？')" OnClick="LinkAskFormDel_Click" Text=">> 刪除 >>"></asp:LinkButton>                                
                            </td>
                            <td>
                                <table>
                                    <tr><td>表單：</td><td><asp:TextBox ID="TextAskFormNo" runat="server" ForeColor="Green" OnKeyDown="if(event.keyCode <=57 | event.keyCode >=96 & event.keyCode <=105 | event.keyCode=189) {event.returnValue=true;} else{event.returnValue=false;}"> </asp:TextBox></td></tr>
                                    <tr><td>事項：</td><td><asp:TextBox ID="TextAskFormMemo" runat="server" ForeColor="Green"></asp:TextBox></td></tr>
                                    <tr><td>日期：</td><td>
                                        <asp:DropDownList ID="SelFormYYYY" runat="server" ForeColor="#009933"></asp:DropDownList>年
                                        <asp:DropDownList ID="SelFormMM" runat="server" ForeColor="#009933"></asp:DropDownList>月
                                        <asp:DropDownList ID="SelFormDD" runat="server" ForeColor="#009933"></asp:DropDownList>日
                                    </td></tr>
                                </table>                                                                    
                            </td>                            
                        </tr>
                    </table> 
                    <asp:Label ID="Label2" runat="server" Font-Size="Small" ForeColor="#FF0000" Text="* 請完成新增資料後後再編輯此欄位，否則會有問題!"></asp:Label>
                </asp:TableCell>
                <asp:TableCell runat="server" Font-Size="Small">
                    <asp:Label ID="helpAskFormNo" runat="server" Font-Size="Small" ForeColor="#003300"></asp:Label>
                </asp:TableCell>
            </asp:TableRow>

             <asp:TableRow ID="rowStatus" runat="server" HorizontalAlign="Left">
                <asp:TableCell runat="server" Font-Bold="True">授權狀態</asp:TableCell>
                <asp:TableCell runat="server" Font-Size="Small">
                    <asp:DropDownList ID="SelStatus" runat="server" ForeColor="#009933" OnSelectedIndexChanged="SelStatus_SelectedIndexChanged"
                        DataSourceID="SqlDataSourceAskStatus" DataTextField="Item" DataValueField="Item" AutoPostBack="true">
                    </asp:DropDownList>&nbsp;&nbsp;
                    軟體狀態：<asp:Label ID="lblSwStatus" runat="server"></asp:Label>
                    <asp:SqlDataSource ID="SqlDataSourceAskStatus" runat="server" ConnectionString="<%$ ConnectionStrings:IDMSConnectionString %>"
                        SelectCommand="SELECT [Item] FROM [Config] WHERE [Kind] = '授權狀態' ORDER BY [mark]">
                    </asp:SqlDataSource>
                </asp:TableCell>
                <asp:TableCell ID="TableCell19" runat="server" Font-Size="Small">
                    <asp:Label ID="helpStatus" runat="server" Font-Size="Small" ForeColor="#003300"></asp:Label>
                </asp:TableCell>
            </asp:TableRow>  

            <asp:TableRow ID="rowKeyDay" runat="server" HorizontalAlign="Left" VerticalAlign="Middle">
                <asp:TableCell ID="TableCell1" runat="server" Font-Bold="True">填表日期</asp:TableCell>
                <asp:TableCell ID="TableCell2" runat="server" Font-Size="Small">
                    <asp:DropDownList ID="SelKeyYYYY" runat="server" ForeColor="#009933"></asp:DropDownList>年 &nbsp;&nbsp;
                    <asp:DropDownList ID="SelKeyMM" runat="server" ForeColor="#009933"></asp:DropDownList>月 &nbsp;&nbsp;
                    <asp:DropDownList ID="SelKeyDD" runat="server" ForeColor="#009933"></asp:DropDownList>日                    
                </asp:TableCell>
                <asp:TableCell ID="TableCell3" runat="server" Font-Size="Small">
                    <asp:Label ID="helpKeyDay" runat="server" Font-Size="Small" ForeColor="#003300"></asp:Label>
                </asp:TableCell>
            </asp:TableRow>

            <asp:TableRow ID="rowKeyer" runat="server" HorizontalAlign="Left" VerticalAlign="Middle">
                <asp:TableCell ID="TableCell4" runat="server" Font-Bold="True">填表人員</asp:TableCell>
                <asp:TableCell ID="TableCell5" runat="server" Font-Size="Small">
                    <asp:TextBox ID="TextKeyer" runat="server" ForeColor="Green"></asp:TextBox>&nbsp;&nbsp;
                    <asp:DropDownList ID="SelKeyerUnit" runat="server" AutoPostBack="true" ForeColor="#009933" OnSelectedIndexChanged="SelKeyerUnit_SelectedIndexChanged"
                        DataSourceID="SqlDataSourceKeyerUnit" DataTextField="Item" DataValueField="Item">
                    </asp:DropDownList>
                    <asp:SqlDataSource ID="SqlDataSourceKeyerUnit" runat="server" ConnectionString="<%$ ConnectionStrings:IDMSConnectionString %>"
                        SelectCommand="SELECT [Item] FROM [Config] WHERE [Kind]='數值資訊組' order by [Mark]">
                    </asp:SqlDataSource>&nbsp;&nbsp;
                    <asp:DropDownList ID="SelKeyer" runat="server" ForeColor="#009933" DataSourceID="SqlDataSourceKeyer"
                        DataTextField="Item" DataValueField="Item" AutoPostBack="True" OnSelectedIndexChanged="SelKeyer_SelectedIndexChanged">
                    </asp:DropDownList>
                    <asp:SqlDataSource ID="SqlDataSourceKeyer" runat="server" ConnectionString="<%$ ConnectionStrings:IDMSConnectionString %>"
                        SelectCommand="SELECT [Item] FROM [Config] WHERE ([Kind] = @Kind) ORDER BY [mark]">
                        <SelectParameters>
                            <asp:ControlParameter ControlID="SelKeyerUnit" Name="Kind" PropertyName="SelectedValue" Type="String" />
                        </SelectParameters>
                    </asp:SqlDataSource>
                </asp:TableCell>
                <asp:TableCell ID="TableCell6" runat="server" Font-Size="Small">
                    <asp:Label ID="helpKeyer" runat="server" Font-Size="Small" ForeColor="#003300"></asp:Label>
                </asp:TableCell>
            </asp:TableRow> 
            
            <asp:TableRow ID="rowSwNo" runat="server" HorizontalAlign="Left" VerticalAlign="Middle">
                <asp:TableCell ID="TableCell10" runat="server" Font-Bold="True" ForeColor="Red">保管單(★)</asp:TableCell>
                <asp:TableCell ID="TableCell11" runat="server" Font-Size="Small">
                    編號：<asp:TextBox ID="TextSwNo" runat="server" Columns="2" Font-Size="Large" ForeColor="Red" CssClass="style1" 
                        OnKeyDown="if(event.keyCode <=57 | event.keyCode >=96 & event.keyCode <=105) {event.returnValue=true;} else{event.returnValue=false;}"></asp:TextBox>&nbsp;&nbsp;        
                    <span onclick="window.open('Sw.aspx?SelMode=Y','_blank');" >
                        <font style="cursor:pointer;font-size:Small;color:Blue"><u>選取</u></font>                    
                    </span>&nbsp;&nbsp;
                    <asp:LinkButton ID="LinkSwEdit" runat="server" onclick="LinkSwEdit_Click">編輯</asp:LinkButton> &nbsp;&nbsp;                    
                    <asp:LinkButton ID="LinkSwIn" runat="server" onclick="LinkSwIn_Click">帶入</asp:LinkButton> 
                    &nbsp;&nbsp;&nbsp;&nbsp;<font color="green"><b>或</b></font>&nbsp;&nbsp;&nbsp;&nbsp;
                    表單：<asp:TextBox ID="TextFormNo" runat="server" Columns="6" MaxLength="9"  ForeColor="Green"
                        OnKeyDown="if(event.keyCode <=57 | event.keyCode >=96 & event.keyCode <=105 | event.keyCode=189) {event.returnValue=true;} else{event.returnValue=false;}"></asp:TextBox>&nbsp;&nbsp;
                    <asp:LinkButton ID="LinkFormEdit" runat="server" onclick="LinkFormEdit_Click">編輯</asp:LinkButton> &nbsp;&nbsp;
                    <asp:LinkButton ID="LinkFormIn" runat="server" onclick="LinkFormIn_Click">帶入</asp:LinkButton> 
                    <br /> <br />                    
                    <asp:Label ID="LblSwName" runat="server" ForeColor="#003300" Font-Size="Large" Font-Bold="true"></asp:Label> 
                </asp:TableCell>
                <asp:TableCell ID="TableCell12" runat="server" Font-Size="Small">
                    <asp:Label ID="helpSwNo" runat="server" Font-Size="Small" ForeColor="#003300"></asp:Label>
                </asp:TableCell>
            </asp:TableRow>

            <asp:TableRow ID="rowIdentify" runat="server" HorizontalAlign="Left" VerticalAlign="Middle">
                <asp:TableCell ID="TableCell13" runat="server" Font-Bold="True">授權識別(★)</asp:TableCell>
                <asp:TableCell ID="TableCell14" runat="server" Font-Size="Small">
                    <asp:TextBox ID="TextIdentify" runat="server" Columns="2" Font-Size="Large" ForeColor="Red" CssClass="style1" OnKeyDown="if(event.keyCode <=57 | event.keyCode >=96 & event.keyCode <=105) {event.returnValue=true;} else{event.returnValue=false;}"></asp:TextBox>  &nbsp;&nbsp;
                    <asp:LinkButton ID="LinkIdentify" runat="server" Font-Size="Small" OnClick="LinkIdentify_Click">識別查詢</asp:LinkButton>            
                </asp:TableCell>
                <asp:TableCell ID="TableCell15" runat="server" Font-Size="Small">
                    <asp:Label ID="helpIdentify" runat="server" Font-Size="Small" ForeColor="#003300"></asp:Label>
                </asp:TableCell>
            </asp:TableRow>                

            <asp:TableRow ID="rowVersion" runat="server" HorizontalAlign="Left" VerticalAlign="Middle">
                <asp:TableCell runat="server" Font-Bold="True">軟體版本</asp:TableCell>
                <asp:TableCell runat="server" Font-Size="Small">
                    購買：<asp:Label ID="lblVersBuy" runat="server"></asp:Label> <br/>
                    可用：<asp:Label ID="lblVersCan" runat="server"></asp:Label> <br/>
                    申請：<asp:TextBox ID="TextVersAsk" runat="server" ForeColor="Green"></asp:TextBox> &nbsp;&nbsp;
                    <asp:DropDownList ID="SelVersAsk" runat="server" ForeColor="#009933" DataSourceID="SqlDataSourceVersAsk"
                        DataTextField="申請版本" DataValueField="申請版本" AutoPostBack="True" OnSelectedIndexChanged="SelVersAsk_SelectedIndexChanged">
                    </asp:DropDownList>
                    <asp:SqlDataSource ID="SqlDataSourceVersAsk" runat="server" ConnectionString="<%$ ConnectionStrings:IDMSConnectionString %>"
                        SelectCommand="SELECT DISTINCT [申請版本] FROM [軟體授權] where [軟體編號]=@軟體編號 order by [申請版本]">
                        <SelectParameters>
                            <asp:ControlParameter ControlID="TextSwNo" Name="軟體編號" PropertyName="Text" Type="Int16" />
                        </SelectParameters>
                    </asp:SqlDataSource>                    
                </asp:TableCell>
                <asp:TableCell runat="server" Font-Size="Small">
                    <asp:Label ID="helpVersion" runat="server" Font-Size="Small" ForeColor="#003300"></asp:Label>
                </asp:TableCell>
            </asp:TableRow>

            <asp:TableRow ID="rowLicense" runat="server" HorizontalAlign="Left" VerticalAlign="Middle">
                <asp:TableCell runat="server" Font-Bold="True">軟體授權</asp:TableCell>
                <asp:TableCell runat="server" Font-Size="Small">
                    購買：<asp:Label ID="lblLcsBuy" runat="server"></asp:Label> <br/>
                    申請：
                    <asp:DropDownList ID="SelLcsAsk" runat="server" ForeColor="#009933" DataSourceID="SqlDataSourceLcsAsk" DataTextField="Item" DataValueField="Item">
                    </asp:DropDownList>
                    <asp:SqlDataSource ID="SqlDataSourceLcsAsk" runat="server" ConnectionString="<%$ ConnectionStrings:IDMSConnectionString %>"
                        SelectCommand="SELECT [Item] FROM [Config] WHERE [Kind] = '授權方式' ORDER BY [mark]">
                    </asp:SqlDataSource> <br /><br />
                    說明：<asp:TextBox ID="TextLcsMemo" Width="350px" runat="server" ForeColor="Green"></asp:TextBox>
                </asp:TableCell>
                <asp:TableCell runat="server" Font-Size="Small">
                    <asp:Label ID="helpLicense" runat="server" Font-Size="Small" ForeColor="#003300"></asp:Label>
                </asp:TableCell>
            </asp:TableRow>

            <asp:TableRow ID="rowSN" runat="server" HorizontalAlign="Left" VerticalAlign="Middle">
                <asp:TableCell runat="server" Font-Bold="True">軟體序號</asp:TableCell>
                <asp:TableCell runat="server" Font-Size="Small">
                   購買：<asp:Label ID="lblSnBuy" runat="server"></asp:Label> <br/>
                   申請：<asp:TextBox ID="TextSnAsk" runat="server" Columns="48" ForeColor="Green"></asp:TextBox>
                </asp:TableCell>
                <asp:TableCell runat="server" Font-Size="Small">
                    <asp:Label ID="helpSN" runat="server" Font-Size="Small" ForeColor="#003300"></asp:Label>
                </asp:TableCell>
            </asp:TableRow>

            <asp:TableRow ID="rowAttach" runat="server" HorizontalAlign="Left" VerticalAlign="Middle">
                <asp:TableCell runat="server" Font-Bold="True">授權附件</asp:TableCell>
                <asp:TableCell runat="server" Font-Size="Small">
                   購買：<asp:Label ID="lblAttach" runat="server"></asp:Label> <br/>
                   申請：<asp:TextBox ID="TextAttach" runat="server" Columns="48" Width="500px" Height="50px" Rows="3" ForeColor="Green" TextMode="MultiLine"></asp:TextBox>
                </asp:TableCell>
                <asp:TableCell ID="TableCell21" runat="server" Font-Size="Small">
                    <asp:Label ID="helpAttach" runat="server" Font-Size="Small" ForeColor="#003300"></asp:Label>
                </asp:TableCell>
            </asp:TableRow>

            <asp:TableRow ID="rowApNo" runat="server" HorizontalAlign="Left" VerticalAlign="Middle">
                <asp:TableCell runat="server" Font-Bold="True" ForeColor="Red">作業編號</asp:TableCell>
                <asp:TableCell runat="server" Font-Size="Small">
                    <asp:TextBox ID="TextApNo" runat="server" Columns="2" Font-Size="Large" ForeColor="Red" CssClass="style1" 
                        OnKeyDown="if(event.keyCode <=57 | event.keyCode >=96 & event.keyCode <=105) {event.returnValue=true;} else{event.returnValue=false;}"></asp:TextBox> &nbsp;&nbsp;
                    <asp:Label ID="LblApNo" runat="server" ForeColor="#003300"></asp:Label> &nbsp;
                    <span onclick="window.open('../AP/AP.aspx?SelMode=Y','_blank');" >
                        <font style="cursor:pointer;font-size:Small;color:Blue"><u>選取</u></font>                    
                    </span>&nbsp;&nbsp;
                    <span onclick="if(TextApNo.value != '' & TextApNo.value != '0') window.open('../AP/ApEdit.aspx?ApNo='+form1.TextApNo.value,'_self');else alert('您尚未選取要授權的作業主機！');" >
                        <font style="cursor:pointer;font-size:Small;color:Blue"><u>編輯</u></font>                    
                    </span>&nbsp;&nbsp;
                    <asp:LinkButton ID="LinkApIn" runat="server" onclick="LinkApIn_Click">帶入</asp:LinkButton> &nbsp;&nbsp;
                    <asp:LinkButton ID="LinkApClear" runat="server" onclick="LinkApClear_Click">清除</asp:LinkButton>
                </asp:TableCell>
                <asp:TableCell ID="TableCell18" runat="server" Font-Size="Small">
                    <asp:Label ID="helpApNo" runat="server" Font-Size="Small" ForeColor="#003300"></asp:Label>
                </asp:TableCell>
            </asp:TableRow>

            <asp:TableRow ID="rowUnit" runat="server" HorizontalAlign="Left" VerticalAlign="Middle">
                <asp:TableCell runat="server" Font-Bold="True">申請單位</asp:TableCell>
                <asp:TableCell runat="server" Font-Size="Small">
                    授權：
                    <asp:DropDownList ID="SelUnitLcs" runat="server" AutoPostBack="true" ForeColor="#009933"  OnSelectedIndexChanged="SelUnitLcs_SelectedIndexChanged"
                        DataSourceID="SqlDataSourceUnitLcs" DataTextField="Item" DataValueField="Item">
                    </asp:DropDownList>
                    <asp:SqlDataSource ID="SqlDataSourceUnitLcs" runat="server" ConnectionString="<%$ ConnectionStrings:IDMSConnectionString %>"
                        SelectCommand="SELECT [Item] FROM [Config] WHERE [Kind]='數值資訊組' order by [Mark]">
                    </asp:SqlDataSource>&nbsp;&nbsp;
                    使用：<asp:Label ID="lblUnitUse" runat="server"></asp:Label>
                </asp:TableCell>
                <asp:TableCell runat="server" Font-Size="Small">
                    <asp:Label ID="helpUnit" runat="server" Font-Size="Small" ForeColor="#003300"></asp:Label>
                </asp:TableCell>
            </asp:TableRow>

            <asp:TableRow ID="rowUser" runat="server" HorizontalAlign="Left" VerticalAlign="Middle">
                <asp:TableCell runat="server" Font-Bold="True">申請人員</asp:TableCell>
                <asp:TableCell runat="server" Font-Size="Small">
                    授權：<asp:TextBox ID="TextUserLcs" runat="server" ForeColor="Green"></asp:TextBox>&nbsp;&nbsp;
                    <asp:DropDownList ID="SelUserLcs" runat="server" ForeColor="#009933" DataSourceID="SqlDataSourceUserLcs"
                        DataTextField="Item" DataValueField="Item" AutoPostBack="True" OnSelectedIndexChanged="SelUserLcs_SelectedIndexChanged">
                    </asp:DropDownList>
                    <asp:SqlDataSource ID="SqlDataSourceUserLcs" runat="server" ConnectionString="<%$ ConnectionStrings:IDMSConnectionString %>"
                        SelectCommand="SELECT [Item] FROM [Config] WHERE ([Kind] = @Kind) ORDER BY [mark]">
                        <SelectParameters>
                            <asp:ControlParameter ControlID="SelUnitLcs" Name="Kind" PropertyName="SelectedValue" Type="String" />
                        </SelectParameters>
                    </asp:SqlDataSource>
                    <br />
                    使用：<asp:Label ID="lblUserUse" runat="server"></asp:Label>
                </asp:TableCell>
                <asp:TableCell runat="server" Font-Size="Small">
                    <asp:Label ID="helpUser" runat="server" Font-Size="Small" ForeColor="#003300"></asp:Label>
                </asp:TableCell>
            </asp:TableRow>

            <asp:TableRow ID="rowHost" runat="server" HorizontalAlign="Left" VerticalAlign="Middle">
                <asp:TableCell runat="server" Font-Bold="True">主機名稱</asp:TableCell>
                <asp:TableCell runat="server" Font-Size="Small">
                    授權：<asp:TextBox ID="TextHostLcs" runat="server" ForeColor="Green"></asp:TextBox>&nbsp;&nbsp;
                    使用：<asp:Label ID="lblHostUse" runat="server"></asp:Label>
                </asp:TableCell>
                <asp:TableCell runat="server" Font-Size="Small">
                    <asp:Label ID="helpHost" runat="server" Font-Size="Small" ForeColor="#003300"></asp:Label>
                </asp:TableCell>
            </asp:TableRow>

            <asp:TableRow ID="rowIP" runat="server" HorizontalAlign="Left" VerticalAlign="Middle">
                <asp:TableCell runat="server" Font-Bold="True">IP</asp:TableCell>
                <asp:TableCell runat="server" Font-Size="Small">
                    授權：<asp:TextBox ID="TextIPLcs" runat="server" ForeColor="Green"></asp:TextBox>&nbsp;&nbsp;
                    使用：<asp:Label ID="lblIPUse" runat="server"></asp:Label>
                </asp:TableCell>
                <asp:TableCell runat="server" Font-Size="Small">
                    <asp:Label ID="helpIP" runat="server" Font-Size="Small" ForeColor="#003300"></asp:Label>
                </asp:TableCell>
            </asp:TableRow>

            <asp:TableRow ID="rowPropNo" runat="server" HorizontalAlign="Left" VerticalAlign="Middle">
                <asp:TableCell runat="server" Font-Bold="True">財產編號</asp:TableCell>
                <asp:TableCell runat="server" Font-Size="Small">
                    授權：<asp:TextBox ID="TextPropNoALcs" runat="server" ForeColor="Green"></asp:TextBox>&nbsp;&nbsp;
                         <asp:TextBox ID="TextPropNoBLcs" runat="server" ForeColor="Green"></asp:TextBox> <br/>
                    使用：<asp:Label ID="lblPropNoAUse" runat="server"></asp:Label>&nbsp;&nbsp;
                         <asp:Label ID="lblPropNoBUse" runat="server"></asp:Label>
                </asp:TableCell>
                <asp:TableCell runat="server" Font-Size="Small">
                    <asp:Label ID="helpPropNo" runat="server" Font-Size="Small" ForeColor="#003300"></asp:Label>
                </asp:TableCell>
            </asp:TableRow>

            <asp:TableRow ID="rowBrand" runat="server" HorizontalAlign="Left" VerticalAlign="Middle">
                <asp:TableCell runat="server" Font-Bold="True">設備廠牌</asp:TableCell>
                <asp:TableCell runat="server" Font-Size="Small">
                    授權：<asp:TextBox ID="TextBrandLcs" runat="server" ForeColor="Green"></asp:TextBox>&nbsp;&nbsp;
                    <asp:DropDownList ID="SelBrandLcs" runat="server" AutoPostBack="True" OnSelectedIndexChanged="SelBrandLcs_SelectedIndexChanged"
                        ForeColor="#009933" DataSourceID="SqlDataSourceBrandLcs" DataTextField="授權廠牌" DataValueField="授權廠牌">
                    </asp:DropDownList>
                    <asp:SqlDataSource ID="SqlDataSourceBrandLcs" runat="server" ConnectionString="<%$ ConnectionStrings:IDMSConnectionString %>"
                        SelectCommand="SELECT DISTINCT [授權廠牌] FROM [軟體授權] order by [授權廠牌]">
                    </asp:SqlDataSource><br />
                    使用：<asp:Label ID="lblBrandUse" runat="server"></asp:Label>                    
                </asp:TableCell>
                <asp:TableCell runat="server" Font-Size="Small">
                    <asp:Label ID="helpBrand" runat="server" Font-Size="Small" ForeColor="#003300"></asp:Label>
                </asp:TableCell>
            </asp:TableRow>

            <asp:TableRow ID="rowStyle" runat="server" HorizontalAlign="Left" VerticalAlign="Middle">
                <asp:TableCell runat="server" Font-Bold="True">設備型式</asp:TableCell>
                <asp:TableCell runat="server" Font-Size="Small">
                    授權：<asp:TextBox ID="TextStyleLcs" runat="server" ForeColor="Green"></asp:TextBox>&nbsp;&nbsp;
                    <asp:DropDownList ID="SelStyleLcs" runat="server" ForeColor="#009933" DataSourceID="SqlDataSourceStyleLcs"
                        DataTextField="授權型式" DataValueField="授權型式" AutoPostBack="True" OnSelectedIndexChanged="SelStyleLcs_SelectedIndexChanged">
                    </asp:DropDownList>
                    <asp:SqlDataSource ID="SqlDataSourceStyleLcs" runat="server" ConnectionString="<%$ ConnectionStrings:IDMSConnectionString %>"
                        SelectCommand="SELECT DISTINCT [授權型式] FROM [軟體授權] where [授權廠牌]=@授權廠牌 order by [授權型式]">
                        <SelectParameters>
                            <asp:ControlParameter ControlID="SelBrandLcs" Name="授權廠牌" PropertyName="SelectedValue" Type="String" />
                        </SelectParameters>
                    </asp:SqlDataSource>
                    <br />
                    使用：<asp:Label ID="lblStyleUse" runat="server"></asp:Label>
                </asp:TableCell>
                <asp:TableCell runat="server" Font-Size="Small">
                    <asp:Label ID="helpStyle" runat="server" Font-Size="Small" ForeColor="#003300"></asp:Label>
                </asp:TableCell>
            </asp:TableRow>         
            
            <asp:TableRow ID="rowCause" runat="server" HorizontalAlign="Left" VerticalAlign="Middle">
                <asp:TableCell runat="server" Font-Bold="True">減損原因</asp:TableCell>
                <asp:TableCell runat="server" Font-Size="Small">
                    <asp:TextBox ID="TextCause" runat="server" ForeColor="Green"></asp:TextBox>
                    &nbsp;&nbsp;
                    <asp:DropDownList ID="SelCause" runat="server" AutoPostBack="True" OnSelectedIndexChanged="SelCause_SelectedIndexChanged"
                        DataSourceID="SqlDataSourceCause" DataTextField="減損原因" DataValueField="減損原因" ForeColor="#009933">
                    </asp:DropDownList>
                    <asp:SqlDataSource ID="SqlDataSourceCause" runat="server" ConnectionString="<%$ ConnectionStrings:IDMSConnectionString %>"
                        SelectCommand="SELECT DISTINCT [減損原因] FROM [軟體授權] order by [減損原因]">
                    </asp:SqlDataSource>
                    <span onclick="TextCause.value=SelCause.value;"><font size="2" color="blue" style="cursor:pointer"><u>帶入</u></font></span>
                </asp:TableCell>
                <asp:TableCell runat="server" Font-Size="Small">
                    <asp:Label ID="helpCause" runat="server" Font-Size="Small" ForeColor="#003300"></asp:Label>
                </asp:TableCell>
            </asp:TableRow>

            <asp:TableRow ID="rowExec" runat="server" HorizontalAlign="Left" VerticalAlign="Middle">
                <asp:TableCell runat="server" Font-Bold="True">減損方式</asp:TableCell>
                <asp:TableCell runat="server" Font-Size="Small">
                    <asp:TextBox ID="TextExec" runat="server" ForeColor="Green"></asp:TextBox>
                    &nbsp;&nbsp;
                    <asp:DropDownList ID="SelExec" runat="server" AutoPostBack="True" OnSelectedIndexChanged="SelExec_SelectedIndexChanged"
                        DataSourceID="SqlDataSourceExec" DataTextField="減損方式" DataValueField="減損方式" ForeColor="#009933">
                    </asp:DropDownList>
                    <asp:SqlDataSource ID="SqlDataSourceExec" runat="server" ConnectionString="<%$ ConnectionStrings:IDMSConnectionString %>"
                        SelectCommand="SELECT DISTINCT [減損方式] FROM [軟體授權] order by [減損方式]">
                    </asp:SqlDataSource>
                    <span onclick="TextExec.value=SelExec.value;"><font size="2" color="blue" style="cursor:pointer"><u>帶入</u></font></span>
                </asp:TableCell>
                <asp:TableCell runat="server" Font-Size="Small">
                    <asp:Label ID="helpExec" runat="server" Font-Size="Small" ForeColor="#003300"></asp:Label>
                </asp:TableCell>
            </asp:TableRow>

            <asp:TableRow ID="rowStatusDay" runat="server" HorizontalAlign="Left" VerticalAlign="Middle">
                <asp:TableCell runat="server" Font-Bold="True">減損日期</asp:TableCell>
                <asp:TableCell runat="server" Font-Size="Small">
                    <asp:DropDownList ID="SelStatusYYYY" runat="server" ForeColor="#009933"></asp:DropDownList>年 &nbsp;&nbsp;
                    <asp:DropDownList ID="SelStatusMM" runat="server" ForeColor="#009933"></asp:DropDownList>月 &nbsp;&nbsp;
                    <asp:DropDownList ID="SelStatusDD" runat="server" ForeColor="#009933"></asp:DropDownList>日                    
                </asp:TableCell>
                <asp:TableCell runat="server" Font-Size="Small">
                    <asp:Label ID="helpStatusDay" runat="server" Font-Size="Small" ForeColor="#003300"></asp:Label>
                </asp:TableCell>
            </asp:TableRow>            
                  
            <asp:TableRow ID="rowMemo" runat="server" HorizontalAlign="Left" VerticalAlign="Middle">
                <asp:TableCell runat="server" Font-Bold="True">備註說明</asp:TableCell>
                <asp:TableCell runat="server" Font-Size="Small">
                    <asp:TextBox ID="TextMemo" runat="server" Columns="48" Width="500px" Height="100px" Rows="5" ForeColor="Green" TextMode="MultiLine"></asp:TextBox>
                </asp:TableCell>
                <asp:TableCell runat="server" Font-Size="Small">
                    <asp:Label ID="helpMemo" runat="server" Font-Size="Small" ForeColor="#003300"></asp:Label>
                </asp:TableCell>
            </asp:TableRow>

            <asp:TableRow ID="rowCreate" runat="server" HorizontalAlign="Left" VerticalAlign="Middle">
                <asp:TableCell runat="server" Font-Bold="True">資料建立</asp:TableCell>
                <asp:TableCell runat="server" Font-Size="Small">
                    <asp:Label ID="LblCreater" runat="server" Font-Size="Small"></asp:Label>
                    &nbsp;&nbsp;
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
                    &nbsp;&nbsp;
                    <asp:Label ID="LblUpdateDT" runat="server" Font-Size="Small"></asp:Label>
                </asp:TableCell>
                <asp:TableCell runat="server" Font-Size="Small">
                    <asp:Label ID="helpUpdate" runat="server" Font-Size="Small" ForeColor="#003300"></asp:Label>
                </asp:TableCell>
            </asp:TableRow>

            <asp:TableRow ID="rowChecks" runat="server" HorizontalAlign="Left" VerticalAlign="Middle">
                <asp:TableCell ID="TableCell7" runat="server" Font-Bold="True">資料查核</asp:TableCell>
                <asp:TableCell ID="TableCell8" runat="server" Font-Size="Small">
                    <asp:ValidationSummary ID="ValidationSummary1" runat="server" />　
                    <asp:Label ID="lblChecks" runat="server" ForeColor="red" Text="" Font-Bold="True"></asp:Label>
                </asp:TableCell>
                <asp:TableCell ID="TableCell9" runat="server" Font-Size="Small">
                    <asp:Label ID="helpChecks" runat="server" Font-Size="Small" ForeColor="#003300"></asp:Label>
                </asp:TableCell>
            </asp:TableRow>
        </asp:Table>
    </div>
    <br />
    <div class="style1">
        <asp:Button ID="BtnAdd"  runat="server" Text="　新　增　" OnClick="BtnAdd_Click" OnClientClick="return confirm('您確定要新增這筆資料嗎？')" ForeColor="Red" />　　
        <asp:Button ID="BtnEdit" runat="server" Text="　修　改　" OnClick="BtnEdit_Click" OnClientClick="return confirm('您確定要修改這筆資料嗎？')" />　　
        <asp:Button ID="BtnDel"  runat="server" Text="　刪　除　" OnClick="BtnDel_Click"  OnClientClick="return confirm('您確定要刪除這筆資料嗎？')" CausesValidation="False" />　　
        <asp:Button ID="BtnLife" runat="server" Text=" 生命履歷 " OnClick="BtnLife_Click" CausesValidation="False" />　
        <asp:Button ID="BtnIn" runat="server" Text=" 列印申請單 "   OnClick="BtnPrint_Click" CausesValidation="False" />　
        <asp:Button ID="BtnPre"  runat="server" Text="　預　覽　" OnClick="BtnPre_Click" CausesValidation="False" />　　
        <asp:Button ID="BtnExit" runat="server" Text="　關　閉　" OnClientClick="window.close();" CausesValidation="False" />
        <br />
        <asp:Label runat="server" Font-Size="Small" ForeColor="#006600" Text="(★：必填欄位　　紅色：連結欄位)"></asp:Label>
    </div>    
    </form>
    <asp:Label ID="Label1" runat="server" Font-Size="Small" ForeColor="#003300"></asp:Label>
</body>
</html>
