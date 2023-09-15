<%@ Page Language="C#" Title="清查編輯" AutoEventWireup="true" CodeFile="ChkEdit.aspx.cs" Inherits="SOS_ChkEdit"
    MaintainScrollPositionOnPostback="true" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>清查編輯</title>
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
<body bgcolor="#D7E6EB">
    <form id="form1" runat="server">
    <div style="text-align: center; font-family: 標楷體;">
        <asp:Label ID="LblAskEdit" runat="server" Text="機器清查編輯介面" Style="font-weight: 700;
            color: #0000CC; font-size: xx-large"></asp:Label>
    </div>
    <div class="style1">
        <asp:Table ID="Table1" runat="server" Width="100%" GridLines="Both">
            <asp:TableRow ID="rowTitle" runat="server" BackColor="#33CCCC" Font-Bold="True" ForeColor="#800000">
                <asp:TableCell runat="server" Width="150px">欄位名稱</asp:TableCell>
                <asp:TableCell runat="server" Width="40%">設定</asp:TableCell>
                <asp:TableCell runat="server">說明</asp:TableCell>
            </asp:TableRow>

            <asp:TableRow ID="rowChkYear" runat="server" HorizontalAlign="Left" VerticalAlign="Middle">
                <asp:TableCell runat="server" Font-Bold="True">清查年度(★)</asp:TableCell>
                <asp:TableCell runat="server" Font-Size="Small">
                    <asp:TextBox ID="TextChkYear" runat="server" BackColor="Silver" Columns="4" Font-Size="Large"
                        ForeColor="Red" ReadOnly="True" CssClass="style1"></asp:TextBox>
                </asp:TableCell>
                <asp:TableCell runat="server" Font-Size="Small">
                    <asp:Label ID="helpChkYear" runat="server" Font-Size="Small" ForeColor="#003300"></asp:Label>
                </asp:TableCell>
            </asp:TableRow>

            <asp:TableRow ID="rowDevNo" runat="server" HorizontalAlign="Left" VerticalAlign="Middle">
                <asp:TableCell runat="server" Font-Bold="True">設備編號(★)</asp:TableCell>
                <asp:TableCell runat="server" Font-Size="Small">
                    <asp:TextBox ID="TextDevNo" runat="server" BackColor="Silver" Columns="2" Font-Size="Large"
                        ForeColor="Red" ReadOnly="True" CssClass="style1"></asp:TextBox> &nbsp;&nbsp;
                    <asp:Button ID="BtnDev" runat="server" Text=" 設 備 " OnClick="BtnDev_Click" /> &nbsp;&nbsp;
                    <br />
                    <asp:Label ID="lblHost" runat="server"></asp:Label>
                </asp:TableCell>
                <asp:TableCell runat="server" Font-Size="Small">
                    <asp:Label ID="helpDevNo" runat="server" Font-Size="Small" ForeColor="#003300"></asp:Label>
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
                        &nbsp; &nbsp; 或 &nbsp; &nbsp;<font style="cursor: pointer; font-size: Small; color: Blue"><u>機房坐標定位</u></font>
                    </span>&nbsp;&nbsp;
                    <asp:TextBox ID="TextPointerNo" runat="server" BackColor="Silver" Columns="4" OnKeyDown="event.returnValue=false;"></asp:TextBox>                    
                </asp:TableCell>
                <asp:TableCell runat="server" Font-Size="Small">
                    <asp:Label ID="helpPlace" runat="server" Font-Size="Small" ForeColor="#003300"></asp:Label>
                </asp:TableCell>
            </asp:TableRow>

            <asp:TableRow ID="rowMt" runat="server" HorizontalAlign="Left" VerticalAlign="Middle">
                <asp:TableCell runat="server" Font-Bold="True" BackColor="">維護人員(★)</asp:TableCell>
                <asp:TableCell runat="server" Font-Size="Small">                    
                    <asp:DropDownList ID="SelHwUnit" runat="server" AutoPostBack="True" ForeColor="#009933"
                        DataSourceID="SqlDataSourceHwUnit" DataTextField="Item" DataValueField="Item">
                    </asp:DropDownList>
                    &nbsp;&nbsp;
                    <asp:SqlDataSource ID="SqlDataSourceHwUnit" runat="server" ConnectionString="<%$ ConnectionStrings:IDMSConnectionString %>"
                        SelectCommand="SELECT [Item] FROM [Config] WHERE (Kind='組織架構' and Item=@MIC) or ([Kind] = @Kind or Kind=@MIC) ORDER BY [mark]">
                        <SelectParameters>
                            <asp:Parameter DefaultValue="維護群組" Name="Kind" Type="String" />
                            <asp:Parameter DefaultValue="數值資訊組" Name="MIC" Type="String" />
                        </SelectParameters>
                    </asp:SqlDataSource>
                    <asp:DropDownList ID="SelMt" runat="server" AutoPostBack="True" ForeColor="#009933"
                        DataSourceID="SqlDataSourceMt" DataTextField="Item" DataValueField="Item">
                    </asp:DropDownList>
                    <asp:SqlDataSource ID="SqlDataSourceMt" runat="server" ConnectionString="<%$ ConnectionStrings:IDMSConnectionString %>"
                        SelectCommand="SELECT [Item] FROM [Config] WHERE ([Kind] = @Kind) ORDER BY [mark]">
                        <SelectParameters>
                            <asp:ControlParameter ControlID="SelHwUnit" Name="Kind" PropertyName="SelectedValue"
                                Type="String" />
                        </SelectParameters>
                    </asp:SqlDataSource>：                 
                    <asp:Label ID="LblMt" runat="server"></asp:Label>                    
                </asp:TableCell>
                <asp:TableCell runat="server" Font-Size="Small">
                    <asp:Label ID="helpMt" runat="server" ForeColor="#003300"></asp:Label>
                </asp:TableCell>
            </asp:TableRow>

            <asp:TableRow ID="rowChecks" runat="server" HorizontalAlign="Left" VerticalAlign="Middle">
                <asp:TableCell runat="server" Font-Bold="True">清查結果</asp:TableCell>
                <asp:TableCell runat="server" Font-Size="Small">
                    <asp:CheckBoxList ID="ChkChecks" runat="server" DataSourceID="SqlDataSourceChecks"
                        Font-Size="Small" DataTextField='Item' DataValueField="Item" RepeatDirection="Horizontal"
                        RepeatLayout="Flow" RepeatColumns="4">
                    </asp:CheckBoxList>
                    <asp:SqlDataSource ID="SqlDataSourceChecks" runat="server" ConnectionString="<%$ ConnectionStrings:IDMSConnectionString %>"
                        SelectCommand="SELECT [Item] FROM [Config] WHERE ([Kind] = @Kind) ORDER BY [mark]">
                        <SelectParameters>
                            <asp:Parameter DefaultValue="清查結果" Name="Kind" Type="String" />
                        </SelectParameters>
                    </asp:SqlDataSource>
                </asp:TableCell>
                <asp:TableCell runat="server" Font-Size="Small">
                    <asp:Label ID="helpChecks" runat="server" Font-Size="Small" ForeColor="#003300"></asp:Label>
                </asp:TableCell>
            </asp:TableRow>

            <asp:TableRow ID="rowMemo" runat="server" HorizontalAlign="Left" VerticalAlign="Middle">
                <asp:TableCell runat="server" Font-Bold="True">
                    <span onclick="window.open('../Help/IDMS機器清查介面-需求與流程.doc','_blank');">
                        <font style="cursor: pointer; font-size: Small; color: Blue"><u>備註說明</u></font>
                    </span>
                </asp:TableCell>
                <asp:TableCell runat="server" Font-Size="Small">
                    <asp:TextBox ID="TextMemo" runat="server" Columns="48" Width="500px" Height="100px"
                        Rows="5" TextMode="MultiLine"></asp:TextBox>
                </asp:TableCell>
                <asp:TableCell ID="TableCell3" runat="server" Font-Size="Small">
                    <asp:Label ID="helpMemo" runat="server" Font-Size="Small" ForeColor="#003300"></asp:Label>
                </asp:TableCell>
            </asp:TableRow>            
            
            <asp:TableRow ID="rowStatus" runat="server" HorizontalAlign="Left">
                <asp:TableCell runat="server" Font-Bold="True">清查狀態(★)</asp:TableCell>
                <asp:TableCell runat="server" Font-Size="Small">
                    <asp:DropDownList ID="SelStatus" runat="server" ForeColor="#009933"
                        DataSourceID="SqlDataSourceStatus" DataTextField="Item" DataValueField="Item">
                    </asp:DropDownList>
                    <asp:SqlDataSource ID="SqlDataSourceStatus" runat="server" ConnectionString="<%$ ConnectionStrings:IDMSConnectionString %>"
                        SelectCommand="SELECT [Item] FROM [Config] WHERE [Kind] = '清查狀態' ORDER BY [mark]">
                    </asp:SqlDataSource>
                </asp:TableCell>
                <asp:TableCell runat="server" Font-Size="Small">
                    <asp:Label ID="helpStatus" runat="server" Font-Size="Small" ForeColor="#003300"></asp:Label>
                </asp:TableCell>
            </asp:TableRow>            
            
            <asp:TableRow ID="rowCheck" runat="server" HorizontalAlign="Left" VerticalAlign="Middle">
                <asp:TableCell runat="server" Font-Bold="True">資料清查</asp:TableCell>
                <asp:TableCell runat="server" Font-Size="Small">
                    <asp:Label ID="LblChecker" runat="server" Font-Size="Small"></asp:Label>
                    &nbsp;&nbsp;
                    <asp:Label ID="LblCheckDT" runat="server" Font-Size="Small"></asp:Label>
                </asp:TableCell>
                <asp:TableCell runat="server" Font-Size="Small">
                    <asp:Label ID="helpCheck" runat="server" Font-Size="Small" ForeColor="#003300"></asp:Label>
                </asp:TableCell>
            </asp:TableRow>

            <asp:TableRow ID="rowUpdate" runat="server" HorizontalAlign="Left" VerticalAlign="Middle">
                <asp:TableCell runat="server" Font-Bold="True">資料更新</asp:TableCell>
                <asp:TableCell runat="server" Font-Size="Small">
                    <asp:Label ID="LblUpdater" runat="server" Font-Size="Small"></asp:Label>
                    &nbsp;&nbsp;
                    <asp:Label ID="LblUpdateDT" runat="server" Font-Size="Small"></asp:Label>
                </asp:TableCell>
                <asp:TableCell runat="server" Font-Size="Small">
                    <asp:Label ID="helpUpdate" runat="server" Font-Size="Small" ForeColor="#003300"></asp:Label>
                </asp:TableCell>
            </asp:TableRow>
        </asp:Table>
    </div>

    <div class="style1">
        <br />
        <asp:Button ID="BtnAdd" runat="server" Text="　新　增　" OnClick="BtnAdd_Click" OnClientClick="return confirm('您確定要新增這筆資料嗎？')"
            ForeColor="Red" /> &nbsp;&nbsp;&nbsp;&nbsp;
        <asp:Button ID="BtnEdit" runat="server" Text="　修　改　" OnClick="BtnEdit_Click" OnClientClick="return confirm('您確定要修改這筆資料嗎？')" />&nbsp;&nbsp;&nbsp;&nbsp;
        <asp:Button ID="BtnDel" runat="server" Text="　刪　除　" OnClick="BtnDel_Click" OnClientClick="return confirm('您確定要刪除這筆資料嗎？')"
            CausesValidation="False" />&nbsp;&nbsp;&nbsp;&nbsp;
        <asp:Button ID="BtnExit" runat="server" Text="　關　閉　" OnClientClick="window.close();"
            CausesValidation="False" />
        <br />
    </div>
    </form>
</body>
</html>
