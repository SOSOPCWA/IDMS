<%@ Page Language="C#" Title="設備迴路" AutoEventWireup="true"  Trace="False"
CodeFile="TreeEdit.aspx.cs" Inherits="Device_TreeEdit" MaintainScrollPositionOnPostback="true" %>

    <!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
    <html xmlns="http://www.w3.org/1999/xhtml">

    <head runat="server">
        <style type="text/css">
            .style1 {
                text-align: center;
            }
            
            .style2 {
                font-size: Small;
            }
        </style>
    </head>

    <body bgcolor="LightGray">
        <form id="form1" runat="server">
            <div style="text-align: center; font-family: 標楷體;">
                <asp:Label ID="LblDevEdit" runat="server" Text="設備迴路設定介面" Style="font-weight: 700;
            color: #0000CC; font-size: xx-large"></asp:Label>
            </div>
            <div class="style1">
                <asp:Table ID="Table1" runat="server" Width="100%" GridLines="Both">
                    <asp:TableRow ID="rowDevice" runat="server" HorizontalAlign="Left" VerticalAlign="Top">
                        <asp:TableCell runat="server" Font-Size="Small" ColumnSpan="3">
                            <b><font size="4">設備編號：</font></b>
                            <asp:TextBox ID="TextDevNo" Text="0" runat="server" BackColor="Silver" Columns="2" Font-Size="Large" ForeColor="Red" ReadOnly="True" CssClass="style1"></asp:TextBox>&nbsp;&nbsp;
                            <asp:Label ID="lblDevName" runat="server" Font-Size="Large" ForeColor="Green"></asp:Label>&nbsp;&nbsp;
                            <asp:Label ID="lblDevice" runat="server" Font-Size="Small" ForeColor="Green"></asp:Label>&nbsp;&nbsp;
                            <asp:Label ID="lblMt" runat="server" Font-Size="Small" ForeColor="Green"></asp:Label> <br />
                            <asp:Label ID="lblTree" runat="server" Font-Size="Small"></asp:Label>
                        </asp:TableCell>
                        <asp:TableCell runat="server" Font-Size="Small" Width="40%">
                            <asp:Label ID="helpDevice" runat="server" Font-Size="Small" ForeColor="#003300"></asp:Label>
                        </asp:TableCell>
                    </asp:TableRow>

                    <asp:TableRow runat="server" HorizontalAlign="Center" VerticalAlign="Middle">
                        <asp:TableCell runat="server" Font-Size="Small">
                            <b><font size="4">迴路待選清單：</font></b>
                        </asp:TableCell>
                        <asp:TableCell runat="server" Font-Size="Small" ColumnSpan="2">
                            <b><font size="4">接電：<font color="red">上游供電迴路</font></font></b>
                        </asp:TableCell>
                        <asp:TableCell runat="server" Font-Size="Small" RowSpan="4" HorizontalAlign="Left" VerticalAlign="Top">
                            <asp:Label ID="helpPower" runat="server" Font-Size="Small" ForeColor="#003300"></asp:Label>
                        </asp:TableCell>
                    </asp:TableRow>

                    <asp:TableRow runat="server" HorizontalAlign="Center" VerticalAlign="Middle">
                        <asp:TableCell runat="server" Font-Size="Small" RowSpan="11">
                            <asp:DropDownList ID="SelPlace" runat="server" AutoPostBack="True" OnSelectedIndexChanged="SelPlace_SelectedIndexChanged" ForeColor="#009933" DataSourceID="SqlDataSourcePlace" DataTextField="區域名稱" DataValueField="區域名稱">
                                <asp:ListItem>空調群組</asp:ListItem>
                            </asp:DropDownList>

                            <asp:SqlDataSource ID="SqlDataSourcePlace" runat="server" ConnectionString="<%$ ConnectionStrings:IDMSConnectionString %>" SelectCommand="SELECT DISTINCT [區域名稱] FROM [定位設定] WHERE ([定位方式] &lt;&gt; @定位方式) order by [區域名稱]">
                                <SelectParameters>
                                    <asp:Parameter DefaultValue="坐標" Name="定位方式" Type="String" />
                                </SelectParameters>
                            </asp:SqlDataSource>&nbsp;

                            <asp:DropDownList ID="SelPointer" runat="server" AutoPostBack="True" OnSelectedIndexChanged="SelPointer_SelectedIndexChanged" ForeColor="#009933" DataSourceID="SqlDataSourcePointer" DataTextField="定位名稱" DataValueField="定位編號">
                            </asp:DropDownList>
                            <asp:SqlDataSource ID="SqlDataSourcePointer" runat="server" ConnectionString="<%$ ConnectionStrings:IDMSConnectionString %>">
                            </asp:SqlDataSource><br />

                            <asp:ListBox ID="ListSource" runat="server" ForeColor="#009933" DataSourceID="SqlDataSourceSource" AutoPostBack="true" OnSelectedIndexChanged="ListSource_SelectedIndexChanged" DataTextField="設備名稱" DataValueField="設備編號" AppendDataBoundItems="False" Rows="24">
                            </asp:ListBox>
                            <asp:SqlDataSource ID="SqlDataSourceSource" runat="server" ConnectionString="<%$ ConnectionStrings:IDMSConnectionString %>">
                            </asp:SqlDataSource><br />

                            <asp:LinkButton ID="LinkSourceUp" runat="server" Font-Size="Small" OnClick="LinkSourceUp_Click" Text="上游" ToolTip="依所選取迴路列出其上游"></asp:LinkButton>&nbsp;&nbsp;
                            <asp:LinkButton ID="LinkSourceDn" runat="server" Font-Size="Small" OnClick="LinkSourceDn_Click" Text="下游" ToolTip="依所選取迴路列出其下游"></asp:LinkButton>&nbsp;&nbsp;
                            <asp:LinkButton ID="LinkSourceEdit" runat="server" Font-Size="Small" OnClick="LinkSourceEdit_Click" Text="編輯" ToolTip="到所選取迴路之編輯介面"></asp:LinkButton>&nbsp;&nbsp;
                            <asp:LinkButton ID="LinkSourceConn" runat="server" Font-Size="Small" OnClick="LinkSourceConn_Click" Text="迴路" ToolTip="到所選取迴路之設定介面"></asp:LinkButton>
                            <br />
                            <asp:Label ID="lblSource" runat="server" Font-Size="Small" Text=""></asp:Label>
                        </asp:TableCell>

                        <asp:TableCell runat="server" Font-Size="Small">
                            <asp:LinkButton ID="LinkUpPowerAdd" runat="server" Font-Size="Small" OnClick="LinkUpPowerAdd_Click" Text=">> 新增 >>" ToolTip="新增左方所選取之接電迴路到右方上游"></asp:LinkButton><br /><br />
                            <asp:LinkButton ID="LinkUpPowerDel" runat="server" Font-Size="Small" OnClick="LinkUpPowerDel_Click" Text="<< 刪除 <<" ToolTip="刪除右方上游所選取之接電迴路"></asp:LinkButton><br /><br />
                            <asp:LinkButton ID="LinkUpPowerEdit" runat="server" Font-Size="Small" OnClick="LinkUpPowerEdit_Click" Text="→ 編輯 →" ToolTip="編輯右方上游所選取之接電迴路"></asp:LinkButton><br /><br />
                            <asp:LinkButton ID="LinkUpPowerConn" runat="server" Font-Size="Small" OnClick="LinkUpPowerConn_Click" Text="→ 迴路 →" ToolTip="設定右方上游所選取之接電迴路"></asp:LinkButton>
                        </asp:TableCell>

                        <asp:TableCell runat="server" Font-Size="Small">
                            <asp:ListBox ID="ListUpPower" runat="server" AutoPostBack="true" OnSelectedIndexChanged="ListUpPower_SelectedIndexChanged" ForeColor="Green">
                            </asp:ListBox><br />
                            <asp:Label ID="lblUpPower" runat="server" Font-Size="Small" Text=""></asp:Label>
                        </asp:TableCell>
                    </asp:TableRow>

                    <asp:TableRow runat="server" HorizontalAlign="Center" VerticalAlign="Middle">
                        <asp:TableCell runat="server" Font-Size="Small" ColumnSpan="2">
                            <b><font size="4">接電：<font color="red">下游用電設備</font></font></b>
                        </asp:TableCell>
                    </asp:TableRow>

                    <asp:TableRow runat="server" HorizontalAlign="Center" VerticalAlign="Middle">
                        <asp:TableCell runat="server" Font-Size="Small">
                            <asp:LinkButton ID="LinkDnPowerAdd" runat="server" Font-Size="Small" OnClick="LinkDnPowerAdd_Click" Text=">> 新增 >>" ToolTip="新增左方所選取之接電迴路到右方下游"></asp:LinkButton><br /><br />
                            <asp:LinkButton ID="LinkDnPowerDel" runat="server" Font-Size="Small" OnClick="LinkDnPowerDel_Click" Text="<< 刪除 <<" ToolTip="刪除右方下游所選取之接電迴路"></asp:LinkButton><br /><br />
                            <asp:LinkButton ID="LinkDnPowerEdit" runat="server" Font-Size="Small" OnClick="LinkDnPowerEdit_Click" Text="→ 編輯 →" ToolTip="編輯右方下游所選取之接電迴路"></asp:LinkButton><br /><br />
                            <asp:LinkButton ID="LinkDnPowerConn" runat="server" Font-Size="Small" OnClick="LinkDnPowerConn_Click" Text="→ 迴路 →" ToolTip="設定右方下游所選取之接電迴路"></asp:LinkButton>
                        </asp:TableCell>

                        <asp:TableCell runat="server" Font-Size="Small">
                            <asp:ListBox ID="ListDnPower" runat="server" AutoPostBack="true" OnSelectedIndexChanged="ListDnPower_SelectedIndexChanged" ForeColor="Green">
                            </asp:ListBox>
                            <br />
                            <asp:Label ID="lblDnPower" runat="server" Font-Size="Small" Text=""></asp:Label>
                        </asp:TableCell>
                    </asp:TableRow>

                    <asp:TableRow ID="rowNet" runat="server" HorizontalAlign="Center" VerticalAlign="Middle">
                        <asp:TableCell runat="server" Font-Size="Small" ColumnSpan="2">
                            <b><font size="4">接網：<font color="red">上游供網迴路</font></font></b>
                        </asp:TableCell>

                        <asp:TableCell runat="server" Font-Size="Small" HorizontalAlign="Left" VerticalAlign="Top" RowSpan="4">
                            <asp:Label ID="helpNet" runat="server" Font-Size="Small" ForeColor="#003300"></asp:Label>
                        </asp:TableCell>
                    </asp:TableRow>

                    <asp:TableRow runat="server" HorizontalAlign="Center" VerticalAlign="Middle">
                        <asp:TableCell runat="server" Font-Size="Small">
                            <asp:LinkButton ID="LinkUpNetAdd" runat="server" Font-Size="Small" OnClick="LinkUpNetAdd_Click" Text=">> 新增 >>" ToolTip="新增左方所選取之接網迴路到右方上游"></asp:LinkButton><br /><br />
                            <asp:LinkButton ID="LinkUpNetDel" runat="server" Font-Size="Small" OnClick="LinkUpNetDel_Click" Text="<< 刪除 <<" ToolTip="刪除右方上游所選取之接網迴路"></asp:LinkButton><br /><br />
                            <asp:LinkButton ID="LinkUpNetEdit" runat="server" Font-Size="Small" OnClick="LinkUpNetEdit_Click" Text="→ 編輯 →" ToolTip="編輯右方上游所選取之接網迴路"></asp:LinkButton><br /><br />
                            <asp:LinkButton ID="LinkUpNetConn" runat="server" Font-Size="Small" OnClick="LinkUpNetConn_Click" Text="→ 迴路 →" ToolTip="設定右方上游所選取之接網迴路"></asp:LinkButton>
                        </asp:TableCell>

                        <asp:TableCell runat="server" Font-Size="Small">
                            <asp:ListBox ID="ListUpNet" runat="server" AutoPostBack="true" OnSelectedIndexChanged="ListUpNet_SelectedIndexChanged" ForeColor="Green">
                            </asp:ListBox><br />
                            <asp:Label ID="lblUpNet" runat="server" Font-Size="Small" Text="" />
                        </asp:TableCell>
                    </asp:TableRow>

                    <asp:TableRow runat="server" HorizontalAlign="Center" VerticalAlign="Middle">
                        <asp:TableCell runat="server" Font-Size="Small" ColumnSpan="2">
                            <b><font size="4">接網：<font color="red">下游用網設備</font></font></b>
                        </asp:TableCell>
                    </asp:TableRow>

                    <asp:TableRow runat="server" HorizontalAlign="Center" VerticalAlign="Middle">
                        <asp:TableCell runat="server" Font-Size="Small">
                            <asp:LinkButton ID="LinkDnNetAdd" runat="server" Font-Size="Small" OnClick="LinkDnNetAdd_Click" Text=">> 新增 >>" ToolTip="新增左方所選取之接網迴路到右方下游"></asp:LinkButton><br /><br />
                            <asp:LinkButton ID="LinkDnNetDel" runat="server" Font-Size="Small" OnClick="LinkDnNetDel_Click" Text="<< 刪除 <<" ToolTip="刪除右方下游所選取之接網迴路"></asp:LinkButton><br /><br />
                            <asp:LinkButton ID="LinkDnNetEdit" runat="server" Font-Size="Small" OnClick="LinkDnNetEdit_Click" Text="→ 編輯 →" ToolTip="編輯右方下游所選取之接網迴路"></asp:LinkButton><br /><br />
                            <asp:LinkButton ID="LinkDnNetConn" runat="server" Font-Size="Small" OnClick="LinkDnNetConn_Click" Text="→ 迴路 →" ToolTip="設定右方下游所選取之接網迴路"></asp:LinkButton>
                        </asp:TableCell>

                        <asp:TableCell runat="server" Font-Size="Small">
                            <asp:ListBox ID="ListDnNet" runat="server" AutoPostBack="true" OnSelectedIndexChanged="ListDnNet_SelectedIndexChanged" ForeColor="Green">
                            </asp:ListBox><br />
                            <asp:Label ID="lblDnNet" runat="server" Font-Size="Small" Text="" />
                        </asp:TableCell>
                    </asp:TableRow>

                    <asp:TableRow ID="rowCold" runat="server" HorizontalAlign="Center" VerticalAlign="Middle">
                        <asp:TableCell runat="server" Font-Size="Small" ColumnSpan="2">
                            <b><font size="4">接冷：<font color="red">上游供冷迴路</font></font></b>
                        </asp:TableCell>

                        <asp:TableCell runat="server" Font-Size="Small" HorizontalAlign="Left" VerticalAlign="Top" RowSpan="4">
                            <asp:Label ID="helpCold" runat="server" Font-Size="Small" ForeColor="#003300"></asp:Label>
                        </asp:TableCell>
                    </asp:TableRow>

                    <asp:TableRow runat="server" HorizontalAlign="Center" VerticalAlign="Middle">
                        <asp:TableCell runat="server" Font-Size="Small">
                            <asp:LinkButton ID="LinkUpColdAdd" runat="server" Font-Size="Small" OnClick="LinkUpColdAdd_Click" Text=">> 新增 >>" ToolTip="新增左方所選取之接網迴路到右方上游"></asp:LinkButton><br /><br />
                            <asp:LinkButton ID="LinkUpColdDel" runat="server" Font-Size="Small" OnClick="LinkUpColdDel_Click" Text="<< 刪除 <<" ToolTip="刪除右方上游所選取之接網迴路"></asp:LinkButton><br /><br />
                            <asp:LinkButton ID="LinkUpColdEdit" runat="server" Font-Size="Small" OnClick="LinkUpColdEdit_Click" Text="→ 編輯 →" ToolTip="編輯右方上游所選取之接網迴路"></asp:LinkButton><br /><br />
                            <asp:LinkButton ID="LinkUpColdConn" runat="server" Font-Size="Small" OnClick="LinkUpColdConn_Click" Text="→ 迴路 →" ToolTip="設定右方上游所選取之接網迴路"></asp:LinkButton>
                        </asp:TableCell>

                        <asp:TableCell runat="server" Font-Size="Small">
                            <asp:ListBox ID="ListUpCold" runat="server" AutoPostBack="true" OnSelectedIndexChanged="ListUpCold_SelectedIndexChanged" ForeColor="Green">
                            </asp:ListBox><br />
                            <asp:Label ID="lblUpCold" runat="server" Font-Size="Small" Text="" />
                        </asp:TableCell>
                    </asp:TableRow>

                    <asp:TableRow runat="server" HorizontalAlign="Center" VerticalAlign="Middle">
                        <asp:TableCell runat="server" Font-Size="Small" ColumnSpan="2">
                            <b><font size="4">接冷：<font color="red">下游用冷設備</font></font></b>
                        </asp:TableCell>
                    </asp:TableRow>

                    <asp:TableRow runat="server" HorizontalAlign="Center" VerticalAlign="Middle">
                        <asp:TableCell runat="server" Font-Size="Small">
                            <asp:LinkButton ID="LinkDnColdAdd" runat="server" Font-Size="Small" OnClick="LinkDnColdAdd_Click" Text=">> 新增 >>" ToolTip="新增左方所選取之接網迴路到右方下游"></asp:LinkButton><br /><br />
                            <asp:LinkButton ID="LinkDnColdDel" runat="server" Font-Size="Small" OnClick="LinkDnColdDel_Click" Text="<< 刪除 <<" ToolTip="刪除右方下游所選取之接網迴路"></asp:LinkButton><br /><br />
                            <asp:LinkButton ID="LinkDnColdEdit" runat="server" Font-Size="Small" OnClick="LinkDnColdEdit_Click" Text="→ 編輯 →" ToolTip="編輯右方下游所選取之接網迴路"></asp:LinkButton><br /><br />
                            <asp:LinkButton ID="LinkDnColdConn" runat="server" Font-Size="Small" OnClick="LinkDnColdConn_Click" Text="→ 迴路 →" ToolTip="設定右方下游所選取之接網迴路"></asp:LinkButton>
                        </asp:TableCell>

                        <asp:TableCell runat="server" Font-Size="Small">
                            <asp:ListBox ID="ListDnCold" runat="server" AutoPostBack="true" OnSelectedIndexChanged="ListDnCold_SelectedIndexChanged" ForeColor="Green">
                            </asp:ListBox><br />
                            <asp:Label ID="lblDnCold" runat="server" Font-Size="Small" Text="" />
                        </asp:TableCell>
                    </asp:TableRow>
                </asp:Table>
            </div>
            <br />
            <div class="style1">
                <asp:Button ID="BtnEdit" runat="server" Text=" 回編輯頁 " OnClick="BtnEdit_Click" CausesValidation="False" />&nbsp;&nbsp;&nbsp;&nbsp;
                <asp:Button ID="BtnLife" runat="server" Text=" 生命履歷 " OnClick="BtnLife_Click" CausesValidation="False" />&nbsp;&nbsp;&nbsp;&nbsp;
                <asp:Button ID="BtnExit" runat="server" Text=" 關閉介面 " OnClientClick="window.close();" CausesValidation="False" />
            </div>
        </form>
    </body>

    </html>