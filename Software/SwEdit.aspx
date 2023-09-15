<%@ Page Language="C#" Title="軟體編輯" AutoEventWireup="true" CodeFile="SwEdit.aspx.cs" Inherits="Software_SwEdit" Debug="true" MaintainScrollPositionOnPostback="true" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
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
        <asp:Label ID="LblSwEdit" runat="server" Text="軟體保管單編輯介面" Style="font-weight: 700;
            color: #0000CC; font-size: xx-large"></asp:Label>
    </div>
    <div class="style1">
        <asp:Table ID="Table1" runat="server" Width="100%" GridLines="Both">
            <asp:TableRow ID="rowTitle" runat="server" Font-Bold="True" ForeColor="#800000">
                <asp:TableCell ID="TableCell4" runat="server" Width="150px">欄位名稱</asp:TableCell>
                <asp:TableCell ID="TableCell5" runat="server">
                    設定 (<asp:LinkButton ID="LinkHideHelp" runat="server" Font-Size="Small" ForeColor="Blue" OnClick="LinkHideHelp_Click" ToolTip="隱藏右方說明欄位，以簡化操作介面">隱藏說明</asp:LinkButton>)
                </asp:TableCell>
                <asp:TableCell ID="TableCell10" runat="server" Width="35%">說明</asp:TableCell>
            </asp:TableRow>

            <asp:TableRow ID="rowSwNo" runat="server" HorizontalAlign="Left" VerticalAlign="Middle">
                <asp:TableCell runat="server" Font-Bold="True" ForeColor="Red">軟體編號</asp:TableCell>
                <asp:TableCell runat="server" Font-Size="Small">
                    <asp:TextBox ID="TextSwNo" runat="server" BackColor="Silver" Columns="2" Font-Size="Large"
                        ForeColor="Red" ReadOnly="True" CssClass="style1"></asp:TextBox>
                    &nbsp;&nbsp;表單編號：
                    <asp:DropDownList ID="SelFormYYYY" runat="server" AutoPostBack="True" ForeColor="#009933"></asp:DropDownList>
                    －
                    <asp:TextBox ID="TextFormXXXX" runat="server" CssClass="style1" Columns="2" MaxLength="4" ForeColor="Green" OnKeyDown="if(event.keyCode <=57 | event.keyCode >=96 & event.keyCode <=105 | event.keyCode ==110 | event.keyCode ==190) {event.returnValue=true;} else{event.returnValue=false;}"></asp:TextBox>
                    &nbsp;&nbsp;
                    <asp:LinkButton ID="LinkLcs" runat="server" Font-Size="Small" OnClick="LinkLcs_Click">授權查詢</asp:LinkButton>
                    <br /><br /><br />
                    批次修改：<asp:TextBox ForeColor="Green" OnKeyDown="if(event.keyCode <=57 | event.keyCode >=96 & event.keyCode <=105 | event.keyCode ==110 | event.keyCode ==190) {event.returnValue=true;} else{event.returnValue=false;}"
                        ID="TextNewSwNo" runat="server" Columns="6"> </asp:TextBox>(請填新軟體編號)&nbsp;&nbsp;
                    <asp:LinkButton ID="LinkNewSwNo" runat="server" Font-Size="Small" OnClick="LinkNewSwNo_Click">立刻執行</asp:LinkButton>
                </asp:TableCell>
                <asp:TableCell runat="server" Font-Size="Small">
                    <asp:Label ID="helpSwNo" runat="server" Font-Size="Small" ForeColor="#003300"></asp:Label>
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

            <asp:TableRow ID="rowSwAsk" runat="server" HorizontalAlign="Left" VerticalAlign="Middle">
                <asp:TableCell runat="server" Font-Bold="True">軟體申請</asp:TableCell>
                <asp:TableCell runat="server" Font-Size="Small">
                    單位：<asp:DropDownList ID="SelUnit" runat="server" ForeColor="#009933" DataSourceID="SqlDataSourceUnit" DataTextField="Item" DataValueField="Item"
                         AutoPostBack="True" OnSelectedIndexChanged="SelUnit_SelectedIndexChanged">
                    </asp:DropDownList>
                    <asp:SqlDataSource ID="SqlDataSourceUnit" runat="server" ConnectionString="<%$ ConnectionStrings:IDMSConnectionString %>"
                        SelectCommand="SELECT [Item] FROM [Config] WHERE [Kind]='數值資訊組' order by [Mark]">
                    </asp:SqlDataSource>&nbsp;&nbsp;
                    人員：<asp:TextBox ID="TextAsker" runat="server" ForeColor="Green"></asp:TextBox>&nbsp;
                    <asp:DropDownList ID="SelAsker" runat="server" ForeColor="#009933" DataSourceID="SqlDataSourceAsker"
                        DataTextField="Item" DataValueField="Item" AutoPostBack="True" OnSelectedIndexChanged="SelAsker_SelectedIndexChanged">
                    </asp:DropDownList>
                    <asp:SqlDataSource ID="SqlDataSourceAsker" runat="server" ConnectionString="<%$ ConnectionStrings:IDMSConnectionString %>"
                        SelectCommand="SELECT [Item] FROM [Config] WHERE ([Kind] = @Kind) ORDER BY [mark]">
                        <SelectParameters>
                            <asp:ControlParameter ControlID="SelUnit" Name="Kind" PropertyName="SelectedValue" Type="String" />
                        </SelectParameters>
                    </asp:SqlDataSource>
                </asp:TableCell>
                <asp:TableCell runat="server" Font-Size="Small">
                    <asp:Label ID="helpSwAsk" runat="server" Font-Size="Small" ForeColor="#003300"></asp:Label>
                </asp:TableCell>
            </asp:TableRow>            

            <asp:TableRow ID="rowSwName" runat="server" HorizontalAlign="Left" VerticalAlign="Middle">
                <asp:TableCell runat="server" Font-Bold="True">軟體名稱(★)</asp:TableCell>
                <asp:TableCell runat="server" Font-Size="Small">
                    <asp:TextBox ID="TextSwName" runat="server" Width="250px" ForeColor="Green"></asp:TextBox>
                    <asp:RequiredFieldValidator ID="RequiredFieldValidatorSwName" runat="server" ControlToValidate="TextSwName"
                        ErrorMessage="未輸入軟體名稱" Font-Size="Small" SetFocusOnError="True" Display="Dynamic"></asp:RequiredFieldValidator>
                    &nbsp;&nbsp;
                    <asp:DropDownList ID="SelSwName" runat="server" AutoPostBack="True" OnSelectedIndexChanged="SelSwName_SelectedIndexChanged"
                        ForeColor="#009933" DataSourceID="SqlDataSourceSwName" DataTextField="軟體名稱" DataValueField="軟體名稱">
                    </asp:DropDownList>
                    <asp:SqlDataSource ID="SqlDataSourceSwName" runat="server" ConnectionString="<%$ ConnectionStrings:IDMSConnectionString %>"
                        SelectCommand="SELECT DISTINCT [軟體名稱] FROM [軟體主檔] order by [軟體名稱]">
                    </asp:SqlDataSource> 
                    <span onclick="TextSwName.value=SelSwName.value;"><font size="2" color="blue" style="cursor:pointer"><u>帶入</u></font></span>                   
                </asp:TableCell>
                <asp:TableCell runat="server" Font-Size="Small">
                    <asp:Label ID="helpSwName" runat="server" Font-Size="Small" ForeColor="#003300"></asp:Label>
                </asp:TableCell>
            </asp:TableRow>  
                     
            <asp:TableRow ID="rowVersion" runat="server" HorizontalAlign="Left" VerticalAlign="Middle">
                <asp:TableCell runat="server" Font-Bold="True">軟體版本</asp:TableCell>
                <asp:TableCell runat="server" Font-Size="Small">
                    購買：<asp:TextBox ID="TextVersBuy" Width="250px" runat="server" ForeColor="Green"></asp:TextBox> &nbsp;&nbsp;
                    <asp:DropDownList ID="SelVersBuy" runat="server" ForeColor="#009933" DataSourceID="SqlDataSourceVersBuy"
                        DataTextField="購買版本" DataValueField="購買版本" AutoPostBack="True" OnSelectedIndexChanged="SelVersBuy_SelectedIndexChanged">
                    </asp:DropDownList>
                    <asp:SqlDataSource ID="SqlDataSourceVersBuy" runat="server" ConnectionString="<%$ ConnectionStrings:IDMSConnectionString %>"
                        SelectCommand="SELECT DISTINCT [購買版本] FROM [軟體主檔] where [軟體名稱]=@軟體名稱 order by [購買版本]">
                        <SelectParameters>
                            <asp:ControlParameter ControlID="SelSwName" Name="軟體名稱" PropertyName="SelectedValue" Type="String" />
                        </SelectParameters>
                    </asp:SqlDataSource>
                    <span onclick="TextVersBuy.value=SelVersBuy.value;"><font size="2" color="blue" style="cursor:pointer"><u>帶入</u></font></span>
                    <br />
                    可用：<asp:TextBox ID="TextVersCan" runat="server" Columns="48" Width="250px" Height="100px" Rows="5" TextMode="MultiLine" ForeColor="Green"></asp:TextBox>&nbsp;&nbsp;
                    <asp:DropDownList ID="SelVersCan" runat="server" ForeColor="#009933" DataSourceID="SqlDataSourceVersCan"
                        DataTextField="可用版本" DataValueField="可用版本" AutoPostBack="True" OnSelectedIndexChanged="SelVersCan_SelectedIndexChanged">
                    </asp:DropDownList>
                    <asp:SqlDataSource ID="SqlDataSourceVersCan" runat="server" ConnectionString="<%$ ConnectionStrings:IDMSConnectionString %>"
                        SelectCommand="SELECT DISTINCT [可用版本] FROM [軟體主檔] where [軟體名稱]=@軟體名稱 order by [可用版本]">
                        <SelectParameters>
                            <asp:ControlParameter ControlID="SelSwName" Name="軟體名稱" PropertyName="SelectedValue" Type="String" />
                        </SelectParameters>
                    </asp:SqlDataSource>
                    <span onclick="TextVersCan.value=SelVersCan.value;"><font size="2" color="blue" style="cursor:pointer"><u>帶入</u></font></span>
                </asp:TableCell>
                <asp:TableCell runat="server" Font-Size="Small">
                    <asp:Label ID="helpVersion" runat="server" Font-Size="Small" ForeColor="#003300"></asp:Label>
                </asp:TableCell>
            </asp:TableRow>

            <asp:TableRow ID="rowLicense" runat="server" HorizontalAlign="Left" VerticalAlign="Middle">
                <asp:TableCell runat="server" Font-Bold="True">軟體授權</asp:TableCell>
                <asp:TableCell runat="server" Font-Size="Small">
                    方式：<asp:DropDownList ID="SelLcsKind" runat="server" ForeColor="#009933"
                        DataSourceID="SqlDataSourceLcsKind" DataTextField="Item" DataValueField="Item">
                    </asp:DropDownList>
                    <asp:SqlDataSource ID="SqlDataSourceLcsKind" runat="server" ConnectionString="<%$ ConnectionStrings:IDMSConnectionString %>"
                        SelectCommand="SELECT [Item] FROM [Config] WHERE [Kind] = '授權方式' ORDER BY [mark]">
                    </asp:SqlDataSource>
                    &nbsp;&nbsp;
                    數量：<asp:TextBox OnKeyDown="if(event.keyCode <=57 | event.keyCode >=96 & event.keyCode <=105 | event.keyCode ==110 | event.keyCode ==190) {event.returnValue=true;} else{event.returnValue=false;}"
                        ID="TextLcsNum" runat="server" Columns="6" MaxLength="4" ForeColor="Green"> </asp:TextBox>
                    <asp:RangeValidator ID="RangeValidatorLcsNum" runat="server" ControlToValidate="TextLcsNum"
                        Display="Dynamic" ErrorMessage="授權數量輸入格式有誤" Font-Size="Small" MaximumValue="1000"
                        MinimumValue="0" SetFocusOnError="True" Type="Integer">
                    </asp:RangeValidator>
                    <asp:RequiredFieldValidator ID="RequiredFieldValidatorLcsNum" runat="server" ControlToValidate="TextLcsNum"
                        ErrorMessage="未輸入授權數量" Font-Size="Small" SetFocusOnError="True" Display="Dynamic">
                    </asp:RequiredFieldValidator>  <br /><br />
                    說明：<asp:TextBox ID="TextLcsMemo" Width="350px" runat="server" ForeColor="Green"></asp:TextBox>
                    <br /><br />
                    申請授權(方式*數量)：<asp:Label ID="lblLicense" runat="server" Font-Size="Small"></asp:Label>                  
                </asp:TableCell>
                <asp:TableCell runat="server" Font-Size="Small">
                    <asp:Label ID="helpLicense" runat="server" Font-Size="Small" ForeColor="#003300"></asp:Label>
                </asp:TableCell>
            </asp:TableRow>

            <asp:TableRow ID="rowSource" runat="server" HorizontalAlign="Left" VerticalAlign="Middle">
                <asp:TableCell runat="server" Font-Bold="True">軟體來源</asp:TableCell>
                <asp:TableCell runat="server" Font-Size="Small">
                    <asp:TextBox ID="TextSource" runat="server" ForeColor="Green"></asp:TextBox>
                    &nbsp;&nbsp;
                    <asp:DropDownList ID="SelSource" runat="server" AutoPostBack="True" OnSelectedIndexChanged="SelSource_SelectedIndexChanged"
                        DataSourceID="SqlDataSourceSource" DataTextField="軟體來源" DataValueField="軟體來源" ForeColor="#009933">
                    </asp:DropDownList>
                    <asp:SqlDataSource ID="SqlDataSourceSource" runat="server" ConnectionString="<%$ ConnectionStrings:IDMSConnectionString %>"
                        SelectCommand="SELECT DISTINCT [軟體來源] FROM [軟體主檔] order by [軟體來源]">
                    </asp:SqlDataSource>
                    <span onclick="TextSource.value=SelSource.value;"><font size="2" color="blue" style="cursor:pointer"><u>帶入</u></font></span> 
                </asp:TableCell>
                <asp:TableCell runat="server" Font-Size="Small">
                    <asp:Label ID="helpSource" runat="server" Font-Size="Small" ForeColor="#003300"></asp:Label>
                </asp:TableCell>
            </asp:TableRow>

            <asp:TableRow ID="rowFunctions" runat="server" HorizontalAlign="Left" VerticalAlign="Middle">
                <asp:TableCell runat="server" Font-Bold="True">軟體功能</asp:TableCell>
                <asp:TableCell runat="server" Font-Size="Small">
                    <asp:TextBox ID="TextFunctions" runat="server" ForeColor="Green"></asp:TextBox>
                    &nbsp;&nbsp;
                    <asp:DropDownList ID="SelFunctions" runat="server" AutoPostBack="True" OnSelectedIndexChanged="SelFunctions_SelectedIndexChanged"
                        DataSourceID="SqlDataSourceFunctions" DataTextField="軟體功能" DataValueField="軟體功能" ForeColor="#009933">
                    </asp:DropDownList>
                    <asp:SqlDataSource ID="SqlDataSourceFunctions" runat="server" ConnectionString="<%$ ConnectionStrings:IDMSConnectionString %>"
                        SelectCommand="SELECT DISTINCT [軟體功能] FROM [軟體主檔] order by [軟體功能]">
                    </asp:SqlDataSource>
                    <span onclick="TextFunctions.value=SelFunctions.value;"><font size="2" color="blue" style="cursor:pointer"><u>帶入</u></font></span>
                </asp:TableCell>
                <asp:TableCell runat="server" Font-Size="Small">
                    <asp:Label ID="helpFunctions" runat="server" Font-Size="Small" ForeColor="#003300"></asp:Label>
                </asp:TableCell>
            </asp:TableRow>

            <asp:TableRow ID="rowType" runat="server" HorizontalAlign="Left" VerticalAlign="Middle">
                <asp:TableCell runat="server" Font-Bold="True">機器型號</asp:TableCell>
                <asp:TableCell runat="server" Font-Size="Small">
                    適用廠牌：<asp:TextBox ID="TextBrand" runat="server" ForeColor="Green"></asp:TextBox>&nbsp;&nbsp;
                    <asp:DropDownList ID="SelBrand" runat="server" AutoPostBack="True" OnSelectedIndexChanged="SelBrand_SelectedIndexChanged"
                        ForeColor="#009933" DataSourceID="SqlDataSourceBrand" DataTextField="適用廠牌" DataValueField="適用廠牌">
                    </asp:DropDownList>
                    <asp:SqlDataSource ID="SqlDataSourceBrand" runat="server" ConnectionString="<%$ ConnectionStrings:IDMSConnectionString %>"
                        SelectCommand="SELECT DISTINCT [適用廠牌] FROM [軟體主檔] order by [適用廠牌]">
                    </asp:SqlDataSource>
                    <span onclick="TextBrand.value=SelBrand.value;"><font size="2" color="blue" style="cursor:pointer"><u>帶入</u></font></span>
                    <br />
                    適用型式：<asp:TextBox ID="TextStyle" runat="server" ForeColor="Green"></asp:TextBox>&nbsp;&nbsp;
                    <asp:DropDownList ID="SelStyle" runat="server" ForeColor="#009933" DataSourceID="SqlDataSourceStyle"
                        DataTextField="適用型式" DataValueField="適用型式" AutoPostBack="True" OnSelectedIndexChanged="SelStyle_SelectedIndexChanged">
                    </asp:DropDownList>
                    <asp:SqlDataSource ID="SqlDataSourceStyle" runat="server" ConnectionString="<%$ ConnectionStrings:IDMSConnectionString %>"
                        SelectCommand="SELECT DISTINCT [適用型式] FROM [軟體主檔] where [適用廠牌]=@適用廠牌 order by [適用型式]">
                        <SelectParameters>
                            <asp:ControlParameter ControlID="SelBrand" Name="適用廠牌" PropertyName="SelectedValue" Type="String" />
                        </SelectParameters>
                    </asp:SqlDataSource>
                </asp:TableCell>
                <asp:TableCell runat="server" Font-Size="Small">
                    <asp:Label ID="helpType" runat="server" Font-Size="Small" ForeColor="#003300"></asp:Label>
                </asp:TableCell>
            </asp:TableRow>

            <asp:TableRow ID="rowSN" runat="server" HorizontalAlign="Left" VerticalAlign="Middle">
                <asp:TableCell runat="server" Font-Bold="True">軟體序號</asp:TableCell>
                <asp:TableCell runat="server" Font-Size="Small">
                   購買：<asp:TextBox ID="TextSnBuy" runat="server" Columns="48" Width="500px" ForeColor="Green" ></asp:TextBox>
                   <br />
                   降級：<asp:TextBox ID="TextSnDown" runat="server" Columns="48" Width="500px" Height="100px" Rows="5" TextMode="MultiLine" ForeColor="Green"></asp:TextBox>
                </asp:TableCell>
                <asp:TableCell runat="server" Font-Size="Small">
                    <asp:Label ID="helpSN" runat="server" Font-Size="Small" ForeColor="#003300"></asp:Label>
                </asp:TableCell>
            </asp:TableRow>

            <asp:TableRow ID="rowLife" runat="server" HorizontalAlign="Left" VerticalAlign="Middle">
                <asp:TableCell runat="server" Font-Bold="True">軟體期限</asp:TableCell>
                <asp:TableCell runat="server" Font-Size="Small">
                    期限說明：<asp:TextBox ID="TextLife" runat="server"></asp:TextBox>
                    &nbsp;&nbsp;
                    <asp:DropDownList ID="SelLife" runat="server" AutoPostBack="True" OnSelectedIndexChanged="SelLife_SelectedIndexChanged"
                        DataSourceID="SqlDataSourceLife" DataTextField="期限說明" DataValueField="期限說明" ForeColor="#009933">
                    </asp:DropDownList>
                    <asp:SqlDataSource ID="SqlDataSourceLife" runat="server" ConnectionString="<%$ ConnectionStrings:IDMSConnectionString %>"
                        SelectCommand="SELECT DISTINCT [期限說明] FROM [軟體主檔] order by [期限說明]">
                    </asp:SqlDataSource>
                    <span onclick="TextLife.value=SelLife.value;"><font size="2" color="blue" style="cursor:pointer"><u>帶入</u></font></span>
                    <br />
                    購買日期：<asp:DropDownList ID="SelBuyYYYY" runat="server" ForeColor="#009933"></asp:DropDownList>年 &nbsp;&nbsp;
                    <asp:DropDownList ID="SelBuyMM" runat="server" ForeColor="#009933"></asp:DropDownList>月 &nbsp;&nbsp;
                    <asp:DropDownList ID="SelBuyDD" runat="server" ForeColor="#009933"></asp:DropDownList>日 &nbsp;&nbsp;
                    <br />
                    使用期限：<asp:DropDownList ID="SelUseYYYY" runat="server" ForeColor="#009933"></asp:DropDownList>年 &nbsp;&nbsp;
                    <asp:DropDownList ID="SelUseMM" runat="server" ForeColor="#009933"></asp:DropDownList>月 &nbsp;&nbsp;
                    <asp:DropDownList ID="SelUseDD" runat="server" ForeColor="#009933"></asp:DropDownList>日 &nbsp;&nbsp;
                    <br />
                    更新期限：<asp:DropDownList ID="SelUpYYYY" runat="server" ForeColor="#009933"></asp:DropDownList>年 &nbsp;&nbsp;
                    <asp:DropDownList ID="SelUpMM" runat="server" ForeColor="#009933"></asp:DropDownList>月 &nbsp;&nbsp;
                    <asp:DropDownList ID="SelUpDD" runat="server" ForeColor="#009933"></asp:DropDownList>日 &nbsp;&nbsp;
                </asp:TableCell>
                <asp:TableCell runat="server" Font-Size="Small">
                    <asp:Label ID="helpLife" runat="server" Font-Size="Small" ForeColor="#003300"></asp:Label>
                </asp:TableCell>
            </asp:TableRow>

            <asp:TableRow ID="rowPrice" runat="server" HorizontalAlign="Left" VerticalAlign="Middle">
                <asp:TableCell runat="server" Font-Bold="True">軟體價格</asp:TableCell>
                <asp:TableCell runat="server" Font-Size="Small">
                    單價：<asp:TextBox ID="TextPrice" Columns="10" MaxLength="10" runat="server" ForeColor="Green" OnKeyDown="if(event.keyCode <=57 | event.keyCode >=96 & event.keyCode <=105 | event.keyCode ==110 | event.keyCode ==190) {event.returnValue=true;} else{event.returnValue=false;}"></asp:TextBox>
                    &nbsp;&nbsp;
                    <asp:RangeValidator ID="RangeValidatorPrice" runat="server" ControlToValidate="TextPrice"
                        Display="Dynamic" ErrorMessage="單價欄位輸入有誤！必須輸入純數字，不要包含逗點。" Font-Size="Small" MaximumValue="10000000000"
                        MinimumValue="0" SetFocusOnError="True" Type="Double">
                    </asp:RangeValidator>
                    總價：<asp:TextBox ID="TextTotal" Columns="10" MaxLength="10" runat="server" ForeColor="Green" OnKeyDown="if(event.keyCode <=57 | event.keyCode >=96 & event.keyCode <=105 | event.keyCode ==110 | event.keyCode ==190) {event.returnValue=true;} else{event.returnValue=false;}"></asp:TextBox>
                    &nbsp;&nbsp;
                    <asp:RangeValidator ID="RangeValidatorTotal" runat="server" ControlToValidate="TextTotal"
                        Display="Dynamic" ErrorMessage="總價欄位輸入有誤！必須輸入純數字，不要包含逗點。" Font-Size="Small" MaximumValue="10000000000"
                        MinimumValue="0" SetFocusOnError="True" Type="Double">
                    </asp:RangeValidator>
                    說明：<asp:TextBox ID="TextPriceMemo" runat="server" ForeColor="Green"></asp:TextBox>
                </asp:TableCell>
                <asp:TableCell runat="server" Font-Size="Small">
                    <asp:Label ID="helpPrice" runat="server" Font-Size="Small" ForeColor="#003300"></asp:Label>
                </asp:TableCell>
            </asp:TableRow>

            <asp:TableRow ID="rowSupply" runat="server" HorizontalAlign="Left" VerticalAlign="Middle">
                <asp:TableCell runat="server" Font-Bold="True">提供單位</asp:TableCell>
                <asp:TableCell runat="server" Font-Size="Small">
                    <asp:TextBox ID="TextSupply" runat="server" Width="500px" ForeColor="Green" ></asp:TextBox>
                    <span onclick="TextSupply.value=SelSupply.value;"><font size="2" color="blue" style="cursor:pointer"><u>帶入</u></font></span>
                    <br />
                    <asp:DropDownList ID="SelSupply" runat="server" AutoPostBack="True" OnSelectedIndexChanged="SelSupply_SelectedIndexChanged"
                        DataSourceID="SqlDataSourceSupply" DataTextField="提供單位" DataValueField="提供單位" ForeColor="#009933">
                    </asp:DropDownList>
                    <asp:SqlDataSource ID="SqlDataSourceSupply" runat="server" ConnectionString="<%$ ConnectionStrings:IDMSConnectionString %>"
                        SelectCommand="SELECT DISTINCT [提供單位] FROM [軟體主檔] order by [提供單位]">
                    </asp:SqlDataSource>
                </asp:TableCell>
                <asp:TableCell runat="server" Font-Size="Small">
                    <asp:Label ID="helpSupply" runat="server" Font-Size="Small" ForeColor="#003300"></asp:Label>
                </asp:TableCell>
            </asp:TableRow>

            <asp:TableRow ID="rowMedia" runat="server" HorizontalAlign="Left" VerticalAlign="Middle">
                <asp:TableCell runat="server" Font-Bold="True">存放媒體</asp:TableCell>
                <asp:TableCell runat="server" Font-Size="Small">
                    <asp:TextBox ID="TextMedia" runat="server" Columns="48" Width="500px" Height="100px" Rows="5" ForeColor="Green" TextMode="MultiLine"></asp:TextBox>
                </asp:TableCell>
                <asp:TableCell runat="server" Font-Size="Small">
                    <asp:Label ID="helpMedia" runat="server" Font-Size="Small" ForeColor="#003300"></asp:Label>
                </asp:TableCell>
            </asp:TableRow>

            <asp:TableRow ID="rowBookNo" runat="server" HorizontalAlign="Left" VerticalAlign="Middle">
                <asp:TableCell runat="server" Font-Bold="True">圖書編號</asp:TableCell>
                <asp:TableCell runat="server" Font-Size="Small">
                    <asp:TextBox ID="TextBookNo" runat="server" Columns="48" Width="500px" ForeColor="Green"></asp:TextBox>
                </asp:TableCell>
                <asp:TableCell runat="server" Font-Size="Small">
                    <asp:Label ID="helpBookNo" runat="server" Font-Size="Small" ForeColor="#003300"></asp:Label>
                </asp:TableCell>
            </asp:TableRow>

            <asp:TableRow ID="rowAttach" runat="server" HorizontalAlign="Left" VerticalAlign="Middle">
                <asp:TableCell runat="server" Font-Bold="True">軟體附件</asp:TableCell>
                <asp:TableCell runat="server" Font-Size="Small">
                    <asp:TextBox ID="TextAttach" runat="server" Columns="48" Width="500px" Height="100px" Rows="5" ForeColor="Green" TextMode="MultiLine"></asp:TextBox>
                </asp:TableCell>
                <asp:TableCell runat="server" Font-Size="Small">
                    <asp:Label ID="helpAttach" runat="server" Font-Size="Small" ForeColor="#003300"></asp:Label>
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
                        SelectCommand="SELECT DISTINCT [減損原因] FROM [軟體主檔] order by [減損原因]">
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
                        SelectCommand="SELECT DISTINCT [減損方式] FROM [軟體主檔] order by [減損方式]">
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

            <asp:TableRow ID="rowStatus" runat="server" HorizontalAlign="Left">
                <asp:TableCell runat="server" Font-Bold="True">軟體狀態</asp:TableCell>
                <asp:TableCell runat="server" Font-Size="Small">
                    <asp:DropDownList ID="SelStatus" runat="server" ForeColor="#009933"
                        DataSourceID="SqlDataSourceStatus" DataTextField="Item" DataValueField="Item">
                    </asp:DropDownList>
                    <asp:SqlDataSource ID="SqlDataSourceStatus" runat="server" ConnectionString="<%$ ConnectionStrings:IDMSConnectionString %>"
                        SelectCommand="SELECT [Item] FROM [Config] WHERE [Kind] = '軟體狀態' ORDER BY [mark]">
                    </asp:SqlDataSource><br /><br />
                    <asp:Label ID="lblStatus" runat="server" Font-Size="Small"></asp:Label>
                </asp:TableCell>
                <asp:TableCell ID="TableCell6" runat="server" Font-Size="Small">
                    <asp:Label ID="helpStatus" runat="server" Font-Size="Small" ForeColor="#003300"></asp:Label>
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
                    <asp:Label ID="Label2" runat="server" ForeColor="red" Text="" Font-Bold="True"></asp:Label>
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
        <asp:Button ID="BtnPre"  runat="server" Text="　預　覽　" OnClick="BtnPre_Click" CausesValidation="False" />　　
        <asp:Button ID="BtnExit" runat="server" Text="　關　閉　" OnClientClick="window.close();" CausesValidation="False" />
        <br />
        <asp:Label runat="server" Font-Size="Small" ForeColor="#006600" Text="(★：必填欄位　　紅色：連結欄位)"></asp:Label>
    </div>    
    </form>
    <asp:Label ID="lblChecks" runat="server" ForeColor="red" Text="" Font-Bold="True"></asp:Label>
    <br />
    <asp:Label ID="Label1" runat="server" Font-Size="Small" ForeColor="#003300"></asp:Label>
</body>
</html>
