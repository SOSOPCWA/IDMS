<%@ Page Language="C#" Title="群組編輯" MasterPageFile="~/IDMS/MasterPage.master" AutoEventWireup="true"
    CodeFile="Group.aspx.cs" Inherits="GroupEdit" MaintainScrollPositionOnPostback="true" Debug="true" EnableEventValidation="false" %>


    <asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server">

        <head>
            <title>設備編輯</title>
            <style type="text/css">
                .div1 {
                    display: inline-block;
                    width: 300px;
                    margin: 20px 30px 0 10px;
                }
                
                .div2 {
                    width: 70%;
                    margin: 20px 30px 0 10px;
                }
                
                .divbut {
                    margin-left: 2.5rem;
                    margin-top: 1rem;
                }
                
                ul li {
                    color: black;
                    margin: 10px 0 5px 1.5rem;
                }
                
                .main {
                    display: flex;
                    justify-content: space-around;
                    margin-bottom: 30px;
                }
                
                .mytbl {
                    font-size: 18px;
                    color: black!important;
                    vertical-align: bottom!important;
                }
                
                .trname {
                    overflow: hidden;
                    margin-bottom: 0px;
                    text-overflow: ellipsis;
                    display: -webkit-box;
                    -webkit-line-clamp: 2;
                    -webkit-box-orient: vertical;
                }
                
                .putbox {
                    height: 38px;
                    padding: 6px 10px;
                    /* The 6px vertically centers text on FF, ignored by Webkit */
                    background-color: #fff;
                    border: 1px solid #D1D1D1;
                    border-radius: 4px;
                    box-shadow: none;
                    box-sizing: border-box;
                    margin-bottom: 10px;
                }
                
                .input-dev {
                    min-height: 40px;
                }
                
                #TreeDev:hover {
                    font: bold;
                    color: black;
                }
            </style>




            <script src="assets/js/jquery-3.6.0.min.js"></script>
            <script type="text/javascript" src="assets/js/vendor/modernizr.custom.js"></script>
            <script type="text/javascript" src="assets/js/vendor/jquery.hideseek.min.js"></script>
            <script type="text/javascript" src="assets/js/vendor/rainbow-custom.min.js"></script>
            <script type="text/javascript" src="assets/js/vendor/jquery.anchor.js"></script>
            <script type="text/javascript" src="assets/js/bootstrap.bundle.min.js"></script>
            <script src="assets/js/vendor/initializers.js"></script>

            <link rel="stylesheet" type="text/css" href="assets/css/bootstrap.min.css">
            <link rel="stylesheet" type="text/css" href="assets/css/vendor/normalize.css">

            <link rel="stylesheet" type="text/css" href="assets/css/vendor/search.css">
            <link rel="stylesheet" type="text/css" href="assets/css/vendor/github.css">


            <script>
                var data = '<%=Get_list() %>]';

                var b = JSON.parse(data);

                $(function() {
                    $("[data-toggle='tooltip']").tooltip();
                });

                $(document).ready(function() {
                    b.forEach(element => {
                        $("#input-ul").append(
                            '<li class="Listname" style="display: none;"><a class="ske-a">(' + element.ID + ') -- ' + element.Name + '</a></li>'
                        );
                    });

                    $(".btnedit").click(function() {
                        //return confirm('Sure Delete?');
                    })

                    $(".btndelete").click(function() {
                        var tar_tr = $(this).parent().parent().children();

                        return confirm(`確定刪除這筆群組關聯資料嗎?\n <${tar_tr[0].innerText}>  ${tar_tr[1].innerText}`);
                    })

                    $(".Listname").click(function() { //搜尋列輸入
                        $("#search-hidden-mode").val($(this).text());
                        console.log($(this).text());
                    })

                    var Gid = $(".putbox").val().slice(0, 5); //群組編號
                    $("#TreeDev").attr("href", "/IDMS/Device/TreeEdit.aspx?DevNo=" + Gid);
                    $("#TreeDev").attr("target", "_blank");
                })

                function InputCheck(mes) {
                    if ($(".putbox").val().slice(0, 5).length < 5) {
                        alert("尚未指定群組，請先新增或選擇群組");
                        return false;
                    } else {
                        return confirm(mes);
                    }

                }

                function btndelete_click(id) {
                    var row = $("#del" + id).parent().parent().children()[0].innerText;
                    $('span[id$="TEST"]').text(row);
                    $('span[id$="TEST"]').val(row);
                    console.log($('span[id$="TEST"]').text())
                }
            </script>

        </head>

        <body style="background-color:LightGray;;">



            <script>
            </script>

            <div class="main">



                <div class="div2 my-3">
                    <div class="d-flex align-items-center">
                        <div>
                            <label>選擇群組</label><br />
                            <asp:DropDownList CssClass="putbox" ID="GroupList" runat="server" ForeColor="#1E8449" AutoPostBack="True" OnSelectedIndexChanged="GroupList_SelectedIndexChanged" />
                            </asp:DropDownList><br />
                        </div>
                        <div class="divbut">
                            <asp:button runat="server" text="新增群組" onclientclick="return false;" class="btn btn-secondary " id="groupmodal" data-bs-toggle="offcanvas" href="#offcanvasExample" role="button" data-target="#offcanvasExample" data-toggle="tooltip" data-placement="left"
                                title="檢視、編輯群組資料" aria-controls="offcanvasExample" />

                            <div class="offcanvas offcanvas-start" tabindex="-1" id="offcanvasExample" aria-labelledby="offcanvasExampleLabel">
                                <div class="offcanvas-header">
                                    <h5 class="offcanvas-title" id="offcanvasExampleLabel">群組編輯頁面</h5>
                                    <button type="button" class="btn-close text-reset" data-bs-dismiss="offcanvas" aria-label="Close"></button>
                                </div>

                                <div class="offcanvas-body">
                                    <div class="div-mod mt-2">
                                        <label>群組名稱</label>
                                        <asp:TextBox runat="server" CssClass="putbox" id="add_gname" placeholder="空調系統" class="w-50"></asp:TextBox>
                                        <label>群組用途</label><br>
                                        <asp:DropDownList runat="server" CssClass="putbox" id="add_guse">
                                            <asp:ListItem Selected="true" value="空調">空調</asp:ListItem>
                                            <asp:ListItem value="電力">電力</asp:ListItem>
                                            <asp:ListItem value="消防">消防</asp:ListItem>
                                            <asp:ListItem value="其他">其他</asp:ListItem>
                                        </asp:DropDownList><br>
                                        <label>群組說明</label>
                                        <asp:TextBox runat="server" CssClass="putbox" id="add_gps" placeholder="使用方式、備援形式"></asp:TextBox>
                                        <label>擺放地點</label>
                                        <asp:TextBox runat="server" CssClass="putbox" id="add_gloc"></asp:TextBox>
                                    </div>
                                </div>

                                <div class="modal-footer">
                                    <asp:button runat="server" class="button-primary btn btn-primary" Text="新增" onClientclick="return confirm('確定新增此群組嗎?\n')" onclick="BtnAddGroup_Click" />
                                    <asp:button runat="server" id="Btn_gedit" class="btn btn-warning" visible="false" Text="修改" onClientclick="return confirm('確定修改此群組嗎?\n')" onclick="BtnEditGroup_Click" />
                                    <asp:button runat="server" id="Btn_gdelete" class="btn btn-danger" visible="false" Text="刪除" onClientclick="return confirm('確定刪除此群組嗎?\n\n注意!!  此動作會刪除此群組及底下的群組關聯\n若群組下還有設備，請先移至其他群組')" onclick="BtnDeleteGroup_Click" />
                                </div>
                            </div>

                            <asp:button runat="server" visible="false" text="上下游設備" onclientclick="return false;" class="btn btn-secondary " id="linkmodal" data-bs-toggle="offcanvas" href="#offcanvasExample2" role="button" data-target="#offcanvasExample2" aria-controls="offcanvasExample2"
                                data-toggle="tooltip" data-bs-placement="right" title="檢視、編輯上下游設備" />

                            <div class="offcanvas offcanvas-end" tabindex="-1" id="offcanvasExample2" aria-labelledby="offcanvasExampleLabel">
                                <div class="offcanvas-header">
                                    <h5 class="offcanvas-title" id="offcanvasExampleLabel">
                                        <h3>上下游設備</h3>
                                        <a class="btn btn-warning btn-sm" ID="TreeDev">編輯</a>
                                    </h5>
                                    <button type="button" class="btn-close text-reset" data-bs-dismiss="offcanvas" aria-label="Close"></button>
                                </div>

                                <div class="offcanvas-body">
                                    <div class="div-mod mt-2">
                                        <label style="color: rgb(28, 28, 199);">上游設備</label><br>
                                        <asp:DropDownList id="toplink" runat="server" onclientclick="return false;" size="13" class="w-100">

                                        </asp:DropDownList>

                                    </div>

                                    <div class="div-mod mt-2">
                                        <label style="color: rgb(28, 28, 199);">下游設備</label><br>
                                        <asp:DropDownList id="downlink" runat="server" onclientclick="return false;" size="13" class="w-100">

                                        </asp:DropDownList>

                                    </div>
                                </div>


                            </div>



                        </div>
                    </div>
                    <!-- Modal -->


                    <!-- Button trigger modal -->


                    <asp:GridView ID="GridView1" runat="server" AllowSorting="True" AutoGenerateColumns="False" BorderStyle="None" BorderWidth="1px" CellPadding="3" DataSourceID="SqlDataSource1" GridLines="Vertical" class="table table-hover mytbl" onrowcommand="Row_Command">

                        <Columns>
                            <asp:BoundField DataField="ID" HeaderText="設備編號"></asp:BoundField>
                            <asp:BoundField DataField="DName" HeaderText="設備名稱"></asp:BoundField>
                            <asp:BoundField DataField="Loc" HeaderText="放置地點"></asp:BoundField>
                            <asp:ButtonField Text="編輯" Commandname="Edit1" ControlStyle-CssClass="btn btn-success me-3 btnedit" />
                        </Columns>
                        <Columns>
                            <asp:ButtonField Text="刪除" Commandname="Delete1" ControlStyle-CssClass="btn btn-danger me-3 btndelete" />



                        </Columns>

                    </asp:GridView>
                    <asp:SqlDataSource ID="SqlDataSource1" runat="server" ConnectionString="<%$ ConnectionStrings:IDMSConnectionString %>" SelectCommand="" ProviderName="<%$ ConnectionStrings:IDMSConnectionString.ProviderName %>">
                    </asp:SqlDataSource>


                </div>

                <div class="div1">
                    <article>
                        <label for="search">新增設備至群組</label>

                        <div class="mt-1">
                            <input id="search-hidden-mode" data-list=".hidden_mode_list" data-nodata="No results found" autocomplete="off" name="search" class="form-control input-dev rounded-2" type="text" placeholder="請輸入設備名稱或設備編號" aria-label=".form-control-lg example">
                            <asp:Button runat="server" ID="btnadd" class="btn  btn-outline-primary" OnClientClick="return InputCheck('確定新增嗎?')" onclick="BtnAdd_Click" Text="新增" data-toggle="tooltip" data-bs-placement="left" title="請先搜尋設備，點擊設備全名再按新增" />
                            <ul id="input-ul" class="vertical hidden_mode_list"></ul>
                        </div>

                    </article>
                </div>

            </div>




        </body>
    </asp:Content>