<%@ Page Language="C#" Title="設備上下游" MasterPageFile="~/IDMS/MasterPage.master" AutoEventWireup="true" CodeFile="DeviceTree.aspx.cs" Inherits="Device_test" MaintainScrollPositionOnPostback="true" Debug="true" %>


    <asp:Content ID="Content1" ContentPlaceHolderID="head" runat="Server">

        <style type="text/css">
            .div_ups {
                margin: 20px 0 0 0;
                display: flex;
                justify-content: space-between;
            }
            
            .div_tree {
                background-color: rgba(169, 204, 201, 0.373);
                margin: 5px 0px 5px 0px;
                min-width: 1250px;
                min-height: 720px;
            }
            
            .mitch-d3-tree {
                max-width: 1250px;
                max-height: 720px;
            }
            
            .DevTABLE {
                margin-left: 20px;
            }
            
            .funp {
                overflow: hidden;
                margin-bottom: 0px;
                text-overflow: ellipsis;
                display: -webkit-box;
                -webkit-line-clamp: 2;
                -webkit-box-orient: vertical;
            }
            
            .psdiv {
                color: #CC6600;
                line-height: 2;
                font-size: small;
            }
            
            .clicked_but {
                background-color: rgb(177, 177, 177);
            }
            
            tr,
            td {
                text-align: center;
            }
        </style>


        <meta http-equiv="X-UA-Compatible" content="IE=edge">
        <script src="assets/js/d3-mitch-tree.min.js"></script>
        <script src="assets/js/jquery-3.6.0.min.js"></script>
        <script>
            function save_svg() {
                var e = document.createElement('script');
                e.setAttribute('src', 'assets/js/svg-crowbar.js');
                e.setAttribute('class', 'svg-crowbar');
                document.body.appendChild(e);
                console.log("DONE");
            }

            function Device_edit() {
                var devID = document.getElementsByName('DevID')[0].innerText;
                devID = devID.split('-')[0];
                if (devID) {
                    var url = "DevEdit.aspx?DevNo=" + devID;
                    window.open(url, '_blank');
                } else {
                    alert('請先點擊設備');
                }
            }

            function button_click() {
                console.log("CLICK");
            }
            var cbo1 = false;
            var cbo2 = false;
            $(document).ready(function() {
                $('#cbox1').click(function() {
                    cbo1 = !cbo1;
                    if (cbo1) {
                        $('#cbox1').addClass("clicked_but");
                        treePlugin.setAllowFocus(true);

                    } else {
                        $('#cbox1').removeClass("clicked_but");
                        treePlugin.setAllowFocus(false);

                    }
                })
                $('#cbox2').click(function() {
                    cbo2 = !cbo2;

                    if (cbo2) {
                        $('#cbox2').addClass("clicked_but");
                        var nodes = treePlugin.getNodes();
                        nodes.forEach(function(node, index, arr) {
                            treePlugin.expand(node);
                        });
                        treePlugin.update(treePlugin.getRoot());
                        console.log(true);
                    } else {
                        $('#cbox2').removeClass("clicked_but");
                        var nodes = treePlugin.getNodes();
                        nodes.forEach(function(node, index, arr) {
                            treePlugin.collapse(node);
                        });
                        treePlugin.update(treePlugin.getRoot());
                        console.log(false);
                    }
                })



            })
        </script>

        <link rel="stylesheet" type="text/css" href="assets/css/d3-mitch-tree.min.css">
        <link rel="stylesheet" type="text/css" href="assets/css/d3-mitch-tree-theme-default.min.css">
    </asp:Content>

    <asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1" runat="Server">
        <div style="display:flex;align-items: center;">

            <div>

                <div class="div_ups">
                    <div>
                        選擇UPS:
                        <asp:DropDownList id="UPSList" AutoPostBack="True" OnSelectedIndexChanged="Selection_Change" runat="server">
                            <asp:ListItem Selected="True" Value="UPS 825"> UPS 825(A) </asp:ListItem>
                            <asp:ListItem Value="UPS 550(B)"> UPS 550(B) </asp:ListItem>
                            <asp:ListItem Value="UPS 550(C)"> UPS 550(C) </asp:ListItem>
                            <asp:ListItem Value="UPS 550(D)"> UPS 550(D) </asp:ListItem>
                            <asp:ListItem Value="UPS 60"> UPS 60 </asp:ListItem>
                            <asp:ListItem Value="UPS 60(O+)"> UPS 60(O+) </asp:ListItem>
                            <asp:ListItem Value="UPS 160"> UPS 160 </asp:ListItem>
                            <asp:ListItem Value="UPS 175"> UPS 175 </asp:ListItem>
                            <asp:ListItem Value="市電"> 市電 </asp:ListItem>
                            <asp:ListItem Value="發電機"> 發電機 </asp:ListItem>
                            <asp:ListItem Value="總電源"> 總電源 </asp:ListItem>
                        </asp:DropDownList>
                    </div>



                    <div>
                        <input type="button" id="cbox1" value="單一節點">
                        <input type="button" id="cbox2" value="展開節點">
                        <input type="text" id="focusText" placeholder="輸入欲找尋設備" />
                        <input type="button" id="focusButton" value="搜尋節點">
                    </div>


                </div>
                <div class="div_tree">
                    <span>*滾輪縮放、點擊展開</span>
                    <section id="my_tree3"></section>
                    <section id="my_tree2"></section>
                </div>

            </div>


            <asp:Table class="DevTABLE" ID="Table1" runat="server" width="320px" height="40px" GridLines="Both">
                <asp:TableRow ID="row1" runat="server">
                    <asp:TableCell runat="server" Width="100px" height="40px">設備編號</asp:TableCell>
                    <asp:TableCell runat="server">
                        <asp:Label id="DevID" name="DevID" runat="server"></asp:Label>
                    </asp:TableCell>
                </asp:TableRow>

                <asp:TableRow ID="row2" runat="server">
                    <asp:TableCell runat="server" Width="100px" height="40px">設備名稱</asp:TableCell>
                    <asp:TableCell runat="server">
                        <label id="DevNAME" class="funp"></label>
                    </asp:TableCell>
                </asp:TableRow>

                <asp:TableRow ID="row3" runat="server">
                    <asp:TableCell runat="server" Width="100px" height="40px">設備地點</asp:TableCell>
                    <asp:TableCell runat="server">
                        <label id="DevLoc" class="funp"></label>
                    </asp:TableCell>
                </asp:TableRow>

                <asp:TableRow ID="row11" runat="server">
                    <asp:TableCell ColumnSpan="2" runat="server" Width="100px">
                        <input type="button" value="編輯" onclick="Device_edit()">
                    </asp:TableCell>

                </asp:TableRow>
            </asp:Table>

        </div>


        <input type="button" value="儲存" onclick="save_svg()">
        <span class="psdiv">*調整至欲截圖畫面及適當大小後儲存，可再列印成PDF格式</span>
        <p onClick="save_svg2()"></p>
        <label id="link"></label>
        <asp:Label id="l3" runat="server"></asp:Label>


        <script>
            var s = '<%=TestF() %>';

            var tree3dd = '<%=Get_Json3() %>]';
            //var tree2 = '<%=Get_Json2() %>]';
            tree3dd = JSON.parse(tree3dd);


            var truedata = [];
            var tree = {};

            //var devinfo = 'Get_Devinfo() %>]';
            //devinfo = JSON.parse(devinfo);

            var treePlugin = new d3.mitchTree.boxedTree()
                .setIsFlatData(true)
                .setData(tree3dd)

            .setAllowNodeCentering(false)
                .setAllowFocus(false)
                .setElement(document.getElementById("my_tree3"))
                .setMinScale(0.3)
                .setIdAccessor(function(data) {
                    return data.id;
                })
                .setParentIdAccessor(function(data) {
                    return data.parentId;
                })
                .setBodyDisplayTextAccessor(function(data) {
                    return data.description;
                })
                .setTitleDisplayTextAccessor(function(data) {
                    return data.type;
                })
                .getNodeSettings()
                .setSizingMode('nodesize')
                .setVerticalSpacing(30)
                .setHorizontalSpacing(2)
                .back()
                .on("nodeClick", function(event) { //node click event
                    //document.getElementsByName('DevID')[0].innerHTML = '<a href="DevEdit.aspx?DevNo=' + event.data.id + '" target="_blank">' + event.data.id;
                    document.getElementsByName('DevID')[0].innerHTML = event.data.id;
                    document.getElementById('DevNAME').innerHTML = event.data.description;
                    document.getElementById('DevLoc').innerHTML = event.data.loca;


                })
                .initialize();

            document.getElementById('focusButton').addEventListener('click', function() {
                var value = document.getElementById('focusText').value;
                if (value == "") {
                    alert("欄位空白");
                    return null;
                }
                console.log(value);
                var nodeMatchingText = (treePlugin.getNodes()).find(function(node) {
                    return node.data.description == value;
                });

                try {
                    if (nodeMatchingText)
                        treePlugin.focusToNode(nodeMatchingText);
                    else if (value != null)
                        treePlugin.focusToNode(value);
                } catch (error) {
                    var re = "找不到設備 : " + value;
                    alert(re);
                }

            });

            treePlugin.getZoomListener().scaleTo(treePlugin.getSvg(), 0.8);
            treePlugin.getZoomListener().translateTo(treePlugin.getSvg(), treePlugin.getWidthWithoutMargins(), treePlugin.getHeightWithoutMargins());
        </script>

    </asp:Content>