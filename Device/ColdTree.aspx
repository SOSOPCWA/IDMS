<%@ Page Language="C#" Title="群組關聯圖" 
MasterPageFile="~/IDMS/MasterPage.master" AutoEventWireup="true" 
CodeFile="ColdTree.aspx.cs" Inherits="Device_test" 
MaintainScrollPositionOnPostback="true" Debug="true" %>



    <asp:Content ID="Content1" ContentPlaceHolderID="head" runat="Server">


        <style>
            #myDiagramDiv {
                background-color: rgba(169, 204, 201, 0.373);
            }
            
            .headtext>span {
                margin: 12px 16px 0 0;
                font-size: 20px;
                font-weight: 800;
            }
            
            .headtext>a {
                font-family: Helvetica, sans-serif;
            }
            
            .content {
                margin-top: 20px;
                z-index: -1;
            }
            
            .mybtn {
                color: #fff;
                border: 1px solid transparent;
                border-radius: .25rem;
                background-color: #6c757d;
                border-color: #6c757d;
                line-height: 1.5;
                cursor: pointer;
            }
        </style>
    </asp:Content>



    <asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1" runat="Server">

        <script src="./assets/js/go.js"></script>
        <script src="./assets/js/jquery-3.6.0.min.js"></script>


        <div id="allSampleContent" class="content">
            <script id="code">
                function init() {

                    // Since 2.2 you can also author concise templates with method chaining instead of GraphObject.make
                    // For details, see https://gojs.net/latest/intro/buildingObjects.html
                    const $ = go.GraphObject.make; // for conciseness in defining templates

                    myDiagram =
                        $(go.Diagram, "myDiagramDiv", // must name or refer to the DIV HTML element
                            {
                                "toolManager.mouseWheelBehavior": go.ToolManager.WheelScroll,
                                allowCopy: false,
                                allowDelete: false,
                                draggingTool: $(CustomDraggingTool),
                                layout: $(CustomLayout),
                                //layout: $(go.LayeredDigraphLayout),
                                autoScale: go.Diagram.initialAutoScale,
                                // enable undo & redo
                                hasHorizontalScrollbar: false,
                                hasVerticalScrollbar: false,
                                "undoManager.isEnabled": true,

                            });

                    //圖的點擊行為
                    myDiagram.addDiagramListener("ObjectSingleClicked", function(e) {
                        var part = e.subject.part;
                        if (!(part instanceof go.Link)) {
                            console.log("Clicked on " + part.data.key + part.data.text);
                            jQuery("#devid").attr({
                                    "href": `DevEdit.aspx?DevNo=${part.data.key}`
                                }, ) //Device/DevEdit.aspx?DevNo=2596
                            jQuery("#devid").attr({
                                    "target": `_blank`
                                }, ) //open in another window
                            jQuery("#devid").text(part.data.key)
                            jQuery("#devname").text(part.data.text)
                        } else {
                            console.log("LINE: " + part.data.from + " --> " + part.data.to);
                        }

                    });
                    // 節點屬性
                    myDiagram.nodeTemplate =
                        $(go.Node, "Auto",
                            new go.Binding("location", "loc", go.Point.parse).makeTwoWay(go.Point.stringify),
                            // define the node's outer shape, which will surround the TextBlock
                            $(go.Shape, "RoundedRectangle", {
                                fill: "rgb(254, 201, 0)",
                                stroke: "black",
                                parameter1: 20, // the corner has a large radius
                                portId: "",
                                fromSpot: go.Spot.AllSides,
                                toSpot: go.Spot.AllSides
                            }),
                            $(go.TextBlock,
                                new go.Binding("text", "text").makeTwoWay())
                        );
                    //群組屬性
                    myDiagram.nodeTemplateMap.add("Super",
                        $(go.Node, "Auto", {
                                locationObjectName: "BODY"
                            },
                            $(go.Shape, "RoundedRectangle", {
                                fill: "rgba(128, 128, 64, 0.5)",
                                strokeWidth: 1.5,
                                parameter1: 20,
                                spot1: go.Spot.TopLeft,
                                spot2: go.Spot.BottomRight,
                                minSize: new go.Size(30, 30)
                            }),
                            $(go.Panel, "Vertical", {
                                    margin: 10
                                },
                                $(go.TextBlock, {
                                        font: "bold 14pt sans-serif",
                                        margin: new go.Margin(0, 0, 5, 0)
                                    },
                                    new go.Binding("text")),
                                $(go.Shape, {
                                    name: "BODY",
                                    opacity: 0
                                })
                            )
                        ));

                    // 連線屬性
                    myDiagram.linkTemplate =
                        $(go.Link, // the whole link panel
                            {
                                routing: go.Link.Orthogonal,
                                corner: 10
                            },
                            $(go.Shape, // the link shape
                                {
                                    strokeWidth: 1.6
                                }),
                            $(go.Shape, // the arrowhead
                                {
                                    toArrow: "Standard",
                                    fromArrow: "Standard",
                                    stroke: null
                                })
                        );

                    // read in the JSON-format data from the "mySavedModel" element
                    load();
                }

                // A custom layout that sizes each "Super" node to be big enough to cover all of it member nodes
                class CustomLayout extends go.Layout {
                        doLayout(coll) {
                            coll = this.collectParts(coll);

                            const supers = new go.Set( /*go.Node*/ );
                            coll.each(p => {
                                if (p instanceof go.Node && p.category === "Super") supers.add(p);
                            });

                            function membersOf(sup, diag) {
                                const set = new go.Set( /*go.Part*/ );
                                const arr = sup.data._members;
                                for (let i = 0; i < arr.length; i++) {
                                    const d = arr[i];
                                    set.add(diag.findNodeForData(d));
                                }
                                return set;
                            }

                            function isReady(sup, supers, diag) {
                                const arr = sup.data._members;

                                for (let i = 0; i < arr.length; i++) {
                                    const d = arr[i];
                                    if (d.category !== "Super") continue;
                                    const n = diag.findNodeForData(d);
                                    if (supers.has(n)) return false;
                                }
                                return true;
                            }

                            // implementations of doLayout that do not make use of a LayoutNetwork
                            // need to perform their own transactions
                            this.diagram.startTransaction("Custom Layout");
                            //console.log(supers.iterator)
                            while (supers.count > 0) {
                                let ready = null;
                                const it = supers.iterator;
                                while (it.next()) {
                                    if (isReady(it.value, supers, this.diagram)) {
                                        ready = it.value;
                                        break;
                                    }
                                }

                                supers.remove(ready);
                                const b = this.diagram.computePartsBounds(membersOf(ready, this.diagram));
                                ready.location = b.position;
                                const body = ready.findObject("BODY");

                                if (body) body.desiredSize = b.size;

                            }
                            this.diagram.commitTransaction("Custom Layout");

                        }
                    }
                    // end CustomLayout
                function layout() {
                    myDiagram.startTransaction("change Layout");
                    var lay = myDiagram.layout;

                    var direction = getRadioValue("direction");
                    direction = parseFloat(direction, 10);
                    lay.direction = direction;

                    var layerSpacing = document.getElementById("layerSpacing").value;
                    layerSpacing = parseFloat(layerSpacing, 10);
                    lay.layerSpacing = layerSpacing;

                    var columnSpacing = document.getElementById("columnSpacing").value;
                    columnSpacing = parseFloat(columnSpacing, 10);
                    lay.columnSpacing = columnSpacing;

                    var cycleRemove = getRadioValue("cycleRemove");
                    if (cycleRemove === "CycleDepthFirst") lay.cycleRemoveOption = go.LayeredDigraphLayout.CycleDepthFirst;
                    else if (cycleRemove === "CycleGreedy") lay.cycleRemoveOption = go.LayeredDigraphLayout.CycleGreedy;

                    var layering = getRadioValue("layering");
                    if (layering === "LayerOptimalLinkLength") lay.layeringOption = go.LayeredDigraphLayout.LayerOptimalLinkLength;
                    else if (layering === "LayerLongestPathSource") lay.layeringOption = go.LayeredDigraphLayout.LayerLongestPathSource;
                    else if (layering === "LayerLongestPathSink") lay.layeringOption = go.LayeredDigraphLayout.LayerLongestPathSink;

                    var initialize = getRadioValue("initialize");
                    if (initialize === "InitDepthFirstOut") lay.initializeOption = go.LayeredDigraphLayout.InitDepthFirstOut;
                    else if (initialize === "InitDepthFirstIn") lay.initializeOption = go.LayeredDigraphLayout.InitDepthFirstIn;
                    else if (initialize === "InitNaive") lay.initializeOption = go.LayeredDigraphLayout.InitNaive;

                    var aggressive = getRadioValue("aggressive");
                    if (aggressive === "AggressiveLess") lay.aggressiveOption = go.LayeredDigraphLayout.AggressiveLess;
                    else if (aggressive === "AggressiveNone") lay.aggressiveOption = go.LayeredDigraphLayout.AggressiveNone;
                    else if (aggressive === "AggressiveMore") lay.aggressiveOption = go.LayeredDigraphLayout.AggressiveMore;

                    //TODO implement pack option
                    var pack = document.getElementsByName("pack");
                    var packing = 0;
                    for (var i = 0; i < pack.length; i++) {
                        if (pack[i].checked) packing = packing | parseInt(pack[i].value, 10);
                    }
                    lay.packOption = packing;

                    var setsPortSpots = document.getElementById("setsPortSpots");
                    lay.setsPortSpots = setsPortSpots.checked;

                    myDiagram.commitTransaction("change Layout");
                }

                // Define a custom DraggingTool
                class CustomDraggingTool extends go.DraggingTool {
                        moveParts(parts, offset, check) {
                            super.moveParts(parts, offset, check);

                            this.diagram.layoutDiagram(true);

                        }

                        computeEffectiveCollection(parts) {
                            const coll = new go.Set( /*go.Part*/ ).addAll(parts);

                            parts.each(p => this.walkSubTree(p, coll));
                            return super.computeEffectiveCollection(coll);
                        };

                        // Find other attached nodes.
                        walkSubTree(sup, coll) {
                            if (sup === null) return;

                            coll.add(sup);
                            if (sup.category !== "Super") return;
                            // recurse through this super state's members
                            const model = this.diagram.model;
                            const mems = sup.data._members;
                            //console.log(this.diagram)
                            if (mems) {
                                for (let i = 0; i < mems.length; i++) {
                                    const mdata = mems[i];
                                    this.walkSubTree(this.diagram.findNodeForData(mdata), coll);
                                }
                            }
                        }
                    }
                    // end CustomDraggingTool class

                // Show the diagram's model in JSON format
                function save() {
                    document.getElementById("mySavedModel").value = myDiagram.model.toJson();
                    myDiagram.isModified = false;
                }


                function load() {
                    myDiagram.model = go.Model.fromJson(data);

                    // make sure all data have up-to-date "members" collections
                    // this does not prevent any "cycles" of membership, which would result in undefined behavior
                    const arr = myDiagram.model.nodeDataArray;
                    for (let i = 0; i < arr.length; i++) {
                        const data = arr[i];
                        const supers = data.supers;
                        if (supers) {
                            for (let j = 0; j < supers.length; j++) {
                                const sdata = myDiagram.model.findNodeDataForKey(supers[j]);
                                if (sdata) {
                                    // update _supers to be an Array of references to node data
                                    if (!data._supers) {

                                        data._supers = [sdata];
                                    } else {

                                        data._supers.push(sdata);
                                    }
                                    // update _members to be an Array of references to node data
                                    if (!sdata._members) {

                                        sdata._members = [data];
                                    } else {
                                        sdata._members.push(data);
                                    }
                                }
                            }
                        }
                    }
                }

                function test_button() { //控制位置函數
                    console.log("WORKING");
                    myDiagram.model.commit(function(m) {
                        var data = m.nodeDataArray[0]; // get the first node data
                        console.log(data)
                        m.set(data, "loc", "500 500");
                    }, "flash");
                }


                window.addEventListener('DOMContentLoaded', init);

                function callBackend() {
                    var xhr = new XMLHttpRequest();
                    xhr.open("POST", "ColdTree.aspx/GetData", true);
                    xhr.setRequestHeader("Content-Type", "application/json; charset=utf-8");
                    xhr.onreadystatechange = function() {
                        if (xhr.readyState === 4 && xhr.status === 200) {
                            var response = xhr.responseText;
                            // Process the response from the server here
                        }
                    };

                    var data = JSON.stringify({
                        parameter1: "value1",
                        parameter2: "value2"
                    });
                    xhr.send(data);
                    console.log(xhr)
                }
                var data1 = '<%=Get_Json() %>]}';
                var data = {}
                data = JSON.parse(data1);
            </script>

            <div id="sample">
                <div>
                    選擇圖形:
                    <asp:DropDownList id="GroupList" AutoPostBack="True" OnSelectedIndexChanged="Selection_Change" runat="server">
                        <asp:ListItem Value="0" selected="true"> 總空調 </asp:ListItem>
                    </asp:DropDownList>

                    <asp:button runat="server" text="群組編輯" onclick="Linkbtn_click" class="mybtn"></asp:button>
                </div>
                <!--go.js a.preventDefault()-->
                <div id="myDiagramDiv" style="margin-top: 1rem; border: 1px solid black; width: 100%; height: 700px; position: relative; -webkit-tap-highlight-color: rgba(255, 255, 255, 0);">
                    <canvas tabindex="0" width="1296" height="497" style="position: absolute; top: 0px; left: 0px; z-index: -1; user-select: none; touch-action: none; width: 1037px; height: 398px;">This text is displayed if your browser does not support the Canvas HTML element.</canvas>
                    <div style="position: absolute; overflow: auto; width: 1054px; height: 398px; z-index: -2!important ;">
                        <div style="position: absolute; width: 1px; height: 403.57px;"></div>
                    </div>
                </div>


                <div id="context">
                    <div class="headtext">
                        <span>設備編號:</span><a id="devid">節點編號</a>
                    </div>
                    <div class="headtext">
                        <span>設備名稱:</span><a id="devname">節點名稱</a>
                    </div>
                </div>
                <br>
                <!--<asp:Label runat="server" ID="Testa" Text="d"></asp:Label>--->

            </div>

            <asp:ScriptManager ID="ScriptManager1" runat="server" EnablePageMethods="true" />

        </div>

    </asp:Content>