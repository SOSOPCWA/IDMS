<%@ Page Language="C#" Title="機房平面圖" AutoEventWireup="true" CodeFile="map.aspx.cs" Inherits="Lib_map" %>

    <html>

    <head runat="server">
        <script type="text/javascript">
            function Fobj(str) {
                return document.getElementById(str);
            }

            //定位坐標X 定位坐標Y 圖寬格數   圖高格數  每格長(pix) 每格寬(pix)
            X = 1;
            Y = 1;
            Xno = 105;
            Yno = 33;
            Xw = 13;
            Yh = 14;
            //每格長寬在Page_Load pos1 pos2定位時，亦有引用，須一併修改

            IO = "<%Response.Write(Request[" IO "]);%>";
            if (IO == "") IO = "I" //移入或移出機房

            function PageLoad() {
                
                window.resizeTo(Xno * Xw + 16, Yno * Yh + 85); //調整顯示視窗大小
                window.moveTo((screen.width - Xno * Xw) / 2, (screen.height - Yno * Yh) / 2); //移動視窗顯示位置    

                Fobj("bigplot").style.width = Xno * Xw; //調整機房平面圖長度
                Fobj("bigplot").style.height = Yno * Yh; //調整機房平面圖寬度
                Fobj("pos1").style.width = Xw - 4; //調整移入定位點寬
                Fobj("pos1").style.height = Yh - 2; //調整移入定位點高
                Fobj("pos2").style.width = Xw - 4; //調整移出定位點寬
                Fobj("pos2").style.height = Yh - 2; //調整移出定位點高

                Fobj("BtnOK").style.left = 64 * Xw; //調整按鈕位置
                Fobj("BtnOK").style.top = 2 * Yh; //調整按鈕位置            
                Fobj("BtnNo").style.left = 64 * Xw; //調整按鈕位置
                Fobj("BtnNo").style.top = 5 * Yh; //調整按鈕位置
                Fobj("BtnSearch").style.left = 64 * Xw; //調整按鈕位置
                Fobj("BtnSearch").style.top = 8 * Yh; //調整按鈕位置
            }

            function MouseMove() //坐標指針	
            {

                Lx.style.left = Fobj("bigplot").style.posLeft;
                Lx.style.top = window.event.y - 12;
                Rx.style.left = Fobj("bigplot").style.posLeft + Fobj("bigplot").width - 50;
                Rx.style.top = window.event.y - 12;
                Ty.style.left = window.event.x - 4;
                //Ty.style.top = Fobj("bigplot").style.posTop - 8;
                Ty.style.top = -4;
                By.style.left = window.event.x - 4;
                //By.style.top = Fobj("bigplot").style.posTop + Fobj("bigplot").height - 28;
                By.style.top = -4;
            }

            function MapClick() //標示選取位置
            {

                switch (IO) {
                    case "I":
                        Fobj("pos2").style.left = -100;
                        Fobj("pos2").style.top = -100;

                        X = Math.floor((window.event.x) / Xw, 10); //int無條件捨去
                        Y = Math.floor((window.event.y) / Yh, 10); //int無條件捨去
                        Fobj("pos1").style.left = X * Xw + 2;
                        Fobj("pos1").style.top = Y * Yh + 2;

                        break;
                    case "O":
                        Fobj("pos1").style.left = -100;
                        Fobj("pos1").style.top = -100;

                        X = Math.floor((window.event.x) / Xw, 10); //int無條件捨去
                        Y = Math.floor((window.event.y) / Yh, 10); //int無條件捨去
                        Fobj("pos2").style.left = X * Xw + 2;
                        Fobj("pos2").style.top = Y * Yh + 2;

                        break;
                }
            }

            function ReturnPointerNo(PointerNo) //回傳設定位編號	//opener.document.form1.TextPointerNo.value=-(X-2)*31-Y; 
            {
                for (j = 0; j < opener.document.all.length; j++) { //IDMS用
                    if (opener.document.all(j).id.indexOf("TextPointerNo") >= 0) opener.document.all(j).value = PointerNo;
                }
                for (j = 0; j < opener.document.all.length; j++) { //設備進出用
                    if (opener.document.all(j).id.indexOf("TextDevNo") >= 0) opener.document.all(j).value = PointerNo;
                    if (opener.document.all(j).id.indexOf("Xall") >= 0) opener.document.all(j).value = X;
                    if (opener.document.all(j).id.indexOf("Yall") >= 0) opener.document.all(j).value = Y;
                }
            }

            function Search() {
                window.open('../Device/Device.aspx?Search=Map&X=' + (X - 3) + '&Y=' + Y);
            }

            function Cancel() {
                window.close();
            }

            function OKfun() {
                if (Y >= 0) {
                    ReturnPointerNo(-(X - 2) * 31 - Y);
                    window.close();
                } else alert('您尚未定位 ! ');
            }
        </script>
    </head>

    <body onload="PageLoad();">
        <form id="form1" runat="server">
            <img runat="server" id="bigplot" alt="" border="0" src="../images/bigplot.jpg" style="position: absolute;
        left: 0; top: 0" width="1000" height="560" onclick="MapClick();" onmousemove='MouseMove();' />
            <img runat="server" id="pos1" alt="" border="0" src="../images/In.gif" onclick="MapClick();" onmousemove="MouseMove();" style="z-index: 5; position: absolute; left: -100; top: -100;" />
            <img runat="server" id="pos2" alt="" border="0" src="../images/Out.gif" onclick="MapClick();" onmousemove="MouseMove();" style="z-index: 5; position: absolute; left: -100; top: -100; width:5; height:6;" /> &nbsp;
            <div runat="server" id='Lx' style='z-index: 1; color: red; font-weight: bold; position: absolute'>
                －
            </div>
            <div runat="server" id='Rx' style='z-index: 1; color: red; font-weight: bold; position: absolute'>
                －
            </div>
            <div runat="server" id='Ty' style='z-index: 1; color: red; position: absolute'>
                |
            </div>
            <div runat="server" id='By' style='z-index: 1; color: red; position: absolute'>
                |
            </div>
            <input runat="server" type="button" id="BtnOK" style="z-index: 2; font-weight: bold;
        position: absolute; left: 0; top: 0; color: #FF0000; font-size: medium; background-color: #FFFF00" value="確定" onclick="OKfun()" />

            <input runat="server" type="button" id="BtnSearch" style="z-index: 2; font-weight: bold;
        position: absolute; left: 0; top: 0; color: #FF0000; font-size: medium; background-color: #FFFF00" value="查詢" onclick="Search()" />

            <input runat="server" type="button" id="BtnNo" style="z-index: 2; font-weight: bold;
        position: absolute; left: 0; top: 0; color: #FF0000; font-size: medium; background-color: #FFFF00" value="取消" onclick="Cancel()" />
        </form>
    </body>

    </html>