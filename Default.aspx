<%@ Page Title="資源管理" Language="C#" MasterPageFile="~/IDMS/MasterPage.master" AutoEventWireup="true" CodeFile="Default.aspx.cs" Inherits="_Default" %>

<asp:Content ID="head" ContentPlaceHolderID="head" Runat="Server">
    <style type="text/css">
        .style4
        {
            color: #FF0000;
        }
    </style>
</asp:Content>
<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server">
<br />
【公告】<br />
<!--
1. 2022/7/15 起恢復連結 ISMS 取得資訊資產清冊<br />
&nbsp;&nbsp;&nbsp;&nbsp; a. 最新版本之資產編號已匯入至實體設備、系統資源編輯頁面選單(每日凌晨 02:40 同步更新)。<br />
&nbsp;&nbsp;&nbsp;&nbsp; b. 部分於連結中斷期間以手動方式新增之資產編號，與目前ISMS資產清冊衝突(編號重複但內容不一致)之調整情形請參閱 <a href="http://10.6.1.11/IDMS/Help/資產編號異動清單_20220715.pdf" target="_blank">資產編號異動清單_20220715.pdf</a> 。<br />																				
&nbsp;&nbsp;&nbsp;&nbsp; c. 對於設定已註銷資產編號之實體設備，系統將自動於每週一上午 07:50 Mail發送提醒訊息(詳見 生命履歷 -> 資料查核 -> 無資產編號)，請相關負責人員留意。<br />
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; - 已於111年8月底恢復維護人員例行通知，若您有資料批次修改需求(如 設備編號 -> 資產編號、原資產編號 -> 新資產編號 等)，請於111年9月至112年2月期間與本組聯繫討論。相關人員通知將於後續擇日啟用。<br>
若有任何問題，請與IDMS負責人(資訊中心操作課 SSM小組)聯繫反映，或請機房OP代為轉達，感謝您的配合與協助。<br />
<br />
<br /> -->
1. http://10.6.1.4:8888資料庫已轉移至 http://10.6.1.11:8888 <br />
http://10.6.1.4:8888已下線，兩者資料不互通。<br />
<font color="red">因秘總財產資料庫已不支援，若在秘總財產有更動的資料，請手動填入相關表格欄位</font> <br /><br />
若新版有任何問題請聯絡IDMS負責人，或請機房OP代為轉達(留日誌)
<br />
<br />
<br />
<br />
<br />
<br />
<br />
<br />
<br />
【系統定義】：各課不自有機制，而以機房為中心，所共同遵循的一套資源資料庫共用系統。<br />
<br />
【測試路徑】：<a href="http://10.6.1.11:12345/IDMS" target="_self"><font color="blue">說明文件 -> (測試機)</font></a> ，此處資料可任意異動！<br />
<br />
【系統架構】：[系統資源]是以安裝於[實體設備]上的[作業主機]所具有的種種[系統功能]所構成<p align="center"><img src="images/IDMS.jpg" /></p>
【資源範圍】：實體設備(硬體、通訊、環境)、作業主機(軟體)、系統資源(軟體、資料)、秘總財產(硬體)、套裝軟體(軟體)、組織群組(人員)、放置地點(環境)。<br />
<br />
【管理機制】：<span class="style4"><strong>維護人員</strong></span>對資料負有完整性(有用的資料都有輸入)、即時性(在需要運用之前，資料已經輸入)、正確性(和實際狀況或需求相符)之責，故：<br />
&nbsp;&nbsp;&nbsp; (1) 資料問題請找維護人員，系統問題則向機房(sosop@cwb.gov.tw)反映 <br />
&nbsp;&nbsp;&nbsp; (2) 詳細內容請參考 <a href="Help/權限管理與生命履歷.txt" target="_blank"><font color="blue">說明文件 -> 系統手冊 -> 權限說明</font></a>。<br />
&nbsp;&nbsp;&nbsp; (3) 基於互惠原則：使用IDMS發現資料有問題，應向維護人員反映更正。<br />
<br />
【權責分工】：<br />
&nbsp;&nbsp;&nbsp; 1.【設備管理】由硬體負責人負責<br />
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; a.系統設備、零附件、週邊設備多半由系統課負責，以實際維護人員為準<br />
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; b.網路設備由網路課網管小組負責<br />
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; c.環境設備由操作課環境小組負責<br />
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; d.辦公設備由保管人員負責<br />
<br />
&nbsp;&nbsp;&nbsp; 2.【作業管理】由資訊系統負責人負責<br />
<br />
&nbsp;&nbsp;&nbsp; 3.【套裝軟體】由軟體小組負責<br />
<br />
&nbsp;&nbsp;&nbsp; 4.【秘總財產】由保管人員/技術支援課負責<br />
<br />
&nbsp;&nbsp;&nbsp; 5.【機房專區】由操作課環境小組負責<br />
<br />
&nbsp;&nbsp;&nbsp; 6.【系統設定】程式、設定與文件由操作課SSM小組負責<br />

</asp:Content>
