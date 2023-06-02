'-----------------------------------------------------------------
strConn="Driver={SQL SERVER};Server=10.6.1.4;Trusted_Connection=True;Database=" 
Dim connIDMS    : Set connIDMS=CreateObject("ADODB.Connection")   : connIDMS.Open strConn & "IDMS"

'------------------取得起始編號-------------------------------------
'on error resume next

dim rs1 : set rs1=createobject("ADODB.Recordset")
rs1.open "select distinct [上游編號] from [接電] where [上游編號]<>-1 and [上游編號]<>-2",connIDMS
while not rs1.eof
  str=GetRoot(rs1(0))
  if str<>-1 then strs=strs & str & ","
  rs1.movenext
wend
rs1.close

Function GetRoot(DevNo)
  GetRoot=DevNo
  dim rs : set rs=createobject("ADODB.Recordset")
  rs.open "select [上游編號] from [接電] where [下游編號]=" & DevNo,connIDMS
  if not rs.eof then GetRoot=GetRoot(rs(0))
  rs.close
End Function

msgbox "斷點：" & strs

