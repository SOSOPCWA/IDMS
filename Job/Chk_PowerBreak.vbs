'-----------------------------------------------------------------
strConn="Driver={SQL SERVER};Server=10.6.1.4;Trusted_Connection=True;Database=" 
Dim connIDMS    : Set connIDMS=CreateObject("ADODB.Connection")   : connIDMS.Open strConn & "IDMS"

'------------------���o�_�l�s��-------------------------------------
'on error resume next

dim rs1 : set rs1=createobject("ADODB.Recordset")
rs1.open "select distinct [�W��s��] from [���q] where [�W��s��]<>-1 and [�W��s��]<>-2",connIDMS
while not rs1.eof
  str=GetRoot(rs1(0))
  if str<>-1 then strs=strs & str & ","
  rs1.movenext
wend
rs1.close

Function GetRoot(DevNo)
  GetRoot=DevNo
  dim rs : set rs=createobject("ADODB.Recordset")
  rs.open "select [�W��s��] from [���q] where [�U��s��]=" & DevNo,connIDMS
  if not rs.eof then GetRoot=GetRoot(rs(0))
  rs.close
End Function

msgbox "�_�I�G" & strs

