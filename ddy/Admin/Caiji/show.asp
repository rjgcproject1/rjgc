
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" type="text/css" href="css/Admin_Style.css">
<title></title>
<style type="text/css">
<!--
body {
	background-color: #FFFFFF;
}
-->
</style></head>

<body>
<!--#include file="inc/setup.asp"-->

<%
ID=Trim(skcj.G("ID"))
Lx=Trim(skcj.G("Lx"))
if ID="" then Response.end 
Select Case Lx
Case 1
	Sql="select * from SK_Article where ArticleID=" & id 
Case 2
	Sql="select * from SK_Photo where ID=" & id 
Case 3
	Sql="select * from SK_DownLoad where ID=" & id 
Case 4
	Sql="select * from SK_Flash where ID=" & id 	
End Select
Set Rs=ConnItem.execute(Sql)
If not rs.eof then
%>
<table width="70%" border="0" align="center" bordercolor="#CCCCCC" bgcolor="#FFFFFF" style="border:#CCCCCC 1px solid">
  <%
  Select Case Lx
  Case 1
  	Call Get_Article
	Dim tRs,RsItem,StrTemp,Id,Sql
	     Dim ZdType,Zds,Zdo,Zd,SKZDStr,SKZDStr_arr
		 Sql="select * from SK_ZDField where Flag=1 Order By ID DESC"
		 Set tRs=ConnItem.execute(Sql)
		 if not tRs.bof and not tRs.eof then
		 	While not tRs.eof
				Id=tRs("ID")
				FieldName=tRs("FieldName")
				Response.Write "<tr>"
    			Response.Write "<td height=""25"" width=""30%"" align=""center"" valign=""middle"" bgcolor=""#E4E4E4"">"& trs("FieldIntro") &"</td>"
    			Response.Write "<td align=""center"" valign=""middle"" bgcolor=""#F9F9F9"">"& Rs(FieldName) &"</td>"
 				Response.Write "</tr>"
				tRs.movenext
			Wend
		End if
  Case 2
  	Call Get_PHOTO
  Case 3
  	Call Get_soft
  Case 4
    Call Get_Flash	
  End Select 
  %>
</table>
<%
End if
rs.close : set rs=nothing
Sub Get_Article
  Set ARs=ConnItem.execute("Select * from SK_Article Where ArticleID=" & Id)
  IF Not Rs.eof then
	  Response.Write "<tr><td height=""25"" width=""30%"" align=""center"" valign=""middle"" bgcolor=""#E4E4E4"">标题</td>"
	  Response.Write "<td align=""center"" valign=""middle"" bgcolor=""#F9F9F9"">"& Ars("Title") &"</td></tr>"
	  Response.Write "<tr><td height=""25"" width=""30%"" align=""center"" valign=""middle"" bgcolor=""#E4E4E4"">作者</td>"
	  Response.Write "<td align=""center"" valign=""middle"" bgcolor=""#F9F9F9"">"& Ars("Author") &"</td></tr>"
	  Response.Write "<tr><td height=""25"" width=""30%"" align=""center"" valign=""middle"" bgcolor=""#E4E4E4"">来源</td>"
	  Response.Write "<td align=""center"" valign=""middle"" bgcolor=""#F9F9F9"">"& Ars("CopyFrom") &"</td></tr>"
	  Response.Write "<tr><td height=""25"" width=""30%"" align=""center"" valign=""middle"" bgcolor=""#E4E4E4"">录入者</td>"
	  Response.Write "<td align=""center"" valign=""middle"" bgcolor=""#F9F9F9"">"& Ars("Inputer") &"</td></tr>"
	  Response.Write "<tr><td height=""25"" width=""30%"" align=""center"" valign=""middle"" bgcolor=""#E4E4E4"">关键字</td>"
	  Response.Write "<td align=""center"" valign=""middle"" bgcolor=""#F9F9F9"">"& Ars("Keyword") &"</td></tr>"
	  Response.Write "<tr><td height=""25"" width=""30%"" align=""center"" valign=""middle"" bgcolor=""#E4E4E4"">更新时间</td>"
	  Response.Write "<td align=""center"" valign=""middle"" bgcolor=""#F9F9F9"">"& Ars("UpdateTime") &"</td></tr>"
	  Response.Write "<tr><td height=""25"" width=""30%"" align=""center"" valign=""middle"" bgcolor=""#E4E4E4"">文章内容</td>"
	  Response.Write "<td align=""center""   valign=""middle"" bgcolor=""#F9F9F9"">"& Ars("Content")  &"</td></tr>"
	  Response.Write "<tr><td height=""25"" width=""30%"" align=""center"" valign=""middle"" bgcolor=""#E4E4E4"">小图片</td>"
	  Response.Write "<td align=""center"" valign=""middle"" bgcolor=""#F9F9F9"">"& Ars("picpath") &"</td></tr>"
  End if
  Ars.Close
  Set Ars=nothing
End Sub
Sub Get_Photo
  Set ARs=ConnItem.execute("Select * from SK_Photo Where ID=" & Id)
  IF Not Rs.eof then
	  Response.Write "<tr><td height=""25"" width=""30%"" align=""center"" valign=""middle"" bgcolor=""#E4E4E4"">图片名称</td>"
	  Response.Write "<td align=""center"" valign=""middle"" bgcolor=""#F9F9F9"">"& Ars("Title") &"</td></tr>"
	  Response.Write "<tr><td height=""25"" width=""30%"" align=""center"" valign=""middle"" bgcolor=""#E4E4E4"">缩图</td>"
	  Response.Write "<td align=""center"" valign=""middle"" bgcolor=""#F9F9F9"">"& Ars("PhotoUrl") &"</td></tr>"
	  Response.Write "<tr><td height=""25"" width=""30%"" align=""center"" valign=""middle"" bgcolor=""#E4E4E4"">图片地址</td>"
	  Response.Write "<td align=""center"" valign=""middle"" bgcolor=""#F9F9F9"">"& Ars("PicUrls") &"</td></tr>"
	  Response.Write "<tr><td height=""25"" width=""30%"" align=""center"" valign=""middle"" bgcolor=""#E4E4E4"">图片简介</td>"
	  Response.Write "<td align=""center"" valign=""middle"" bgcolor=""#F9F9F9"">"& Ars("PictureContent") &"</td></tr>"
	  Response.Write "<tr><td height=""25"" width=""30%"" align=""center"" valign=""middle"" bgcolor=""#E4E4E4"">作者</td>"
	  Response.Write "<td align=""center"" valign=""middle"" bgcolor=""#F9F9F9"">"& Ars("Author") &"</td></tr>"
	  Response.Write "<tr><td height=""25"" width=""30%"" align=""center"" valign=""middle"" bgcolor=""#E4E4E4"">来源</td>"
	  Response.Write "<td align=""center"" valign=""middle"" bgcolor=""#F9F9F9"">"& Ars("Origin") &"</td></tr>"
	  Response.Write "<tr><td height=""25"" width=""30%"" align=""center"" valign=""middle"" bgcolor=""#E4E4E4"">时间</td>"
	  Response.Write "<td align=""center""   valign=""middle"" bgcolor=""#F9F9F9"">"& Ars("AddDate")  &"</td></tr>"
  End if
  Ars.Close
  Set Ars=nothing
End Sub
Sub Get_soft
  Set ARs=ConnItem.execute("Select * from SK_DownLoad Where ID=" & Id)
  IF Not Rs.eof then
	  Response.Write "<tr><td height=""25"" width=""30%"" align=""center"" valign=""middle"" bgcolor=""#E4E4E4"">下载名称</td>"
	  Response.Write "<td align=""center"" valign=""middle"" bgcolor=""#F9F9F9"">"& Ars("Title") &"</td></tr>"
	  Response.Write "<tr><td height=""25"" width=""30%"" align=""center"" valign=""middle"" bgcolor=""#E4E4E4"">下载类别</td>"
	  Response.Write "<td align=""center"" valign=""middle"" bgcolor=""#F9F9F9"">"& Ars("DownLB") &"</td></tr>"
	  Response.Write "<tr><td height=""25"" width=""30%"" align=""center"" valign=""middle"" bgcolor=""#E4E4E4"">下载语言</td>"
	  Response.Write "<td align=""center"" valign=""middle"" bgcolor=""#F9F9F9"">"& Ars("DownYY") &"</td></tr>"
	  Response.Write "<tr><td height=""25"" width=""30%"" align=""center"" valign=""middle"" bgcolor=""#E4E4E4"">下载授权</td>"
	  Response.Write "<td align=""center"" valign=""middle"" bgcolor=""#F9F9F9"">"& Ars("DownSQ") &"</td></tr>"
	  Response.Write "<tr><td height=""25"" width=""30%"" align=""center"" valign=""middle"" bgcolor=""#E4E4E4"">下载平台</td>"
	  Response.Write "<td align=""center"" valign=""middle"" bgcolor=""#F9F9F9"">"& Ars("DownPT") &"</td></tr>"
	  Response.Write "<tr><td height=""25"" width=""30%"" align=""center"" valign=""middle"" bgcolor=""#E4E4E4"">大小</td>"
	  Response.Write "<td align=""center"" valign=""middle"" bgcolor=""#F9F9F9"">"& Ars("DownSize") &"</td></tr>"
	  Response.Write "<tr><td height=""25"" width=""30%"" align=""center"" valign=""middle"" bgcolor=""#E4E4E4"">演示地址</td>"
	  Response.Write "<td align=""center""   valign=""middle"" bgcolor=""#F9F9F9"">"& Ars("YSDZ")  &"</td></tr>"
	  Response.Write "<tr><td height=""25"" width=""30%"" align=""center"" valign=""middle"" bgcolor=""#E4E4E4"">注册地址</td>"
	  Response.Write "<td align=""center"" valign=""middle"" bgcolor=""#F9F9F9"">"& Ars("ZCDZ") &"</td></tr>"
	  Response.Write "<tr><td height=""25"" width=""30%"" align=""center"" valign=""middle"" bgcolor=""#E4E4E4"">缩略图</td>"
	  Response.Write "<td align=""center"" valign=""middle"" bgcolor=""#F9F9F9"">"& Ars("PhotoUrl") &"</td></tr>"
	  Response.Write "<tr><td height=""25"" width=""30%"" align=""center"" valign=""middle"" bgcolor=""#E4E4E4"">下载简介</td>"
	  Response.Write "<td align=""center"" valign=""middle"" bgcolor=""#F9F9F9"">"& left(server.HTMLEncode(Ars("DownContent")),50) &".......</td></tr>"
	  Response.Write "<tr><td height=""25"" width=""30%"" align=""center"" valign=""middle"" bgcolor=""#E4E4E4"">作者</td>"
	  Response.Write "<td align=""center"" valign=""middle"" bgcolor=""#F9F9F9"">"& Ars("Author") &"</td></tr>"
	  Response.Write "<tr><td height=""25"" width=""30%"" align=""center"" valign=""middle"" bgcolor=""#E4E4E4"">下载来源</td>"
	  Response.Write "<td align=""center"" valign=""middle"" bgcolor=""#F9F9F9"">"& Ars("Origin") &"</td></tr>"
  End if
  Ars.Close
  Set Ars=nothing
End Sub
Sub Get_Flash
  Set ARs=ConnItem.execute("Select * from SK_Flash Where ID=" & Id)
  IF Not Rs.eof then
	  Response.Write "<tr><td height=""25"" width=""30%"" align=""center"" valign=""middle"" bgcolor=""#E4E4E4"">动漫名称</td>"
	  Response.Write "<td align=""center"" valign=""middle"" bgcolor=""#F9F9F9"">"& Ars("Title") &"</td></tr>"
	  Response.Write "<tr><td height=""25"" width=""30%"" align=""center"" valign=""middle"" bgcolor=""#E4E4E4"">动漫缩略图</td>"
	  Response.Write "<td align=""center"" valign=""middle"" bgcolor=""#F9F9F9"">"& Ars("PhotoUrl") &"</td></tr>"
	  Response.Write "<tr><td height=""25"" width=""30%"" align=""center"" valign=""middle"" bgcolor=""#E4E4E4"">动漫地址</td>"
	  Response.Write "<td align=""center"" valign=""middle"" bgcolor=""#F9F9F9"">"& Ars("FlashUrl") &"</td></tr>"
	  Response.Write "<tr><td height=""25"" width=""30%"" align=""center"" valign=""middle"" bgcolor=""#E4E4E4"">动漫简介</td>"
	  Response.Write "<td align=""center"" valign=""middle"" bgcolor=""#F9F9F9"">"& left(server.HTMLEncode(Ars("flashContent")),50) &".......</td></tr>"
	  Response.Write "<tr><td height=""25"" width=""30%"" align=""center"" valign=""middle"" bgcolor=""#E4E4E4"">作者</td>"
	  Response.Write "<td align=""center"" valign=""middle"" bgcolor=""#F9F9F9"">"& Ars("Author") &"</td></tr>"
	  Response.Write "<tr><td height=""25"" width=""30%"" align=""center"" valign=""middle"" bgcolor=""#E4E4E4"">时间</td>"
	  Response.Write "<td align=""center"" valign=""middle"" bgcolor=""#F9F9F9"">"& Ars("AddDate") &"</td></tr>"
	  Response.Write "<tr><td height=""25"" width=""30%"" align=""center"" valign=""middle"" bgcolor=""#E4E4E4"">关键字</td>"
	  Response.Write "<td align=""center"" valign=""middle"" bgcolor=""#F9F9F9"">"& Ars("Keywords") &"</td></tr>"
	  Response.Write "<tr><td height=""25"" width=""30%"" align=""center"" valign=""middle"" bgcolor=""#E4E4E4"">动漫来源</td>"
	  Response.Write "<td align=""center"" valign=""middle"" bgcolor=""#F9F9F9"">"& Ars("Origin") &"</td></tr>"
  End if
  Ars.Close
  Set Ars=nothing
End Sub
Call CloseConnItem()
%>
</body>
</html>
