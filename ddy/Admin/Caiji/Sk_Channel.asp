<!--#include file="inc/setup.asp"-->
<!--#include file="inc/cj_cls.asp"-->

<%
Dim CurrentPage,MaxPerPage
Dim DateStr
If IsSqlDataBase = 1 Then '�Ƿ�SQL���ݿ�
	DateStr = "smalldatetime"
Else
	DateStr = "date"
End If
%>
<html>
<head>
<title><%= Skcj.GetConfig("WebName") %></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" type="text/css" href="css/Admin_Style.css">
</head>

<body>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="tableBorder">
  <tr>
    <td height="22" colspan="2" align="center" bgcolor="#F3F3F3" class="topbg"><strong>�ɼ�ģ�����</strong></td>
  </tr>
</table>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1"  class="tableBorder">
  <tr class="tdbg"> 
    <td width="65" height="30" bgcolor="#F2F2F2"><strong>����������</strong></td>
    <td height="30" bgcolor="#F2F2F2"><a href=SK_Channel.asp>������ҳ</a> | <a href="?Action=add">����ģ��</a></td>
  </tr>
</table>
<%
Action=LCase(Trim(Request("Action")))
Select Case Action
	
Case "add"
	 Call Add()
Case "edit"
	 Call Edit()
Case "del"
	 Call Del()
Case "flag"
	ConnItem.execute("update SK_CJ set Flag="&SKcj.G("Flag")&" Where ID="&Skcj.ChkNumeric(SKcj.G("ID"))) 
	Call Show_Manage()
Case else
	Call Show_Manage()
End Select


Sub Show_Manage()'�ɼ�ģ���б�����
Const MaxPerPage=15
If Request("page")<>"" then
	CurrentPage=Cint(Request("Page"))
Else
	CurrentPage=1
End if  
%>
<table width="100%" align="center" cellpadding="2" cellspacing="1" class="tableBorder">
  <tr>
    <th colspan="9" >�ɼ�ģ���б�</th>
  </tr>
  <tr align="center" >
    <td width="8%" class="title3" style="font-weight: bold"> ģ������ </td>
	<td width="15%" class="title3" style="font-weight: bold"> ����Ŀ¼ </td>
    <td width="4%" class="title3" style="font-weight: bold">����ID</td>
    <td width="4%" class="title3" style="font-weight: bold">״̬</td>
    <td width="10%" class="title3" style="font-weight: bold">����</td>
  </tr>
  <%
dim Rs
Sql = "Select * from SK_cj Order By OrderID DESC"
Set Rs=server.createobject("adodb.recordset") 
Rs.open Sql,ConnItem,1,1
if not Rs.bof and not Rs.eof then
	Rs.PageSize=MaxPerPage
 	Allpage=Rs.PageCount
    If Currentpage>Allpage Then Currentpage=1
    ItemNum=Rs.RecordCount
    Rs.MoveFirst
    Rs.AbsolutePage=CurrentPage
	I=0
	do while not Rs.eof
%>
  <tr>
    <td align="center" bgcolor="#F6F6F6" class="td">
	<%
	Response.Write Rs("CjName") 
	'Set Rsl=ConnItem.execute("Select ID,CjName from SK_CJ where ID="& RS("skcjID"))
	'if not Rsl.eof Then  Response.Write(Rsl(1))
	'Rsl.close : Set Rsl=Nothing 
	%>
	</td>
	<td align="left" bgcolor="#F6F6F6" class="td"><%=Rs("Dir")%></td>
    <td align="left" bgcolor="#F6F6F6" class="td"><%=Rs("OrderID")%></td>
    <td align="center" bgcolor="#F6F6F6" class="td">
		<%if Rs("Flag") = 1 then
		 response.Write "<a href=""?action=Flag&Flag=0&ID="& Rs("ID") &""">��</a>"
		 else  
		response.Write("<a href=""?action=Flag&Flag=1&ID="& Rs("ID") &"""><font color=""#FF0000"">��</font></a>")  
		end if%>    
	</td>
    <td width="7%" align="center" bgcolor="#F6F6F6" class="td">
	<a href="?action=Edit&ID=<%=Rs("ID")%>">�޸�</a>&nbsp;&nbsp;<a href="?action=Del&ID=<%=Rs("ID")%>" onClick='return confirm("ȷ��Ҫɾ���˼�¼��");'>ɾ��</a>	</td>
  </tr>
  <%
  	I=I+1
  	If I>=MaxPerPage Then  Exit  Do
	Rs.movenext
	loop
End if
%>
</table>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="tableBorder" >
  <tr>
    <td height="22" colspan="2" class="tdbg"><%
Response.Write ShowPage("SK_Channel.asp",ItemNum,MaxPerPage,True,True," ����Ŀ")
%>
    </td>
  </tr>
</table>
<%End Sub%>
<%
Sub Add()
If Trim(Request.Form("SaveOk"))="ok" then
	SqlItem="Select top 1 * from SK_Cj where ID=1"
    Set Rs=server.CreateObject("adodb.recordset")
    Rs.Open SqlItem,ConnItem,1,3
	Rs.addnew
	rs("CjName")=Skcj.G("CjName")
	rs("FileName")=Skcj.G("FileName")
	rs("Timeout")=Skcj.ChkNumeric(Skcj.G("Timeout"))
	rs("Dir")=Skcj.G("Dir")
	rs("MaxFileSize")=Skcj.ChkNumeric(Skcj.G("MaxFileSize"))
	rs("FileExtName")=Skcj.G("FileExtName")
	rs("MaxPerPage")=Skcj.ChkNumeric(Skcj.G("MaxPerPage"))
	rs("Flag")=Skcj.ChkNumeric(Skcj.G("Flag"))
	rs("OrderID")=Skcj.ChkNumeric(Skcj.G("OrderID"))
	Rs.UpDate
    Rs.Close	
	set rs=nothing
	response.write "<script>alert('���ӳɹ�!');</script>"'�رմ��� 
End if
%>
<SCRIPT LANGUAGE="JavaScript">
<!--
function CheckForm()
{
  if (document.myform.CjName.value==''){
    alert('ģ�����Ʋ���Ϊ��!');
    document.myform.SK_CJ.focus();
    return false;
  }
}
//-->
</SCRIPT>
<table width="100%" align="center" cellpadding="2" cellspacing="1" class="tableBorder">
      <form name="myform" method="post" action="?action=add" onSubmit="javascript:return CheckForm();">
  <tr>
    <td colspan="2" class="title">���Ӳɼ�ģ��</td>
  </tr>
  <tr>
    <td width="20%" bgcolor="#F5F5F5" class="td">ģ�����ƣ�</td>
    <td width="80%" bgcolor="#F5F5F5" class="td"><input name="SaveOk" class="lostfocus"  type="hidden" value="ok">
	<input name="CjName" type="text" class="lostfocus"> 
      <strong><font color=ff0000  >* </font></td>
    </tr>
  
  <tr>
    <td width="20%" bgcolor="#F5F5F5" class="td">ģ���ļ�����      </td>
    <td width="80%" bgcolor="#F5F5F5" class="td"><input name="FileName" type="text" class="lostfocus" value=".asp"></td>
    </tr>
	
  <tr>
    <td bgcolor="#F5F5F5" class="td">����Ŀ¼��      </td>
    <td bgcolor="#F5F5F5" class="td"><input name="Dir" type="text" class="lostfocus" value="New/">
      <font class="alert">&nbsp;��������"/"����</font></td>
  </tr>
  <tr>
    <td bgcolor="#F5F5F5" class="td">�������ص�ͼƬ��С��</td>
    <td bgcolor="#F5F5F5" class="td"><input name="MaxFileSize" type="text" class="lostfocus" value="5000" size="15" maxlength="9">
    <STRONG>KB </STRONG> <font class="alert">* �����������롰0��</font></td>
  </tr>
  <tr>
    <td bgcolor="#F5F5F5" class="td">�ɼ������ļ����ͣ�</td>
    <td bgcolor="#F5F5F5" class="td"><input name="FileExtName" type="text" class="lostfocus" value="rm|swf" size="50" maxlength="300"></td>
  </tr>
  
  
  <tr>
    <td bgcolor="#F5F5F5" class="td">�ɼ���ʱ���ã�</td>
    <td bgcolor="#F5F5F5" class="td"><input name="Timeout" type="text" class="lostfocus" value="64" size="10" maxlength="9">
    <STRONG>��</STRONG> <font class="alert">* Ĭ��64�� ���128K������64�뻹���ز���(��ÿ��2K���ع���)����ʱ��</font></td>
    </tr>
  <tr>
    <td bgcolor="#F5F5F5" class="td">��Ŀ�������ã�</td>
    <td bgcolor="#F5F5F5" class="td"><input name="MaxPerPage" type="text" class="lostfocus" value="10" size="10" maxlength="9"></td>
  </tr>
  
  <tr>
    <td bgcolor="#F5F5F5" class="td">�������ã�</td>
    <td bgcolor="#F5F5F5" class="td"><input name="OrderID" type="text" class="lostfocus" size="5"></td>
    </tr>
	<tr>
    <td bgcolor="#F5F5F5" class="td">�Ƿ����ã�</td>
    <td bgcolor="#F5F5F5" class="td">
		<input type="radio" name="Flag" value="1" checked>
		��
		<input type="radio" name="Flag" value="0">
        ��	  </td>
    </tr>
  <tr>
    <td colspan="2" align="center" bgcolor="#F5F5F5" class="td"><input type="submit" name="Submit" value="  ȷ������  "></td>
  </tr>
      </form>
</table>
<%End Sub
Sub Edit()
	Dim Rs,ID,FieldName,ItemType,FieldType,AllowNull,FieldIntro,Flag
	ID=Skcj.G("ID")
	If Trim(Request.Form("Saveok"))="ok" then
		SqlItem="Select top 1 * from SK_Cj where ID="& ID
		Set Rs=server.CreateObject("adodb.recordset")
		Rs.Open SqlItem,ConnItem,1,3
		rs("CjName")=Skcj.G("CjName")
		rs("FileName")=Skcj.G("FileName")
		rs("Timeout")=Skcj.ChkNumeric(Skcj.G("Timeout"))
		rs("Dir")=Skcj.G("Dir")
		rs("MaxFileSize")=Skcj.ChkNumeric(Skcj.G("MaxFileSize"))
		rs("FileExtName")=Skcj.G("FileExtName")
		rs("MaxPerPage")=Skcj.ChkNumeric(Skcj.G("MaxPerPage"))
		rs("Flag")=Skcj.ChkNumeric(Skcj.G("Flag"))
		rs("OrderID")=Skcj.ChkNumeric(Skcj.G("OrderID"))
		Rs.UpDate
		Rs.Close	
		set rs=nothing
		response.write "<script>alert('�޸ĳɹ���');location.href='SK_Channel.asp';</script>"'�رմ���
	End if
	
		SqlItem="Select top 1 * from SK_Cj where ID="& ID
		Set Rs=server.CreateObject("adodb.recordset")
		Rs.Open SqlItem,ConnItem,1,3
		IF not rs.eof then 
		CjName=rs("CjName")
		FileName=rs("FileName")
		Timeout=rs("Timeout")
		Dir=rs("Dir")
		MaxFileSize=rs("MaxFileSize")
		FileExtName=rs("FileExtName")
		MaxPerPage=rs("MaxPerPage")
		Flag=rs("Flag")
		OrderID=rs("OrderID")
		Rs.UpDate
		End if
		Rs.Close	
		set rs=nothing
	%>
	<SCRIPT LANGUAGE="JavaScript">
<!--
function CheckForm()
{
  if (document.myform.CjName.value==''){
    alert('��ģ�����ơ�����Ϊ�գ�');
    //document.myform.SaveOk.focus();
    return false;
  }
}
//-->
</SCRIPT>

<table width="100%" align="center" cellpadding="2" cellspacing="1" class="tableBorder">
      <form name="myform" method="post" action="?action=Edit&ID=<%=ID%>" onSubmit="javascript:return CheckForm();">
  <tr>
    <td colspan="2" class="title">�޸Ĳɼ�ģ��</td>
  </tr>
  <tr>
    <td width="20%" bgcolor="#F5F5F5" class="td">ģ�����ƣ�</td>
    <td width="80%" bgcolor="#F5F5F5" class="td"><input name="SaveOk" class="lostfocus"  type="hidden" value="ok">
	<input name="CjName" type="text" class="lostfocus" value=<%= CjName %>> 
      <strong><font color=ff0000  >* </font></td>
    </tr>
  
  <tr>
    <td width="20%" bgcolor="#F5F5F5" class="td">ģ���ļ�����      </td>
    <td width="80%" bgcolor="#F5F5F5" class="td"><input name="FileName" type="text" class="lostfocus" value="<%= FileName %>"></td>
    </tr>
	
  <tr>
    <td bgcolor="#F5F5F5" class="td">����Ŀ¼��      </td>
    <td bgcolor="#F5F5F5" class="td"><input name="Dir" type="text" class="lostfocus" value="<%= Dir %>" >
      <font class="alert">&nbsp;��������"/"����</font></td>
  </tr>
  <tr>
    <td bgcolor="#F5F5F5" class="td">�������ص�ͼƬ��С��</td>
    <td bgcolor="#F5F5F5" class="td"><input name="MaxFileSize" type="text" class="lostfocus" value="<%= MaxFileSize %>" size="15" maxlength="9">
    <STRONG>KB </STRONG> <font class="alert">* �����������롰0��</font></td>
  </tr>
  <tr>
    <td bgcolor="#F5F5F5" class="td">�ɼ������ļ����ͣ�</td>
    <td bgcolor="#F5F5F5" class="td"><input name="FileExtName" type="text" class="lostfocus" value="<%= FileExtName %>" size="50" maxlength="300"></td>
  </tr>
  
  
  <tr>
    <td bgcolor="#F5F5F5" class="td">�ɼ���ʱ���ã�</td>
    <td bgcolor="#F5F5F5" class="td"><input name="Timeout" type="text" class="lostfocus" value="<%= Timeout %>" size="10" maxlength="9">
    <STRONG>��</STRONG> <font class="alert">* Ĭ��64�� ���128K������64�뻹���ز���(��ÿ��2K���ع���)����ʱ��</font></td>
    </tr>
  <tr>
    <td bgcolor="#F5F5F5" class="td">��Ŀ�������ã�</td>
    <td bgcolor="#F5F5F5" class="td"><input name="MaxPerPage" type="text" class="lostfocus" value="<%= MaxPerPage %>" size="10" maxlength="9"></td>
  </tr>
  
  <tr>
    <td bgcolor="#F5F5F5" class="td">�������ã�</td>
    <td bgcolor="#F5F5F5" class="td"><input name="OrderID" type="text" class="lostfocus" size="5" value="<%= OrderID %>"></td>
    </tr>
	<tr>
    <td bgcolor="#F5F5F5" class="td">�Ƿ����ã�</td>
    <td bgcolor="#F5F5F5" class="td">
		<input type="radio" name="Flag" value="1" <% IF Flag=1 then Response.Write"checked" %> >
		��
		<input type="radio" name="Flag" value="0" <% IF Flag=0 then Response.Write"checked" %>>
        ��	  </td>
    </tr>
  <tr>
    <td colspan="2" align="center" bgcolor="#F5F5F5" class="td"><input type="submit" name="Submit" value="  ȷ���޸�  "></td>
  </tr>
      </form>
</table>
	<%
End Sub

Sub Del()
	Dim ID
	ID=Trim(Request("ID"))
	ConnItem.Execute("Delete From SK_CJ Where ID = "& ID )
	response.write "<script>location.href='SK_Channel.asp';</script>"
End Sub

'==================================================
'��������WriteErrMsg
'��  �ã���ʾ������ʾ��Ϣ
'��  ������
'==================================================
sub WriteErrMsg(ErrMsg)
	dim strErr
	strErr=strErr & "<table cellpadding=2 cellspacing=1 border=0 width=400 class=""tableBorder"" align=center>" & vbcrlf
	strErr=strErr & "  <tr align='center' class='title'><td height='22'><strong>������Ϣ</strong></td></tr>" & vbcrlf
	strErr=strErr & "  <tr class='tdbg'><td height='100' valign='top'><b>��������Ŀ���ԭ��</b>" & ErrMsg &"</td></tr>" & vbcrlf
	strErr=strErr & "  <tr align='center' class='tdbg'><td><a href='javascript:history.go(-1)'>&lt;&lt; ������һҳ</a></td></tr>" & vbcrlf
	strErr=strErr & "</table>" & vbcrlf
	response.write strErr
	response.end
end sub
%>
</body>
</html>