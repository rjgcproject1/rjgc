<%
if request("keywords")<>"" then
if checktxt(request("keywords"))<>request("keywords") then
	response.write "<script language='javascript'>"
	response.write "alert('������������ؼ��к��зǷ��ַ���������������룡');"
	response.write "location.href='javascript:history.go(-1)';"							
	response.write "</script>"	
	response.end
end if
end if
%>
<!--#include file="conn.asp"-->
<HTML><HEAD>
<TITLE><%=sitename%></TITLE>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<script src="flashin.js" type="text/javascript"></script>
<meta name="description" content="<%=sitename%>">
<meta name="keywords" content="<%=sitename%>">
<link href="../css.css" rel="stylesheet" type="text/css">
<style type="text/css">
<!--
.style6 {color: #999999}
.style7 {	color: #000000;
	font-weight: bold;
	font-size: 14px;
}
body,td,th {
	font-size: 12px;
}
.style8 {
	font-size: 12px;
	color: #333333;
	font-family: Arial, Helvetica, sans-serif;
	line-height: 25px;
	text-decoration: none;
}
.style9 {	font-size: 12px;
	color: #333333;
	font-family: Arial, Helvetica, sans-serif;
	line-height: 20px;
	font-weight: normal;
	text-decoration: none;
}
.style14 {
	color: #333333;
	font-family: Arial, Helvetica, sans-serif;
	font-size: 12px;
	font-weight: normal;
	text-decoration: none;
	border: 1px solid;
}
.style14 {
	font-size: 12px;
	color: #333333;
	font-family: Arial, Helvetica, sans-serif;
	line-height: 25px;
	font-weight: normal;
	text-decoration: none;
	border: 1px solid;
}
.style19 {font-size: 13px;
	color: #FFFFFF;
	font-family: Arial, Helvetica, sans-serif;
	font-weight: bold;
}
.style19 {font-size: 13px;
	color: #FFFFFF;
	font-family: Arial, Helvetica, sans-serif;
	font-weight: bold;
}
.style20 {	font-size: 12px;
	font-family: Arial, Helvetica, sans-serif;
}
.style12 {	font-size: 12px;
	color: #FFFFFF;
	font-family: Arial, Helvetica, sans-serif;
	line-height: 18px;
	font-weight: normal;
	text-decoration: none;
}
.style21 {	color: #666666;
	font-family: Arial, Helvetica, sans-serif;
	font-size: 12px;
	line-height: 25px;
	font-weight: normal;
	text-decoration: none;
}
.style22 {font-size: 12px; color: #333333; }
.style23 {	color: #0066CC;
	font-weight: normal;
	font-family: Arial, Helvetica, sans-serif;
	font-size: 12px;
	text-decoration: none;
}
-->
</style>
<link href="../image/fish.css" rel="stylesheet" type="text/css">
<style type="text/css">
<!--
.style24 {	font-size: 12px;
	font-family: Arial, Helvetica, sans-serif;
	color: #333333;
	text-decoration: none;
}
a {
	font-family: Arial, Helvetica, sans-serif;
	font-size: 12px;
	color: #333333;
}
a:visited {
	color: #333333;
	text-decoration: none;
}
a:hover {
	color: #333333;
	text-decoration: underline;
}
a:active {
	color: #333333;
	text-decoration: none;
}
a:link {
	color: #333333;
	text-decoration: none;
}
.style16 {color: #333333}
.style17 {	color: #4D679B;
	font-family: Arial, Helvetica, sans-serif;
	font-size: 12px;
	font-weight: normal;
	text-decoration: none;
}
-->
</style>
</HEAD>
<center>
  <table width="770" border="0" cellpadding="0" cellspacing="0">
    <tr>
      <td><table width="100%"  border="0" align="center" cellpadding="0" cellspacing="0">
        <tr>
          <td><table width="99%" height="36"  border="0" align="center" cellpadding="0" cellspacing="0">
              <tr>
                <td><%
set rs=Server.CreateObject("ADODB.RecordSet")
sql="select * from Feedback where online='1' "
keywords=request("keywords")
if keywords<>"" then sql=sql+ " and Comments like '%"&keywords&"%' "
sql=sql + "order by top desc,Postdate desc"
rs.open sql,conn,1
if not (rs.eof and rs.bof) then			'���������ʱ������ʾ���ԡ����е�if�뵹����6�е�end if���Ӧ

if pages=0 or pages="" then pages=5		'ÿҳ��������
rs.pageSize = pages	'ÿҳ��¼��
allPages = rs.pageCount	'��ҳ��
page = Request("page")	'�������ȡ�õ�ǰҳ
'if�ǻ����ĳ�����

If not isNumeric(page) then page=1

if isEmpty(page) or Cint(page) < 1 then
page = 1
elseif Cint(page) >= allPages then
page = allPages 
end if
rs.AbsolutePage = page			'ת��ĳҳͷ��
	Do While Not rs.eof and pages>0
	UserName=rs("UserName")		'�û���
	pic=rs("pic")			'ͷ��
	face=rs("face")			'����
	Comments=rs("Comments")		'����
	bad1=split(bad,"/")		'�����໰
	for t=0 to ubound(bad1)
	Comments=replace(Comments,bad1(t),"***")
	next
	Replay=rs("Replay")		'�ظ�
	Usermail=rs("Usermail")		'�ʼ�
	url=rs("url")			'��ҳ
	I=I+1				'���
	temp=RS.RecordCount-(page-1)*rs.pageSize-I+1
	%>
                    <table width="99%"  border="0" cellspacing="0" cellpadding="0">
                      <tr>
                        <td>&nbsp;</td>
                      </tr>
                  </table></td>
              </tr>
            </table>
              <table width="99%" border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#EFEFEF" class="style14" style="word-break:break-all">
                <tr>
                  <td valign="top" width="25%" bgcolor="#f7f7f7" rowspan="2" align=center><table width="100%"  border="0" cellpadding="0" cellspacing="0">
                    <tr>
                      <td height="215" valign="middle" bgcolor="#FFFFFF"><table width=85% border=0 align="center" class="style8">
                        <tr>
                          <td align=center><img src=images/face/pic<%=pic%>.gif border=0></td>
                        </tr>
                        <tr>
                          <td>������<%=UserName%><br>
      ���ԣ�<%=left(rs("ip"),(len(rs("ip"))-2))+"**"%><br>
      �ʼ���<a href=mailto:<%=Usermail%>><img src=images/mail.gif border=0 alt="<%=Usermail%>"></a><br>
      ��ҳ��<a href="<%=URL%>" target='_blank'><img src=images/home.gif border=0 alt="<%=URL%>"></a></td>
                        </tr>
                      </table></td>
                    </tr>
                  </table></td>
                  <td width="75%" height="20" bgcolor="#EFEFEF"><%if rs("top")<>"1" then response.write "[NO."&temp&"]"%>
                      <img border=0 src=images/face/face<%=face%>.gif> �����ڣ�<%=cstr(rs("Postdate"))%></td>
                </tr>
                <tr>
                  <td width='75%' height=120 valign="top" bgcolor="#FFFFFF" >
                    <%
	'�Ƿ��������������е�html�ַ�
	if html=0 then
	response.write replace(server.htmlencode(Comments),vbCRLF,"<BR>")
	else
	response.write replace(Comments,vbCRLF,"<BR>")
	end if
	%>
                    <br>
                    <br>
                    <%if rs("Replay")<>"" then%>
                    <table width="90%" border="0" align="center" cellpadding="3" cellspacing="1" class="style8">
                      <tr>
                        <td valign="top" bgcolor="#f7f7f7"> <font color=<%=huifucolor%>><%=huifutishi%>��<br>
                              <%=Replay%></font> </td>
                      </tr>
                    </table>
                    <br>
                    <%end if%>                  </td>
                </tr>
              </table>
              <table width="99%"  border="0" align="center" cellpadding="0" cellspacing="0">
                <tr>
                  <td class="style20"><%
pages = pages - 1
rs.movenext
if rs.eof then exit do
loop

response.write "<table border=0 align=center class=style1><tr><td><form action='index.asp' method='post'>�ܼ�����"&RS.RecordCount&"�� "
if page = 1 then
response.write "<font color=darkgray>��ҳ ǰҳ</font>"
else
response.write "<a href=index.asp?page=1>��ҳ</a> <a href=index.asp?keywords="&keywords&"&page="&page-1&">ǰҳ</a>"
end if
if page = allpages then
response.write "<font color=darkgray> ��ҳ ĩҳ</font>"
else
response.write " <a href=index.asp?keywords="&keywords&"&page="&page+1&">��ҳ</a> <a href=index.asp?keywords="&keywords&"&page="&allpages&">ĩҳ</a>"
end if
response.write " ��"&page&"ҳ ��"&allpages&"ҳ &nbsp; ת���� "
response.write "<select name='page'>"
for i=1 to allpages
response.write "<option value="&i&">"&i&"</option>"
next
response.write "</select> ҳ <input type=submit name=go value='Go'><input type=hidden name=keywords value='"&keywords&"'></form></td><td align=right>"
response.write "<form action='index.asp' method='post'><input title='�����ʲô����' type=text name=keywords value='"&keywords&"' size=10 maxlength=20><input type=submit value='����' title='������'></form></td></tr></table>"


else
response.write "<table cellSpacing=0 cellPadding=0 width=680 align=center bgColor=#FFFFFF border=0><TR><TD height=120 align=center>"
if keywords="" then response.write "��ʱû������" else response.write "��Ǹ��û���ҵ���Ҫ�鵽������<br><br><a href='javascript:history.go(-1)'>������һҳ</a>" end if
response.write "</TD></TR></TABLE>"
end if
%></td>
                </tr>
            </table>
            </td>
        </tr>
      </table></td>
    </tr>
  </table>
</center>
</html>