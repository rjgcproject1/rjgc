<%
'�˷���û�������
'����QQ��178575
'����EMail��yliangcf@163.com
'������վ��http://www.qqcf.com
'��ϸ��飺http://www.qqcf.com/?action=list&list=cfcount
'�����г���������ʾ����װ��ʾ��ʹ�����ѽ�����°汾���ص�����
'��Ϊ��Щ���ݿ���ʱ�����£���û�з��ڳ�������Լ�����վ�ϲ鿴
'�������汾����ʾ
%>
<!--#Include File="Conn.asp"-->
<!--#Include File="CF_MyFunCtion.asp"-->
<!--#Include File="CF_Md5.asp"-->
<!--#Include File="CF_Getcookie.asp"-->

<%
Action=Request("Action")
%>
<!--#Include File="CF_Do.asp"-->
<%
If Session("CFCountAdmin")="ok" Then Response.Redirect "CF_Admin_Manage.asp"
%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<HTML>
<HEAD>
<TITLE><%=RsSet("SysTitle")%></TITLE>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<!--#include file="CF_Style.asp"-->
<script language=JavaScript>

 function ccimgchange()
{
 document.getElementById("ccimg").src = 'CF_CheckCode.asp?a=' + Math.random();
}

function adminlogincheck()
{
  if(window.document.form1.admin.value=="")
  {
     alert("���������Ա����!");
	 window.document.form1.admin.focus();
	 return false;
  }
  if(window.document.form1.pwd.value=="")
  {
     alert("���������Ա����!");
	 window.document.form1.pwd.focus();
	 return false;
  }
 if(window.document.form1.checkcode.value=="")
  {
     alert("��������λ���ֵ���֤��!");
	 window.document.form1.checkcode.focus();
	 return false;
  }
  return true;
}

</script>
</HEAD>
<BODY>
<!--#include file="Top.asp"-->

<table class="tb_2">
          <tr class="tr_1"> 
            <td colspan="2">��������Ա���</td>
          </tr>
          <form name="form1" method="post" action="?Action=adminloginsave" onSubmit="return adminlogincheck()">
            <tr> 
              <td width="44%" class="td_3">����Ա���ƣ�</td>
              <td><input name="admin" type="text" id="admin3" size="15"></td>
            </tr>
            <tr> 
              <td class="td_3">����Ա���룺</td>
              <td><input name="pwd" type="password" id="pwd" size="15"></td>
            </tr>
            <tr> 
              <td class="td_3">Cookie��</td>
              <td><select name="cookies_time" id="cookies_time" style="width:115px;">
                  <option value="60" selected>������</option>
                  <option value="1440">����һ��</option>
                  <option value="43200">����һ����</option>
                  <option value="5256000">���ñ���</option>
                </select> </td>
            </tr>
              <TR>
                <TD class="td_3">��֤�룺</TD>
                <TD><INPUT id=checkcode type=text size=6 name=checkcode><a href="javascript:ccimgchange()" title="�����������һ����"><IMG id="ccimg" height=14 src="CF_CheckCode.asp" border="0"></a></TD>
              </TR>
            <tr> 
              <td></td> 
			  <td>  

                  <input type="submit" name="Submit" value="ȷ��">
                  �� 
                  <input type="reset" name="Submit523" value="ȡ��">
&nbsp;&nbsp;&nbsp;&nbsp;<a href="Index.asp">������ҳ</a>
                </td>
            </tr>
          </form>
        </table>   

<!--#include file="Bottom.asp"-->

</BODY></HTML>

<%Call ConnClose()%>