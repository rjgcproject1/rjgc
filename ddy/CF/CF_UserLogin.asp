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
<script language=JavaScript>
function ccimgchange()
{
 document.getElementById("ccimg").src = 'CF_CheckCode.asp?a=' + Math.random();
}

function userlogincheck()
{
  if(window.document.form1.username.value=="")
  {
     alert("�������û���!");
	 window.document.form1.username.focus();
	 return false;
  }
  if(window.document.form1.pwd.value=="")
  {
     alert("����������!");
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

<TABLE class="tb_3" style="width:213px;margin-top:3px;margin-bottom:0px">
      <TR class="tr_1"> 
        <TD colspan="2" class="td_1">�û���¼</TD>
      </TR>
	  
  <form name="form1" method="post" action="Index.asp?Action=userloginsave" onsubmit="return userlogincheck()">
<%
If Session("CFCountUser")="" Then
%>
      <TR> 
        <TD class="td_3" style="padding-top:5px">&nbsp;�û�����</TD>
        <TD style="padding-top:5px"><INPUT name=username id=username value="<%=ChkStr(Request("UserName"),1)%>" size=15></TD>
      </TR>
      <TR> 
        <TD class="td_3" style="padding-top:5px">&nbsp;��&nbsp;&nbsp;&nbsp;&nbsp;�룺</TD>
        <TD style="padding-top:5px"><INPUT id=pwd type=password size=15 name=pwd></TD>
      </TR>
      <TR> 
        <TD class="td_3" style="padding-top:5px">&nbsp;Cookie��</TD>
        <TD style="padding-top:5px"><select name="cookies_time" id="cookies_time" style="width:115px;">
            <option value="60" selected>������</option>
            <option value="1440">����һ��</option>
            <option value="43200">����һ����</option>
            <option value="5256000">���ñ���</option>
          </select></TD>
      </TR>
      <TR> 
        <TD class="td_3" style="padding-top:3px">&nbsp;��֤�룺</TD>
        <TD style="padding-top:3px"><INPUT id=checkcode type=text size=6 name=checkcode><a href="javascript:ccimgchange()" title="�����������һ����"><IMG id="ccimg" height=14 src="CF_CheckCode.asp" border="0"></a></TD>
      </TR>
      <TR> 
        <TD class="td_3" style="padding-top:3px">&nbsp;<a href="PwdRecover.asp"><font color="#0066FF">�����һ�</font></a></TD>
        <TD><INPUT type="submit" value="��&nbsp;¼" name="submit" class="button">

 &nbsp;<a href="Reg.asp">ע&nbsp;&nbsp;��</a></TD>
      </TR>
      <%Else%>
      <TR> 
        <TD colspan="2" style="text-align:center;height:142px;vertical-align:top">
		<INPUT onclick="location.href='Manage.asp';" type='button' value='�������' name='submit' class='button' style="margin-top:10px">
		<br />
		<INPUT onclick="location.href='Manage.asp?action=userlogout';" type='button' value='�˳�����' name='submit' class='button'  style="margin-top:10px">
		</TD>
      </TR>

      <%End if%>
  </FORM>
</TABLE>
