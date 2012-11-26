<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%
Const dbdns = "./"
Const Htmldns = "./"
Const Base_Target = ""
%>

<!--#include file="AppCode/Conn.asp"-->
<!--#include file="AppCode/fun/function.asp"-->
<!--#include file="AppCode/Class/Ok3w_Article.asp"-->
<!--#include file="vbs.asp"-->

<%
id=myCdbl(Request.QueryString("id"))
Set Article = New Ok3w_Article
Call Article.Load(Id)
If Article.IsPass=0 Then Call Page_Err("�����Ѿ��ر�")
If Article.IsDelete=1 Then Call Page_Err("�����Ѿ�ɾ��")
ClassID = Article.ClassID

Sql="select * from Ok3w_Class where ID=" & ClassID
Rs.Open Sql,Conn,0,1
SortPath = Rs("SortPath")
SortName = Rs("SortName")
Rs.Close

%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<!--<title>��ӭ�����й����ʴ�ѧ���人���о�����ҵ��Ϣ��</title>-->
<title><%=Article.Title%></title>
<meta name="keywords" content="<%=Application("Ok3w_SiteKeyWords")%>">
<meta name="description" content="<%=Application("Ok3w_SiteDescription")%>">

<link rel="stylesheet" href="css/reset.css" />
<link rel="stylesheet" href="css/text.css" />
<link rel="stylesheet" href="css/960_24_col.css" />
<link rel="stylesheet" href="css/index.css" />
<link href="css/menu.css" rel="stylesheet" type="text/css" />
<script type="text/javascript" src="js/jquery-1.4.2.min.js"></script>
<script type="text/javascript" src="js/menu.js" ></script>
</head>

<body>
<div class="container_24 all">
			
		<div class="header">
            <div class="grid_24">
                <img src="img/banner.gif" align="banner" />
            </div>
            <div class="grid_24 nav">                
                    <%Call Ok3w_Article_Class_NavMenus()%>
            </div>
		</div>
		<div class="grid_24 content">
		<div class="grid_6 alpha content_left">
        	<!--��Ҫ��ӵ����ݣ���ർ�����˵�-->
            <div class="menu">
            	<div class="text">
                        <%=SortName%>
                </div>
            </div>
            <div class="menudata3">
            	<!--������ID�г���ID�µ���������-->
                <%Call Ok3w_Article_Class_NavLists(ClassID)%>
            	<!--<img src="img/left_img/download2.png"/>-->
            </div>
             
        </div>
        
        <div class="grid_18 omega content_right">
        	<!--��Ҫ��ӵ����ݣ��Ҳ�������-->
            <!--�Ҳ�������ͷ��λ����Ϣ-->
            <div class="content_right_header"> 	
				��ǰλ�ã�    <a href="index.asp">��ҳ</a> &gt;&gt; 
            <%Call Ok3w_Article_Class_Nav(SortPath)%>       
            
            </div>
            
            <!--��������-->
            <div class="content_right_body">
                  <div class="showtitle1"> <%=Article.Title%></div>
				  <div class="showtitle2">
					<%=Format_Time(Article.AddTime,1)%> ��Դ��<%=Article.ComeFrom%> �����<span id="News_Hits"><%=Article.Hits%></span>��
				  </div>
				
				  <div class="showtitle3">
                  		<%If Article.IsPic=1 Then%>
                        <%str = "<img src='" & Article.PicFile & "'/>"%>
                        <%Response.Write(str)%>
                        <%Else%>
						<%End If%>
					    <%If Article.Description="" Then%>
						<%Else%>
						<%End If%>
						<%If Article.vUserGroupID=0 And Article.vUserJifen=0 Then%>
						<%Call OutThisPageContent(Article.ID,Article.Content,"html")%>
						<%Else%>
						<%End If%>
						<%If Article.IsUserAdd=1 Then%>
						<%End If%>
				  </div>
            </div>
        </div>
		<div style="clear:both"></div>
		</div>
        <div class="grid_24 footer">
           <div class="right">
				��Ȩ���У��й����ʴ�ѧ���人����Ϣ����ѧԺ�������ϵ
    		</div>
    	</div>
</div>

</body>
</html>
