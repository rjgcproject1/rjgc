<%
Class FunCls
	Dim AllExtName '������������
	Dim MaxFileSize '������������
	Dim DownTimeout '��ʱ����
	Private Is_Admin'�Ƿ��½
	'===============================================
	'�������¼�
	'===============================================
	Private Sub Class_Initialize()
		On Error Resume Next
		DownTimeout = 64 '��ʱ����
		MaxFileSize = 0'-- ���ش�С����
		AllExtName = "rm|swf"'-- ������������
	End Sub
	'===============================================
	'�ر����¼�
	'===============================================
	Private Sub Class_Terminate()
		'-- Class_Terminate
	End Sub
	'===============================================
	'-- ��ʱ����
	'===============================================
	Public Property Let CjTimeout(ByVal NewValue)
		DownTimeout = NewValue
	End Property
	'===============================================
	'-- ������������
	'===============================================
	Public Property Let DownExtName(ByVal NewValue)
		AllExtName = NewValue
	End Property
	'===============================================
	'-- ���ش�С����
	'===============================================
	Public Property Let MaxSize(ByVal NewValue)
		MaxFileSize = NewValue * 1024
	End Property
	'===============================================
	'����Ա��֤
	'===============================================
	Public Function IsAdmin()
	IsAdmin=Is_Admin
	End FunCtion
	'===============================================
	'����Ա��֤
	'===============================================
	Sub Admin()
		If Session("Ok3w_Caiji")="" Then
			Response.Write("�㻹û�е�½�����ǵ�½�Ѿ���ʱ��")
			Response.End()
		End If
		
		Is_Admin = True
	End Sub	
	'===============================================
	'��������G()
	'��  �ã�'ȡ��Request.Querystring �� Request.Form ��ֵ
	'===============================================
	Public Function G(Str)
	 G = Replace(Replace(Request(Str), "'", ""), """", "")
	End Function
	'===============================================
	'��������ChkNumeric()
	'��  �ã�' ת����LONG ������̬�� 
	'===============================================
	Public Function ChkNumeric(ByVal CHECK_ID)
		If CHECK_ID <> "" And IsNumeric(CHECK_ID) Then
			CHECK_ID = CLng(CHECK_ID)
		Else
			CHECK_ID = 0
		End If
		ChkNumeric = CHECK_ID
	End Function
	'===============================================
	'��������GetConfig
	'��  �ã���ȡϵͳ������Ϣ
	'��  ����  ConfigField��Ӧ���ֶ�����
	'����ֵ����Ӧ�ֶε�ֵ
	'===============================================
	Public Function GetConfig(ByVal ConfigField)
	    IF Application(Site & "SiteConfig_" & ConfigField)="" Then
				   Dim ConfigRS:Set ConfigRS = Server.CreateObject("Adodb.Recordset")
				   On Error Resume Next
				   ConfigRS.Open ("Select * From SK_Config"), ConnItem, 1, 1
				   GetConfig = ConfigRS(ConfigField)
				   If Err.Number <> 0 Then GetConfig = "":Err.clear
				   ConfigRS.Close:Set ConfigRS = Nothing
				   Application(Site & "SiteConfig_" & ConfigField)=GetConfig
		Else
				 GetConfig=Application(Site & "SiteConfig_" & ConfigField)
		End If
	End Function
	'===============================================
	'��������GetItemConfig
	'��  �ã���ȡ�ɼ�����������Ϣ
	'��  ����ConfigField��Ӧ���ֶ�����,CJID�������õ�ID��
	'===============================================
	Public Function GetItemConfig(ByVal ConfigField,CJID)
		   Dim ConfigRS:Set ConfigRS = Server.CreateObject("Adodb.Recordset")
		   On Error Resume Next
		   ConfigRS.Open ("Select * From SK_Cj where ID="& CJID), ConnItem, 1, 1
		   GetItemConfig = ConfigRS(ConfigField)
		   If Err.Number <> 0 Then GetItemConfig = "":Err.clear
		   ConfigRS.Close:Set ConfigRS = Nothing
	End Function
	'===============================================
	'��������GetHttpPage
	'��  �ã���ȡ��ҳԴ��
	'��  ����HttpUrl ------��ҳ��ַ,Cset ����
	'===============================================
	Function GetHttpPage(ByVal URL, ByVal Cset)
	Dim BlockStartTime
	On Error Resume Next
	Dim Http
	 If IsNull(URL)=True Or Len(URL)<18 Or URL="$False$" Then
		  GetHttpPage="$False$"
		  Exit Function
	   End If
	   BlockStartTime = Timer()
	   Set Http=server.createobject("MSXML2.XMLHTTP")
	   Http.open "GET",URL,False
	   Http.Send()
	  'ѭ���ȴ����ݽ���
	   Dim temp,BlockTimeout 	   
	   BlockTimeout = 64
	   While (http.ReadyState <> 4)
		  ' �ж��Ƿ�鳬ʱ
		   temp = Timer() - BlockStartTime
		   Response.Write(Timer())
		   If (temp > BlockTimeout) Then
			   http.abort
			   Set Http=Nothing 
			   GetHttpPage="$False$"
			   Exit function
			   Response.End
		   End If
		   http.waitForResponse 10000'�ȴ�1000����
	   Wend
	   If Http.Readystate<>4 then
		  Set Http=Nothing 
		  GetHttpPage="$False$"
		  Exit function
	   End if
	   GetHTTPPage=bytesToBSTR(Http.responseBody,Cset)
	   Set Http=Nothing
	   
	   If Err.number<>0 then
		  If IsNull(URL)=True Or Len(URL)<18 Or URL="$False$" Then
		  GetHttpPage="$False$"
		  Exit Function
	   End If
	   Set Http=Nothing
		  Err.Clear
	   End If
	   
	End Function
	'===============================================
	'��������BytesToBstr
	'��  �ã�����ȡ��Դ��ת��Ϊ����
	'��  ����Body ------Ҫת���ı���
	'��  ����Cset ------Ҫת��������
	'===============================================
	Function BytesToBstr(Body,Cset)
	   Dim Objstream
	   Set Objstream = Server.CreateObject("adodb.stream")
	   objstream.Type = 1
	   objstream.Mode =3
	   objstream.Open
	   objstream.Write body
	   objstream.Position = 0
	   objstream.Type = 2
	   objstream.Charset = Cset
	   BytesToBstr = objstream.ReadText 
	   objstream.Close
	   set objstream = nothing
	End Function
	
	'===============================================
	'��������PostHttpPage
	'��  �ã���¼
	'===============================================
	Function PostHttpPage(RefererUrl,PostUrl,PostData) 
		Dim xmlHttp 
		Dim RetStr      
		Set xmlHttp = CreateObject("Msxml2.XMLHTTP")  
		xmlHttp.Open "POST", PostUrl, False
		XmlHTTP.setRequestHeader "Content-Length",Len(PostData) 
		xmlHttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		xmlHttp.setRequestHeader "Referer", RefererUrl
		xmlHttp.Send PostData 
		If Err.Number <> 0 Then 
			Set xmlHttp=Nothing
			PostHttpPage = "$False$"
			Exit Function
		End If
		PostHttpPage=bytesToBSTR(xmlHttp.responseBody,"GB2312")
		Set xmlHttp = nothing
	End Function 

	'===============================================
	'��������UrlEncoding
	'��  �ã�ת������
	'===============================================
	Function UrlEncoding(DataStr)
		Dim StrReturn,Si,ThisChr,InnerCode,Hight8,Low8
		StrReturn = ""
		For Si = 1 To Len(DataStr)
			ThisChr = Mid(DataStr,Si,1)
			If Abs(Asc(ThisChr)) < &HFF Then
				StrReturn = StrReturn & ThisChr
			Else
				InnerCode = Asc(ThisChr)
				If InnerCode < 0 Then
				   InnerCode = InnerCode + &H10000
				End If
				Hight8 = (InnerCode  And &HFF00)\ &HFF
				Low8 = InnerCode And &HFF
				StrReturn = StrReturn & "%" & Hex(Hight8) &  "%" & Hex(Low8)
			End If
		Next
		UrlEncoding = StrReturn
	End Function
	'===============================================
	'��������GetBody
	'��  �ã���ȡ�̶����ַ���
	'��  ����strHTML   ----ԭ�ַ���
	'��  ��: start ------ ��ʼ�ַ���
	'��  ��: Over ------ �����ַ���
	'��  ����IncluL ------�Ƿ����StartStr
	'��  ����IncluR ------�Ƿ����OverStr
	'===============================================
	Public Function GetBody(ByVal strHTML, ByVal Start, ByVal Over,IncluL,IncluR)
		Dim SS
		Dim Match
		Dim TempStr
		Dim strPattern
		Dim s,o  
		If IsNull(Start)=True Then GetBody="$False$" : Exit Function
		Start=ReplaceTrim(Start) : Over=ReplaceTrim(Over) : strHTML=strHTML
		s=Len(start) : o=Len(Over)

		If s = 0 Or o = 0  Then GetBody="$False$" : Exit Function
		strPattern = "(" & CorrectPattern(start) & ")(.+?)(" & CorrectPattern(Over) & ")"
		On Error Resume Next
		Dim re
		Set re = New RegExp
		re.IgnoreCase = False
		re.Global = False
		re.Pattern = strPattern
		Set SS = re.Execute(strHTML)
		For Each Match In SS
			TempStr = Match.Value
		Next
		If TempStr="" Then'���ַ���,����������
		   GetBody="$False$"
		   Exit Function
	    End If
	   
		If IncluL=False then
			TempStr=Right(TempStr,Len(TempStr) -S)
	    End if
	    If IncluR=False then
			TempStr=Left(TempStr,Len(TempStr) - O)
	    End if	
		If Err.number<>0 then  '����,����������
		   GetBody="$False$"
		   Exit Function
		End If
		Set SS = Nothing
		Set re = Nothing
		GetBody = TempStr
		Exit Function
	End Function
	'===============================================
	'��������GetArray
	'��  �ã���ȡ���ӵ�ַ����$Array$�ָ�
	'��  ����ConStr ------��ȡ��ַ��ԭ�ַ�
	'��  ����StartStr ------��ʼ�ַ���
	'��  ����OverStr ------�����ַ���
	'��  ����IncluL ------�Ƿ����StartStr
	'��  ����IncluR ------�Ƿ����OverStr
	'===============================================
	Function GetArray(Byval ConStr,StartStr,OverStr,IncluL,IncluR)
	   Dim TempStr,TempStr2,objRegExp,Matches,Match,Templisturl,TempStr_i
	   Dim s,o 
	   On Error Resume Next
	   If ConStr="$False$" or ConStr="" Or IsNull(ConStr)=True or StartStr="" Or OverStr="" or  IsNull(StartStr)=True Or IsNull(OverStr)=True Then
		  GetArray="$False$"
		  Exit Function
	   End If
	   StartStr=ReplaceTrim(StartStr) : OverStr=ReplaceTrim(OverStr) : ConStr=ConStr
	   s=Len(StartStr) : o=Len(OverStr)
	   TempStr=""
	   Set objRegExp = New Regexp 
	   objRegExp.IgnoreCase = True 
	   objRegExp.Global = True
	   objRegExp.Pattern = "("&CorrectPattern(StartStr)&").+?("&CorrectPattern(OverStr)&")"
	   Set Matches =objRegExp.Execute(ConStr) 
	   For Each Match in Matches
		  'If Templisturl =Match.Value then
		  'Else
		  		TempStr_i=Match.Value
		  	  	If IncluL=False then
					TempStr_i=Right(TempStr_i,Len(TempStr_i) -S)
				End if
				If IncluR=False then
					TempStr_i=Left(TempStr_i,Len(TempStr_i) - O)
				End if	
			  	TempStr=TempStr & "$Array$" & TempStr_i
		  '	  	Templisturl = Match.Value
		  'End if
	   Next 
	   Set Matches=nothing
	
	   If TempStr="" Then
		  GetArray="$False$"
		  Exit Function
	   End If
	   TempStr=Right(TempStr,Len(TempStr)-7)
	   Set objRegExp=nothing
	   Set Matches=nothing

	   If TempStr="" then
		  GetArray="$False$"
	   Else
		  GetArray=TempStr
	   End if
	End Function
	'===============================================
	'��������ReplaceSaveRemoteFile
	'��  �ã��滻������Զ��ͼƬ
	'��  ����ConStr ------ Ҫ�滻���ַ���
	'��  ����SaveTf ------ �Ƿ񱣴��ļ���False�����棬True����
	'��  ��: TistUrl------ ��ǰ��ҳ��ַ
	'===============================================
	Function ReplaceSaveRemoteFile(ConStr,strInstallDir,strChannelDir,SaveTf,TistUrl)
	   If ConStr="$False$" or ConStr="" or strChannelDir="" Then
		  ReplaceSaveRemoteFile=ConStr
		  Exit Function
	   End If
	   Dim TempStr,TempStr2,TempStr3,Re,Matches,Match,Tempi,TempArray,TempArray2
	
	   Set Re = New Regexp 
	   Re.IgnoreCase = True 
	   Re.Global = True
	   Re.Pattern ="<img.+?[^\>]>"
	   Set Matches =Re.Execute(ConStr) 
	   For Each Match in Matches
		  If TempStr<>"" then 
			 TempStr=TempStr & "$Array$" & Match.Value
		  Else
			 TempStr=Match.Value
		  End if
	   Next
	   If TempStr<>"" Then
		  TempArray=Split(TempStr,"$Array$")
		  TempStr=""
		  For Tempi=0 To Ubound(TempArray)
			 Re.Pattern ="src\s*=\s*.+?\.(gif|jpg|bmp|jpeg|psd|png|svg|dxf|wmf|tiff)"
			 Set Matches =Re.Execute(TempArray(Tempi)) 
			 For Each Match in Matches
				If TempStr<>"" then 
				   TempStr=TempStr & "$Array$" & Match.Value
				Else
				   TempStr=Match.Value
				End if
			 Next
		  Next
	   End if
	   If TempStr<>"" Then
		  IncludePic=1'ͼƬ����
		  Re.Pattern ="src\s*=\s*"
		  TempStr=Re.Replace(TempStr,"")
	   End If
	   Set Matches=nothing
	   Set Re=nothing
	   If TempStr="" or IsNull(TempStr)=True Then
		  ReplaceSaveRemoteFile=ConStr
		  Exit function
	   End if
	   TempStr=Replace(TempStr,"""","")
	   TempStr=Replace(TempStr,"'","")
	   TempStr=Replace(TempStr," ","")
	   
	   Dim RemoteFileurl,SavePath,PathTemp,DtNow,strFileName,strFileType,ArrSaveFileName,RanNum,Arr_Path
	   DtNow=Now()
	   If SaveTf=True then
	 '***********************************
		  SavePath= strChannelDir & year(DtNow) & right("0" & month(DtNow),2) & "/"
		  response.write "����·����" & savepath & "<br>"
		  Arr_Path=Split(SavePath,"/")
		  PathTemp=""
		  For Tempi=0 To Ubound(Arr_Path)
			 If Tempi=0 Then
				PathTemp=Arr_Path(0) & "/"
			 ElseIf Tempi=Ubound(Arr_Path) Then
				Exit For
			 Else
				PathTemp=PathTemp & Arr_Path(Tempi) & "/"
			 End If
			 If CheckDir(PathTemp)=False Then
				If MakeNewsDir(PathTemp)=False Then
				   SaveTf=False
				   Exit For
				End If
			 End If
		  Next
	   End If
	
	   'ȥ���ظ�ͼƬ��ʼ
	   TempArray=Split(TempStr,"$Array$")
	   TempStr=""
	   For Tempi=0 To Ubound(TempArray)
		  If Instr(Lcase(TempStr),Lcase(TempArray(Tempi)))<1 Then
			 TempStr=TempStr & "$Array$" & TempArray(Tempi)
		  End If
	   Next
	   TempStr=Right(TempStr,Len(TempStr)-7)
	   TempArray=Split(TempStr,"$Array$")
	   'ȥ���ظ�ͼƬ����
	
	   'ת�����ͼƬ��ַ��ʼ
	   TempStr=""
	   For Tempi=0 To Ubound(TempArray)
		  TempStr=TempStr & "$Array$" & DefiniteUrl(TempArray(Tempi),TistUrl)
	   Next
	   TempStr=Right(TempStr,Len(TempStr)-7)
	   TempStr=Replace(TempStr,Chr(0),"")
	   TempArray2=Split(TempStr,"$Array$")
	   TempStr=""
	   'ת�����ͼƬ��ַ����
		'ͼƬ�滻/����
	   Set Re = New Regexp
	   Re.IgnoreCase = True 
	   Re.Global = True
	   For Tempi=0 To Ubound(TempArray2)
		  RemoteFileUrl=TempArray2(Tempi)
		  If RemoteFileUrl<>"$False$" And SaveTf=True Then'����ͼƬ
			 ArrSaveFileName = Split(RemoteFileurl,".")
		 strFileType=Lcase(ArrSaveFileName(Ubound(ArrSaveFileName)))'�ļ�����
			 If strFileType="asp" or strFileType="asa" or strFileType="aspx" or strFileType="cer" or strFileType="cdx" or strFileType="exe" or strFileType="rar" or strFileType="zip" then
				UploadFiles=""
				ReplaceSaveRemoteFile=ConStr
				Exit Function
			 End If
	
			 Randomize
			 RanNum=Int(900*Rnd)+100
		 strFileName = year(DtNow) & right("0" & month(DtNow),2) & right("0" & day(DtNow),2) & right("0" & hour(DtNow),2) & right("0" & minute(DtNow),2) & right("0" & second(DtNow),2) & ranNum & "." & strFileType
			 Re.Pattern =TempArray(Tempi)
			 
		 If SaveRemoteFile(SavePath & strFileName,RemoteFileUrl)=True Then
		 		ImagesNumAll = ImagesNumAll + 1
				
				PathTemp=SavePath & strFileName 
				ConStr=Re.Replace(ConStr,PathTemp)
				Re.Pattern=strInstallDir & strChannelDir 
				UploadFiles=UploadFiles & "|" & Re.Replace(SavePath &strFileName,"")
				Response.Flush()
				response.write " &nbsp;&nbsp;&nbsp;ͼƬ�����ַ��" & PathTemp & "<br>"
				if Thumb_WaterMark=1 then call SKThumb.AddWaterMark(PathTemp)'ˮӡ
			 Else
				PathTemp=RemoteFileUrl
				ConStr=Re.Replace(ConStr,PathTemp)
				'UploadFiles=UploadFiles & "|" & RemoteFileUrl
			 End If
		  ElseIf RemoteFileurl<>"$False$" and SaveTf=False Then'������ͼƬ
			 Re.Pattern =TempArray(Tempi)
			 ConStr=Re.Replace(ConStr,RemoteFileUrl)
			 UploadFiles=UploadFiles & "|" & RemoteFileUrl
		  End If
	   Next   
	   Set Re=nothing
	   If UploadFiles<>"" Then
		  UploadFiles=Right(UploadFiles,Len(UploadFiles)-1)
	   End If
	   ReplaceSaveRemoteFile=ConStr
	End function
	'===============================================
	'��������ReSaveRemoteFile
	'��  �ã������ļ������滻
	'��  ����Str   ----ԭ�ַ���
	'��  ����url   ----��Ȼ��վURL
	'��  ����Dir -----����Ŀ¼
	'��  ����InSave ------�Ƿ񱣴�,True,False
	'����ֵ����ʽ��ȡ����ַ���
	'===============================================
	Public Function ReSaveRemoteFile(ByVal str, ByVal URL, ByVal Dir,InSave)
		Dim s_Content
		Dim re
		Dim ContentFile, ContentFileUrl
		Dim strTempUrl,strFileUrl,DirTemp,PathTemp,FileTemp,Tempi,TempUrlArray,Arr_Path
		s_Content = str
		On Error Resume Next
		Set re = New RegExp
		re.IgnoreCase = True
		re.Global = True
		re.Pattern = "((src=|href=)((\S)+[.]{1}(" & AllExtName & ")))"
		Set ContentFile = re.Execute(s_Content)
		Dim sContentUrl(), n, i, bRepeat
		n = 0
		For Each ContentFileUrl In ContentFile
			strFileUrl = Replace(Replace(Replace(Replace(ContentFileUrl.Value, "src=", "", 1, -1, 1), "href=", "", 1, -1, 1), "'", ""), Chr(34), "")
			If n = 0 Then
				n = n + 1
				ReDim sContentUrl(n)
				sContentUrl(n) = strFileUrl
			Else
				bRepeat = False
				For i = 1 To UBound(sContentUrl)
					If UCase(strFileUrl) = UCase(sContentUrl(i)) Then
						bRepeat = True
						Exit For
					End If
				Next
				If bRepeat = False Then
					n = n + 1
					ReDim Preserve sContentUrl(n)
					sContentUrl(n) = strFileUrl
				End If
			End If
		Next
		If n = 0 Then
			ReSaveRemoteFile = s_Content
			Exit Function
		End If
		For i = 1 To n 
			strTempUrl = sContentUrl(i) : strTempUrl = FormatRemoteUrl(strTempUrl,URL)'�õ��ļ���ַ
			IF InSave=True then
				Arr_Path=Split(Dir,"/")
				'----------��Ŀ¼-----------------------
				  For Tempi=0 To Ubound(Arr_Path)
					 If Tempi=0 Then
						PathTemp=Arr_Path(0) & "/"
					 ElseIf Tempi=Ubound(Arr_Path) Then
						Exit For
					 Else
						PathTemp=PathTemp & Arr_Path(Tempi) & "/"
					 End If
					 If CheckDir(PathTemp)=False Then
						If MakeNewsDir(PathTemp)=False Then
						   SaveTf=False
						   Exit For
						End If
					 End If
				  Next
				 '------------------------------------------------------
				TempUrlArray=Split(strTempUrl,"/")
				'----------����ļ��Ƿ����.������ڻ��ļ���------------------
				Do while True 
					FileTemp=Dir &  MakeRandom(5) & TempUrlArray(Ubound(TempUrlArray))'��������ļ���
					If CheckFile(FileTemp)=False then
						Exit Do
					end if
				loop 
				'-------------------------------------------------------------------

				If SaveRemoteFile(FileTemp,strTempUrl)=True then
					Response.Write FileTemp & "����ɹ�" & "<Br>"
					s_Content = Replace(s_Content,sContentUrl(i),FileTemp, 1, -1, 1)'�滻��ַ	
				Else
					Response.Write FileTemp & "����ʧ��" & "<Br>"
				End if
			Else
				s_Content = Replace(s_Content,sContentUrl(i),strTempUrl, 1, -1, 1)'�滻��ַ		
			End If	
		Next
		Set re = Nothing
		PictureExist = True
		ReSaveRemoteFile = s_Content
		Exit Function
	End Function
	'===============================================
	'��������SaveRemoteFile
	'��  �ã�����Զ�̵��ļ�������
	'��  ����LocalFileName ------ �����ļ���
	'��  ����RemoteFileUrl ------ Զ���ļ�URL
	'===============================================
	Function SaveRemoteFile(LocalFileName,RemoteFileUrl)
		SaveRemoteFile=True
		dim Ads,Retrieval,GetRemoteData	
		On Error Resume Next
		Set Retrieval = Server.CreateObject("Microsoft.XMLHTTP")
		With Retrieval
			.Open "Get", RemoteFileUrl, False, "", ""
			.Send
			If .Readystate<>4 or .Status > 300 then
				SaveRemoteFile=False
				Exit Function
			End If
			GetRemoteData = .ResponseBody
		End With
		Set Retrieval = Nothing

		If MaxFileSize > 0 Then
			If LenB(GetRemoteData) > MaxFileSize Then Exit Function
		End If
		Response.Write(Round(LenB(GetRemoteData)/1024)) & "KB"	
		Set Ads = Server.CreateObject("Adodb.Stream")
		With Ads
			.Type = 1
			.Open
			.Write GetRemoteData
			.SaveToFile server.MapPath(LocalFileName),2
			.Cancel()
			.Close()
		End With
		If Err.number<>0 then
		  SaveRemoteFile=False
		  Exit Function
		  Err.Clear
		End If
		Set Ads=nothing
		
	end Function
	'===============================================
	'��������Sk_GetSaveDir()
	'lx=����
	'��  �ã���ȡ�ļ�����Ŀ¼����
	'===============================================
	Function Sk_GetSaveDir(lx)
		Dim strInstallDir,CacheTemp,rs,strtemp,strChannelDir,Arr_Path,Tempi,PathTemp,SaveTf,TempUrlArray,Ranfilestr
		strInstallDir=trim(request.ServerVariables("SCRIPT_NAME"))
		strInstallDir=left(strInstallDir,instrrev(lcase(strInstallDir),"/")-1)
		set rs = ConnItem.execute("Select top 1 Dir,MaxFileSize,FileExtName,Timeout from SK_Cj where ID=" & lx)
		strtemp = strtemp  & rs("Dir")
		Sk_GetSaveDir = strtemp & SaveFileUrl
		rs.close
		set rs=nothing
	end function
	'===============================================
	'��������Sk_SaveFile()
	'��  ��: Lx=Ƶ��
	'��  ��: FileUrl=Զ���ļ���ַ
	'��  �ã���Ƶ�����ܱ���Զ���ļ��滻��ַ
	'===============================================
	Function Sk_SaveFile(Lx,FileUrl)
		Dim strInstallDir,CacheTemp,rs,strtemp,strChannelDir,Arr_Path,Tempi,PathTemp,SaveTf,TempUrlArray,Ranfilestr,Ranfilestr1
		Dim SqlTemp
		FileUrl=replace(replace(FileUrl,"""","")," ","")
		SqlTemp="Select top 1 Dir,MaxFileSize from SK_Cj where ID="& Lx
		set rs = ConnItem.execute(SqlTemp)
		strtemp = rs("Dir") & SaveFileUrl
			  Arr_Path=Split(strtemp,"/")
			  PathTemp=""
			  For Tempi=0 To Ubound(Arr_Path)
				 If Tempi=0 Then
					PathTemp=Arr_Path(0) & "/"
				 ElseIf Tempi=Ubound(Arr_Path) Then
					Exit For
				 Else
					PathTemp=PathTemp & Arr_Path(Tempi) & "/"
				 End If
				 If CheckDir(PathTemp)=False Then
					If MakeNewsDir(PathTemp)=False Then
					   SaveTf=False
					   Exit For
					End If
				 End If
			  Next
		TempUrlArray=Split(FileUrl,"/")  
		Ranfilestr=GetFileID(strtemp,TempUrlArray(Ubound(TempUrlArray)),3)'�����ļ���
		'Call SaveRemoteFile(Ranfilestr,FileUrl)'����Զ���ļ�
		If SaveRemoteFile(Ranfilestr,FileUrl)<>False then'����Զ���ļ�
			Ranfilestr1=Ranfilestr
			if Thumb_WaterMark=1 And Lx=2 then call SKThumb.AddWaterMark(Ranfilestr)'ˮӡ
			Sk_SaveFile = Ranfilestr1
		Else
			Sk_SaveFile = False
		End if
		rs.close
		Set rs=nothing
	End function
	
	Private Function CorrectPattern(ByVal str)
		str = Replace(str, "\", "\\")
		str = Replace(str, "~", "\~")
		str = Replace(str, "!", "\!")
		str = Replace(str, "@", "\@")
		str = Replace(str, "#", "\#")
		str = Replace(str, "%", "\%")
		str = Replace(str, "^", "\^")
		str = Replace(str, "&", "\&")
		str = Replace(str, "*", "\*")
		str = Replace(str, "(", "\(")
		str = Replace(str, ")", "\)")
		str = Replace(str, "-", "\-")
		str = Replace(str, "+", "\+")
		str = Replace(str, "[", "\[")
		str = Replace(str, "]", "\]")
		str = Replace(str, "<", "\<")
		str = Replace(str, ">", "\>")
		str = Replace(str, ".", "\.")
		str = Replace(str, "/", "\/")
		str = Replace(str, "?", "\?")
		str = Replace(str, "=", "\=")
		str = Replace(str, "|", "\|")
		str = Replace(str, "$", "\$")
		CorrectPattern = str
	End Function
	'===============================================
	'��������FormatRemoteUrl
	'��  �ã���ʽ���ɵ�ǰ��վ������URL-����Ե�ַת��Ϊ���Ե�ַ
	'��  ���� url ----Url�ַ���
	'��  ���� CurrentUrl ----��Ȼ��վURL
	'����ֵ����ʽ��ȡ���Url
	'===============================================
	Public Function FormatRemoteUrl(ByVal URL,ByVal CurrentUrl)
		Dim strUrl
		If Len(URL) < 2 Or Len(URL) > 255 Or Len(CurrentUrl) < 2 Then
			FormatRemoteUrl = vbNullString
			Exit Function
		End If
		CurrentUrl = Trim(Replace(Replace(Replace(Replace(Replace(CurrentUrl, "'", vbNullString), """", vbNullString), vbNewLine, vbNullString), "\", "/"), "|", vbNullString))
		URL = Trim(Replace(Replace(Replace(Replace(Replace(URL, "'", vbNullString), """", vbNullString), vbNewLine, vbNullString), "\", "/"), "|", vbNullString))	
		If InStr(9, CurrentUrl, "/") = 0 Then
			strUrl = CurrentUrl
		Else
			strUrl = Left(CurrentUrl, InStr(9, CurrentUrl, "/") - 1)
		End If

		If strUrl = vbNullString Then strUrl = CurrentUrl
		Select Case Left(LCase(URL), 6)
			Case "http:/", "https:", "ftp://", "rtsp:/", "mms://"
				FormatRemoteUrl = URL
				Exit Function
		End Select

		If Left(URL, 1) = "/" Then
			FormatRemoteUrl = strUrl & URL
			Exit Function
		End If
		
		If Left(URL, 3) = "../" Then
			Dim ArrayUrl
			Dim ArrayCurrentUrl
			Dim ArrayTemp()
			Dim strTemp
			Dim i, n
			Dim c, l
			n = 0
			ArrayCurrentUrl = Split(CurrentUrl, "/")
			ArrayUrl = Split(URL, "../")
			c = UBound(ArrayCurrentUrl)
			l = UBound(ArrayUrl) + 1
			
			If c > l + 2 Then
				For i = 0 To c - l
					ReDim Preserve ArrayTemp(n)
					ArrayTemp(n) = ArrayCurrentUrl(i)
					n = n + 1
				Next
				strTemp = Join(ArrayTemp, "/")
			Else
				strTemp = strUrl
			End If
			URL = Replace(URL, "../", vbNullString)
			FormatRemoteUrl = strTemp & "/" & URL
			Exit Function
		End If
		strUrl = Left(CurrentUrl, InStrRev(CurrentUrl, "/"))
		FormatRemoteUrl = strUrl & Replace(URL, "./", vbNullString)
		Exit Function
	End Function	
	'===============================================
	'��������ReplaceTrim
	'��  �ã����˵��ַ������е�tab�ͻس��ͻ���
	'===============================================
	Public Function ReplaceTrim(ByVal strContent)
		Dim re
		Set re = New RegExp
		re.IgnoreCase = True
		re.Global = True
		re.Pattern = "(" & Chr(8) & "|" & Chr(9) & "|" & Chr(10) & "|" & Chr(13) & ")"
		strContent = re.Replace(strContent, vbNullString)
		Set re = Nothing
		ReplaceTrim = strContent
		Exit Function
	End Function
	'===============================================
	'��������ItemReplaceStr
	'��  �ã���Ŀ�����ַ��滻
	'===============================================
	Public Function ItemReplaceStr(ByVal strContent, ByVal ReplaceList)
		If ReplaceList="" then ItemReplaceStr=strContent : Exit Function
		If  Len(ReplaceList) < 3 Or Len(strContent) = 0  Then Exit Function
		Dim i,ReplaceListArray,ReplaceNameArray
	    On Error Resume Next
		ReplaceListArray = Split(ReplaceList, "$$$")
		For i = 0 To UBound(ReplaceListArray)
			If Len(ReplaceListArray(i)) > 2 Then
				ReplaceNameArray = Split(ReplaceListArray(i), "|")
				strContent = Replace(strContent, ReplaceNameArray(0), ReplaceNameArray(1))
			End If
		Next
		ItemReplaceStr = strContent
	End Function
	'===============================================
	'����ֵ�����زɼ��˵�
	'��  �ã���ȡ�ɼ��˵�
	'===============================================
	Function CjMenu()
		Dim RS,TempStr
		Set Rs=ConnItem.execute("select * from SK_cj where Flag=1 order by ID ASC")
		If Not Rs.eof then
			While not Rs.eof
				TempStr=TempStr & "<TR>" & vbcrlf
				TempStr=TempStr &  " <TD height=30 align=""center"" background=""images/left_bg01.gif"" id=""CjMenu""  style=""cursor:hand"" onClick=""javascript:parent.main.location.href='"& Rs("FileName") &"?Colleclx="&Rs("ID")&"';"" onMouseOver=""leftBgOver(this);"" onMouseOut=""leftBgOut(this,'images/left_bg01.gif');"">"& Rs("CjName") &"�ɼ�</TD>" & vbcrlf
				TempStr=TempStr & "</TR>" & vbcrlf
				Rs.Movenext
			Wend
		End if : Rs.close : Set Rs=Nothing
		CjMenu=TempStr
	End Function
	'===============================================
	'��������Show_Top()
	'��  �ã�'ͷ���� 
	'===============================================
	Sub Show_Top()
	    Dim CJFileName : CJFileName = GetItemConfig("FileName",Colleclx)
		Response.Write "<html>" & vbcrlf
		Response.Write "<head>" & vbcrlf
		Response.Write "<title>�ɼ�ϵͳ</title>" & vbcrlf
		Response.Write "<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"">" & vbcrlf
		Response.Write "<link rel=""stylesheet"" type=""text/css"" href=""css/Admin_Style.css""></head>" & vbcrlf
		Response.Write "<script src=""Inc/Common.JS"" language=""javascript""></script>" & vbcrlf
		Response.Write "<body leftmargin=""0"" topmargin=""0"" marginwidth=""0"" marginheight=""0"">" & vbcrlf
		Response.Write "<table width=""100%"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1"" class=""tableBorder"">" & vbcrlf
		Response.Write "  <tr> " & vbcrlf
		Response.Write "    <td height=""22"" colspan=""2"" align=""center"" bgcolor=""#F3F3F3"" class=""topbg""><strong>"&CjName&"�ɼ�����</td>" & vbcrlf
		Response.Write "  </tr>" & vbcrlf
		Response.Write "</table>" & vbcrlf
		Response.Write "<table width=""100%"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1"" class=""tableBorder"">" & vbcrlf
		Response.Write "  <tr class=""tdbg"">" & vbcrlf
		Response.Write "    <td height=""30"" colspan=""3"" bgcolor=""f3f3f3""><b>˵����</b><br>&nbsp;&nbsp;�١���һ��ʹ�ñ����ܣ����޸�<a href='"& CJFileName &"?action=config'><font color=blue>�ɼ���������</font></a>��<br>" & vbcrlf
		Response.Write "    &nbsp;&nbsp;�ڡ��ɼ�ǰ��<font color=blue>�༭</font>�ɼ���Ŀ��<font color=blue>����</font>��Ŀȷ��������ٽ��вɼ���	</td> " & vbcrlf
		Response.Write "  </tr> " & vbcrlf
		Response.Write "  <tr class=""tdbg"">" & vbcrlf
		Response.Write "    <td width=""79"" height=""30"" bgcolor=""f3f3f3""><strong>����������</strong></td>" & vbcrlf
		Response.Write "    <td width=""600"" bgcolor=""f3f3f3""><a href="& CJFileName &">������ҳ</a> | <a href="""& CJFileName &"?action=add&Colleclx="& Colleclx &""">��������Ŀ</a> | <a href='"& CJFileName &"?action=config&ChannelID=0'>�ɼ���������</a> | <a href=""sk_class.asp?Colleclx="& Colleclx &""">�������� </a>  </td>" & vbcrlf
		'Response.Write "     <form name=""form1"" id=""form1""><td width=""200"" height=""30"" bgcolor=""f3f3f3"">������ʾ��<Select ID=""DClassID"" name=""DClassID"" onchange=""MM_jumpMenu('this',this,0)"">"
		'Response.Write("<option value=""SK_GetArticle.asp?DclassID=0"">ȫ������</option>")
		'Call InitClassSelectOption(1,0,ClassID)
		'Response.Write "</Select></td></form>" & vbcrlf
		Response.Write "  </tr>" & vbcrlf
		Response.Write "</table>" & vbcrlf
	End  Sub 
End Class
%>