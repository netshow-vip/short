<!--#include file="Comm.asp"-->
<%
	Call ConnDB
	dim saveurl
	Select Case Act
		Case "save"	'����
			saveurl=Save(URL)
			Response.Write(saveurl)
		Case "show"	'��ʾ
			Call Show(name)
		' Case Else	'����
			' Response.Redirect DefURL
    End Select
    
	dim s
	s=Request.QueryString("s")
	call show(s)
	Call CloseDB
	Response.End
%>



<%
	'���
	function Save(URLStr)
		dim code62
		Rs.Open "SELECT MAX(ID) as mid FROM daimaw_URL", Conn
        Dim maxID  
        maxID = Rs("mid")+1		
		Rs.Close		
		If Len(URLStr)>0 Then
			SQL = "Select Top 1 * From daimaw_URL Where URL='" & URLStr & "'"
			Rs.Open SQL, Conn, 1, 3
			If Not Rs.Eof Then
				'��ӵ���ַ�Ѿ�����
				Save=Rs("name")
			Else
							
				code62=get62str(maxID)
						
				Rs.AddNew
				Rs("name") = code62
				Rs("URL") = URLStr
				Rs.Update
				'��ַ�Ѿ��ɹ����
				Save=Rs("name")
				
			End If
			Rs.Close
		Else
			Call ShowInfo("����ӵ���ַ���󣬻��߲����ڣ������м�������ӣ�лл��</strong></p><p align=""center"" style=""font-size:14px; color:red""><strong><a href=""javascript:history.go(-1);"">������һҳ</a>")
		End If
	End function
%>

<%
	function get62str(url_id)
		Dim decNum, hexNum,remainder
		decNum = url_id  ' �����ʮ������  
		hexNum = ""  
  
		While decNum > 0  
			remainder = decNum Mod 62  
			If remainder < 10 Then  
				hexNum = CStr(remainder) & hexNum  
			Else  
				hexNum = Chr(remainder - 10 + Asc("A")) & hexNum  
			End If  
			decNum = Int(decNum / 62)  
		Wend  
		get62str=hexNum
	end function
%>

<%
	'��ʾָ����ַ
	Sub Show(URLStr)
			'Response.Write(URLStr)
		If Len(URLStr)>0 Then
			SQL = "Select Top 1 * From daimaw_URL Where name='" & URLStr & "'"
			Rs.Open SQL, Conn, 1, 3
			If Not Rs.Eof Then
			
				Response.Redirect Rs("URL")
			Else
			    
				Response.Redirect DefURL
			End If
			Rs.Close
		Else
			Call ShowInfo("�����ʵ���ַ���󣬻��߲����ڣ������м����ٷ��ʣ�лл��</strong></p><p align=""center"" style=""font-size:14px; color:red""><strong>5���Ӻ�ϵͳ���Զ�����Ĭ��ҳ�棺<a href=""" & DefURL & """>" & DefURL & "</a></strong><meta http-equiv=""Refresh"" content=""5;URL=" & DefURL & """ />")
		End If
	End Sub
%>