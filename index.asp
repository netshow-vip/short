<!--#include file="Comm.asp"-->
<%
	Call ConnDB
	dim saveurl
	Select Case Act
		Case "save"	'保存
			saveurl=Save(URL)
			Response.Write(saveurl)
		Case "show"	'显示
			Call Show(name)
		' Case Else	'其他
			' Response.Redirect DefURL
    End Select
    
	dim s
	s=Request.QueryString("s")
	call show(s)
	Call CloseDB
	Response.End
%>



<%
	'添加
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
				'添加的网址已经存在
				Save=Rs("name")
			Else
							
				code62=get62str(maxID)
						
				Rs.AddNew
				Rs("name") = code62
				Rs("URL") = URLStr
				Rs.Update
				'网址已经成功添加
				Save=Rs("name")
				
			End If
			Rs.Close
		Else
			Call ShowInfo("您添加的网址有误，或者不存在，请自行检查后再添加，谢谢！</strong></p><p align=""center"" style=""font-size:14px; color:red""><strong><a href=""javascript:history.go(-1);"">返回上一页</a>")
		End If
	End function
%>

<%
	function get62str(url_id)
		Dim decNum, hexNum,remainder
		decNum = url_id  ' 输入的十进制数  
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
	'显示指定地址
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
			Call ShowInfo("您访问的网址错误，或者不存在，请自行检查后再访问，谢谢！</strong></p><p align=""center"" style=""font-size:14px; color:red""><strong>5秒钟后系统将自动返回默认页面：<a href=""" & DefURL & """>" & DefURL & "</a></strong><meta http-equiv=""Refresh"" content=""5;URL=" & DefURL & """ />")
		End If
	End Sub
%>