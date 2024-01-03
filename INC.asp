<%
	Const sysDir = "URL.asp?Act=Show&name="	'

	Dim ReqStr
		ReqStr = LCase(Trim(Request.ServerVariables("QUERY_STRING")))

	If Len(ReqStr) > 0 Then
		Response.Redirect sysDir & ReqStr
		Response.End	
	End If
%><a target="_blank" href="http://www.daimaw.com"><font color="#FFFFFF"><span style="font-size: 1pt">ДњТыЭј</span></font></a>