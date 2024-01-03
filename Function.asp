<%
	Sub ConnDB()
		Set Conn = Server.CreateObject("ADODB.Connection")
		Set Rs = Server.CreateObject("ADODB.RecordSet")
		
		Conn.Open "Provider=Microsoft.Jet.OLEDB.4.0; Jet OLEDB:Database Password=''; Data Source=" & Server.MapPath(strDatabasePath)
	End Sub


	Sub CloseDB()
		Conn.Close
		Set Rs = Nothing
		Set Conn = Nothing
	End Sub
%>

<%
	Sub ShowInfo(infoBody)
		Response.Write	"<!DOCTYPE html PUBLIC ""-//W3C//DTD XHTML 1.0 Transitional//EN""" & VbcrLf &_
						"""http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd"">" & VbcrLf &_
						"<html xmlns=""http://www.w3.org/1999/xhtml"">" & VbcrLf &_
						"<head>" & VbcrLf &_
						"<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"" />" & VbcrLf &_
						"<title>永久免费短域名</title>" & VbcrLf &_
						"</head>" & VbcrLf &_
						"<body>" & VbcrLf &_
						"<p>&nbsp;</p>" & VbcrLf &_
						"<p>&nbsp;</p>" & VbcrLf &_
						"<center><p align=""center"" style=""font-size:14px; color:red""><strong>" & InfoBody & "</strong></p>" & VbcrLf &_
						"永久免费短域名" & "    </a></strong><br>" & VbcrLf &_
                                                 "<p>&nbsp;</p>" & VbcrLf &_
						"<p>&nbsp;</p>" & VbcrLf &_
						"<hr />" & VbcrLf &_
						"</body>" & VbcrLf &_
						"</html>"
		Response.End
	End Sub
%>
