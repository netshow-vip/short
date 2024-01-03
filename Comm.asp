<%@ LANGUAGE = VBScript CodePage = 936%>
<%
	Option Explicit
	Response.Buffer = True 
	Response.Expires = -1
	Response.ExpiresAbsolute = Now() - 1 
	Response.Expires = 0 
	Response.CacheControl = "no-cache"
	Response.AddHeader "Access-Control-Allow-Origin", "*"
	' Response.AddHeader "Access-Control-Allow-Headers","Origin"
	' response.addHeader("Access-Control-Allow-Origin","*");
	response.addHeader("Access-Control-Allow-Methods","GET, POST, PUT, DELETE, OPTIONS");
	response.setHeader("Access-Control-Allow-Headers","x-requested-with");
	response.addHeader("Access-Control-Max-Age","1800");//30 min

%>

<!--#include file="Config.asp"-->

<%
	Dim Conn, Rs, SQL

	Dim Act
		Act = LCase(Trim(Request("Act")))

	Dim name
		name = LCase(Trim(Request("name")))

	Dim URL
		URL = LCase(Trim(Request("URL")))
%>

<!--#include file="Function.asp"-->
