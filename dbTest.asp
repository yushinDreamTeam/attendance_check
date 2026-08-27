<%@ Language="VBScript" CodePage="65001" %>
<% 
Response.CodePage = 65001
Response.CharSet = "utf-8"
On Error Resume Next

Dim conn
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "DSN=attendanceDB"

If Err.Number <> 0 Then
    Response.Write "DB 연결 실패<br>"
    Response.Write "오류 번호: " & Err.Number & "<br>"
    Response.Write "오류 내용: " & Err.Description
Else
    Response.Write "DB 연결 성공"
    conn.Close
End If

Set conn = Nothing
%>