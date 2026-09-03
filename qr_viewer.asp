<%@ Language="VBScript" CodePage="65001" %>
<%
Response.Buffer = True
Function StrValue(v)
    Dim value
    If IsObject(v) Then value = v.Value Else value = v
    If IsNull(value) Or IsEmpty(value) Then StrValue = "" Else StrValue = CStr(value)
End Function

Function H(v)
    H = Server.HTMLEncode(StrValue(v))
End Function

Function PositiveID(v)
    Dim s, re
    PositiveID = 0
    s = StrValue(v)
    Set re = New RegExp
    re.Pattern = "^[1-9][0-9]{0,9}$"
    If Not re.Test(s) Then Exit Function
    If CDbl(s) > 2147483647 Then Exit Function
    PositiveID = CLng(s)
End Function

Function ISODate(d)
    Dim value
    If IsObject(d) Then value = d.Value Else value = d
    ISODate = ""
    If IsNull(value) Then Exit Function
    If Not IsDate(value) Then Exit Function
    ISODate = Year(value) & "-" & Right("0" & Month(value), 2) & "-" & Right("0" & Day(value), 2)
End Function
%>
<%
Response.CodePage = 65001
Response.CharSet = "utf-8"

Dim conn, rs, sql
Dim lessonId, lessonTitle, lessonDate, lessonDateText
Dim shell, exec, output, lines, line
Dim wifiSSID, ipAddress
Dim i, pos, inWifi
Dim scriptPath, basePath, attendanceUrl, qrUrl

' 이 코드는 그저 갓, goat 최진형씨의 코드를 일부 가져다썻습니다

lessonId = PositiveID(Request.QueryString("id"))

If lessonId = 0 Then
    Response.Write "잘못된 연수 ID 입니다."
    Response.End
End If


' DB 연결
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "DSN=attendanceDB"


' 연수 정보 가져오기
sql = "SELECT * FROM lesson_list WHERE ID = " & CLng(lessonId)
Set rs = conn.Execute(sql)

If rs.EOF Then
    rs.Close
    Set rs = Nothing
    conn.Close
    Set conn = Nothing

    Response.Write "없는 연수임"
    Response.End
End If

lessonTitle = rs("연수제목")
lessonDate = rs("연수일시")

lessonDateText = ISODate(lessonDate)


' 현재 와이파이 이름이랑 IP 가져오기
wifiSSID = "확인할 수 없음"
ipAddress = "확인할 수 없음"

On Error Resume Next

Set shell = Server.CreateObject("WScript.Shell")


' 현재 연결된 와이파이 SSID 찾기
Set exec = shell.Exec("cmd /c netsh wlan show interfaces | findstr /R /C:""SSID""")
output = exec.StdOut.ReadAll()

lines = Split(output, vbCrLf)

For i = 0 To UBound(lines)

    line = Trim(lines(i))

    If InStr(1, line, "SSID", vbTextCompare) = 1 Then

        If InStr(1, line, "BSSID", vbTextCompare) <> 1 Then

            pos = InStr(line, ":")

            If pos > 0 Then
                wifiSSID = Trim(Mid(line, pos + 1))
                Exit For
            End If

        End If

    End If

Next


' 현재 와이파이의 IPv4 찾기
Set exec = shell.Exec("cmd /c ipconfig")
output = exec.StdOut.ReadAll()

lines = Split(output, vbCrLf)
inWifi = False

For i = 0 To UBound(lines)

    line = Trim(lines(i))

    If Right(line, 6) = "Wi-Fi:" Then

        inWifi = True

    ElseIf inWifi And InStr(line, "IPv4") > 0 Then

        pos = InStr(line, ":")

        If pos > 0 Then
            ipAddress = Trim(Mid(line, pos + 1))
            Exit For
        End If

    End If

Next

On Error GoTo 0

Set exec = Nothing
Set shell = Nothing


' 지금 qr_viewer.asp가 있는 경로 따라가서 출석 주소 만들기
Dim host, port, scheme
host = StrValue(Request.ServerVariables("SERVER_NAME"))
port = StrValue(Request.ServerVariables("SERVER_PORT"))
scheme = "http://"
If LCase(Request.ServerVariables("HTTPS")) = "on" Then scheme = "https://"
If host="localhost" Or host="127.0.0.1" Or host="::1" Then
    host = StrValue(Request.ServerVariables("LOCAL_ADDR"))
    If host="127.0.0.1" Or host="::1" Or host="" Then
        host = ""
        If ipAddress<>"확인할 수 없음" Then host=ipAddress
    End If
End If
attendanceUrl = ""
If host<>"" Then
    If InStr(host,":")>0 And Left(host,1)<>"[" Then host="[" & host & "]"
    scriptPath = Request.ServerVariables("SCRIPT_NAME")
    basePath = Left(scriptPath, InStrRev(scriptPath,"/"))
    attendanceUrl = scheme & host
    If (scheme="http://" And port<>"80") Or (scheme="https://" And port<>"443") Then attendanceUrl=attendanceUrl & ":" & port
    attendanceUrl = attendanceUrl & basePath & "userpage.asp?id=" & lessonId
End If

' 출석 주소로 QR 만들기
qrUrl = ""

If attendanceUrl <> "" Then

    qrUrl = "https://api.qrserver.com/v1/create-qr-code/?size=320x320&data=" & _
            Server.URLEncode(attendanceUrl)

End If

%>

<!DOCTYPE html>
<html lang="ko">

<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>연수 QR</title>
<link rel="stylesheet" href="style.css">
</head>

<body class="qr-page">

<h1><%= H(lessonTitle) %></h1>
<p><%= H(lessonDateText) %></p>

<hr>

<h2>Wifi SSID : <%= H(wifiSSID) %></h2>
<p>위 와이파이에 연결한 뒤 QR을 스캔하세요.</p>

<%
If qrUrl <> "" Then
%>

<img src="<%= H(qrUrl) %>" alt="출석 QR 코드" width="320" height="320">

<p>
출석 주소:<br>
<a href="<%= H(attendanceUrl) %>"><%= H(attendanceUrl) %></a>
</p>

<%
Else
%>

<p>서버의 LAN 주소로 이 페이지를 열어 주세요. 예: http://서버IP:포트/qr_viewer.asp?id=연수번호</p>

<%
End If
%>

<p>
<a href="administer.asp">메인으로</a>
</p>

<%
rs.Close
Set rs = Nothing

conn.Close
Set conn = Nothing
%>

</body>
</html>
