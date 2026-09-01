<%@ Language="VBScript" CodePage="65001" %>
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

lessonId = Request.QueryString("id")

If Not IsNumeric(lessonId) Then
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

lessonDateText = Year(lessonDate) & "-" & _
                 Right("0" & Month(lessonDate), 2) & "-" & _
                 Right("0" & Day(lessonDate), 2)


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
attendanceUrl = ""

If ipAddress <> "확인할 수 없음" Then

    scriptPath = Request.ServerVariables("SCRIPT_NAME")
    basePath = Left(scriptPath, InStrRev(scriptPath, "/"))

    attendanceUrl = "http://" & ipAddress & _
                    basePath & _
                    "userpage.asp?id=" & CLng(lessonId)

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

<h1><%= lessonTitle %></h1>
<p><%= lessonDateText %></p>

<hr>

<h2>Wifi SSID : <%= wifiSSID %></h2>
<p>위 와이파이에 연결한 뒤 QR을 스캔하세요.</p>

<%
If qrUrl <> "" Then
%>

<img src="<%= qrUrl %>" alt="출석 QR 코드" width="320" height="320">

<p>
출석 주소:<br>
<a href="<%= attendanceUrl %>"><%= attendanceUrl %></a>
</p>

<%
Else
%>

<p>IP 주소를 못 찾아서 QR을 만들 수 없음</p>

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
