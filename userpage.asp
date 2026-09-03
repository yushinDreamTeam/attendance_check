<%@ Language="VBScript" CodePage="65001" %>
<%
Response.Buffer = True
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

Sub FailPage(status, message)
    Response.Status = status
    Response.Write "<meta charset='utf-8'><p>" & H(message) & "</p><a href='javascript:history.back()'>돌아가기</a>"
    Response.End
End Sub

Function StrValue(v)
    Dim value
    If IsObject(v) Then value = v.Value Else value = v
    If IsNull(value) Or IsEmpty(value) Then StrValue = "" Else StrValue = CStr(value)
End Function
%>
<%
    Response.CodePage = 65001
    Response.CharSet = "utf-8"
    Dim conn, sql, rs, lessonID
    Set conn = Server.CreateObject("ADODB.Connection")
    conn.Open "DSN=attendanceDB"

    lessonID = PositiveID(Request.QueryString("id"))
    If lessonID = 0 Then
        Response.Status = "400 Bad Request"
        Response.Write "잘못된 연수 ID입니다."
        Response.End
    End If

    sql = "SELECT * FROM lesson_list WHERE ID = " & CLng(lessonID)
    Set rs = conn.Execute(sql)
    If rs.EOF Then FailPage "404 Not Found", "해당 연수가 없습니다."
%>

<!DOCTYPE html>
<html lang="ko">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>출석</title>

    <link rel="stylesheet" href="userpage.css">
</head>

<body>

    <h1><%= H(rs("연수제목")) %></h1>
    <p class="lesson-date"><%= H(ISODate(rs("연수일시"))) %></p>

    <form id="attendance-form" data-lesson-id="<%= lessonID %>">

        <label for="user-name">이름</label>

        <input
            type="text"
            id="user-name"
            maxlength="255"
            required
        >

        <br><br>

        <label>서명</label>

        <br>

        <canvas
            id="signature-canvas"
            width="600"
            height="200">
        </canvas>

        <br>

        <button
            type="button"
            id="clear-signature-button">
            서명 지우기
        </button>

        <p id="submit-status" role="status"></p>
        <button type="submit" id="submit-button">
            제출
        </button>

    </form>

    <script src="./userpage.js"></script>

</body>
</html>
<% rs.Close: conn.Close %>
