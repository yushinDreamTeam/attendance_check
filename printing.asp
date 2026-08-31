<%@ Language=VBScript CodePage="65001"%>
<%
    Response.CodePage = 65001
    Response.CharSet = "utf-8"

    Dim conn, sql, rs, lessonID
    Set conn = Server.CreateObject("ADODB.Connection")
    conn.Open "DSN=attendanceDB"

    lessonId = Request.QueryString("id")
    If Not IsNumeric(lessonId) Then
        Response.Status = "400 Bad Request"
        Response.Write "잘못된 연수 ID입니다."
        Response.End
    End If

    Dim tableName = "signature" & lessonId
    sql = "SELECT * FROM [" & tableName & "] ORDER BY 이름"
    Set rs = conn.Execute(sql)
%>
<!DOCTYPE html>
<html lang="ko">
<head>
    <meta charset="UTF-8">
    <title>출석 명단</title>
</head>
<body>
    <h2>출석 명단</h2>
    <table border="1">
        <thead>
            <tr>
                <th>이름</th>
                <th>서명</th>
            </tr>
        </thead>
        <tbody>
            <% Do While Not rs.EOF %>
            <tr>
                <td><%= rs("이름") %></td>
                <td><%= rs("서명") %></td>
            </tr>
            <% 
            rs.MoveNext
            Loop 
            %>
        </tbody>
    </table>
</body>
</html>